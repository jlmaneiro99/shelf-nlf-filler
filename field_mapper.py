"""Deterministic NLF field label → product value mapping."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

ALLERGENS = [
    'gluten', 'cereal', 'egg', 'peanut', 'soya', 'soybean',
    'soy', 'milk', 'dairy', 'nut', 'celery', 'mustard',
    'sesame', 'sulphite', 'sulphur', 'lupin', 'mollusc',
    'shellfish', 'crustacean', 'fish',
]

SKIP_LABEL_PATTERNS = [
    re.compile(r'^product\s+\d+\s*$', re.I),
    re.compile(r'^fill in product', re.I),
    re.compile(r'^section$', re.I),
    re.compile(r'^field$', re.I),
]


def should_skip_label(field_label: str) -> bool:
    label = field_label.strip()
    if not label or len(label) < 3:
        return True
    return any(p.search(label) for p in SKIP_LABEL_PATTERNS)


def decode_allergen_details(product: Dict[str, Any]) -> List[Dict[str, Any]]:
    if product.get('allergen_details'):
        return list(product['allergen_details'])
    details: List[Dict[str, Any]] = []
    raw = product.get('allergens_raw') or product.get('allergens') or []
    if not isinstance(raw, list):
        return details
    for entry in raw:
        s = str(entry)
        if s.endswith(' (possible)'):
            details.append({'allergen': s[:-11], 'may_contain_traces': True})
        else:
            details.append({'allergen': s, 'present_in_formulation': True})
    return details


def normalize_product(product: Dict[str, Any]) -> Dict[str, Any]:
    p = dict(product)
    p['allergen_details'] = decode_allergen_details(p)
    return p


def is_allergen_field(label: str) -> bool:
    return any(a in label for a in ALLERGENS)


def get_allergen_value(label: str, allergen_details: List[Dict[str, Any]]) -> str:
    present = [
        str(a.get('allergen', '')).lower()
        for a in allergen_details
        if a.get('present_in_formulation') or a.get('present')
    ]
    may = [
        str(a.get('allergen', '')).lower()
        for a in allergen_details
        if a.get('may_contain_traces') or a.get('may_contain')
    ]
    for key in ALLERGENS:
        if key in label:
            if any(key in p or p in key for p in present):
                return 'Present'
            if any(key in m or m in key for m in may):
                return 'May Contain'
            return 'Not Present'
    return 'Not Present'


def map_field_to_value(field_label: str, product: Dict[str, Any]) -> Optional[str]:
    label = field_label.lower().strip().rstrip('*').strip()
    p = normalize_product(product)

    if is_allergen_field(label):
        return get_allergen_value(label, p.get('allergen_details', []))

    if any(x in label for x in ['full product name', 'product name']) and 'description' not in label:
        return str(p.get('product_name', ''))
    if label in ['brand name', 'brand']:
        return str(p.get('brand_name', ''))
    if any(x in label for x in ['supplier name', 'supplier / company', 'company name']):
        return str(p.get('supplier_name', p.get('brand_name', '')))
    if any(x in label for x in ['sku / supplier', 'supplier code', 'supplier reference', 'reference code']):
        return str(p.get('sku_code', ''))
    if ('ean' in label or 'barcode' in label or 'gtin' in label) and ('case' in label or 'outer' in label):
        return str(p.get('case_barcode', '')) or 'N/A'
    if ('ean' in label or 'barcode' in label or 'gtin' in label) and 'case' not in label:
        return str(p.get('ean_barcode', ''))
    if any(x in label for x in ['variant', 'pack size']) and 'case' not in label:
        return str(p.get('variant', ''))
    if 'product description' in label and 'usp' not in label:
        return str(p.get('product_description', ''))
    if any(x in label for x in ['usp', 'key claims', 'unique selling', 'usp / website']):
        return str(p.get('usp', p.get('product_description', '')))

    if any(x in label for x in ['rrp', 'retail price', 'recommended retail', 'msrp']):
        return str(p.get('rrp', ''))
    if any(x in label for x in ['wholesale price', 'trade price', 'cost to', 'normal trade']):
        return str(p.get('trade_price_per_case', ''))
    if 'cost price' in label and 'wholesale' not in label:
        return str(p.get('cost_price_per_case', '')) or 'N/A'
    if 'units per case' in label or 'units/case' in label:
        return str(p.get('units_per_case', ''))
    if 'case size description' in label or 'case configuration' in label or ('case size' in label and 'eg' in label):
        return str(p.get('case_size_description', ''))
    if 'vat rate' in label or label == 'vat':
        return str(p.get('vat_rate', p.get('tax_vat_rate', '')))
    if 'minimum order' in label or 'moq' in label:
        return str(p.get('moq_units', '')) or 'N/A'
    if 'lead time' in label:
        return str(p.get('lead_time_days', '')) or 'N/A'
    if 'payment terms' in label:
        return str(p.get('payment_terms', '')) or 'N/A'

    if 'gross weight' in label:
        return str(p.get('case_gross_weight_kg', '')) or 'N/A'
    if 'net weight' in label:
        return str(p.get('case_net_weight_kg', p.get('unit_net_weight_g', ''))) or 'N/A'
    if 'minimum shelf life' in label or ('shelf life' in label and ('delivery' in label or 'receipt' in label)):
        return str(p.get('min_shelf_life_on_delivery_weeks', '')) or 'N/A'
    if 'shelf life' in label:
        v = p.get('shelf_life_weeks')
        if v:
            if 'day' in label:
                return str(int(v) * 7)
            return str(v)
        return 'N/A'
    if 'storage conditions' in label:
        return str(p.get('storage_conditions', ''))
    if 'storage instructions' in label:
        return str(p.get('storage_instructions', p.get('storage_conditions', ''))) or 'N/A'
    if 'cases per pallet' in label or ('pallet' in label and 'case' in label):
        return str(p.get('cases_per_pallet', '')) or 'N/A'

    if 'country of provenance' in label or 'provenance' in label:
        return str(p.get('country_of_provenance') or p.get('country_of_origin', ''))
    if 'country of origin' in label or 'country of manufacture' in label:
        return str(p.get('country_of_origin', ''))
    if ('hs' in label or 'commodity code' in label or 'tariff' in label) and 'code' in label:
        return str(p.get('hs_commodity_code', '')) or 'N/A'
    if 'meursing' in label:
        return str(p.get('meursing_code', '')) or 'N/A'
    if 'eu address' in label:
        return 'Yes' if p.get('eu_address_on_pack') else 'No'
    if 'uk address' in label:
        return 'Yes' if p.get('uk_address_on_pack') else 'No'

    if 'vegan' in label:
        return 'Yes' if p.get('is_vegan') else 'No'
    if 'vegetarian' in label:
        return 'Yes' if p.get('is_vegetarian') else 'No'
    if 'gluten free' in label or ('gluten' in label and 'free' in label):
        return 'Yes' if p.get('is_gluten_free') else 'No'
    if 'organic certification' in label or 'organic cert' in label or 'certification number' in label:
        return str(p.get('organic_cert_number', '')) or 'N/A'
    if 'organic' in label and 'cert' not in label:
        return 'Yes' if p.get('is_organic') else 'No'
    if 'fairtrade' in label or 'fair trade' in label:
        return 'Yes' if p.get('is_fairtrade') else 'No'
    if 'added sugar' in label:
        return 'Yes' if p.get('contains_added_sugar') else 'No'
    if 'gm free' in label or 'non-gmo' in label or 'gmo free' in label:
        return 'Yes' if p.get('is_gm_free') else 'No'
    if 'hfss scope' in label:
        return 'Yes' if p.get('hfss_scope') else 'No'
    if 'hfss score' in label or 'nutrient profile score' in label or 'npm score' in label:
        return str(p.get('hfss_score', '')) or 'N/A'
    if 'biodynamic' in label:
        return 'Yes' if p.get('is_biodynamic') else 'No'
    if 'irradiated' in label:
        return 'Yes' if p.get('is_irradiated') else 'No'
    if 'contains alcohol' in label or 'alcohol?' in label:
        return 'Yes' if p.get('contains_alcohol') else 'No'
    if 'abv' in label or 'alcohol %' in label:
        return str(p.get('abv_percentage', '')) or 'N/A'
    if 'palm oil free' in label:
        status = str(p.get('palm_oil_status', ''))
        return 'Yes' if 'not contain' in status.lower() else 'No'
    if 'palm oil' in label:
        return str(p.get('palm_oil_status', '')) or 'N/A'
    if 'egg' in label and 'allergen' not in label:
        return str(p.get('egg_status', '')) or 'N/A'

    serving = p.get('serving_size_value') or 100

    def nutritional(val, label_str):
        if val is None:
            return 'N/A'
        if 'per serving' in label_str:
            return str(round(float(val) * float(serving) / 100, 1))
        return str(val)

    if 'energy' in label and 'kcal' in label:
        return nutritional(p.get('energy_kcal'), label)
    if 'energy' in label and ('kj' in label or 'kilojoule' in label):
        return nutritional(p.get('energy_kj'), label)
    if 'saturate' in label:
        return nutritional(p.get('saturates'), label)
    if 'monounsaturate' in label:
        return nutritional(p.get('monounsaturates'), label)
    if 'polyunsaturate' in label:
        return nutritional(p.get('polyunsaturates'), label)
    if 'fat' in label and 'trans' not in label:
        return nutritional(p.get('fat'), label)
    if 'carbohydrate' in label or 'carbs' in label:
        return nutritional(p.get('carbohydrates'), label)
    if 'sugar' in label:
        return nutritional(p.get('sugars'), label)
    if 'fibre' in label or 'fiber' in label:
        return nutritional(p.get('fibre'), label)
    if 'protein' in label:
        return nutritional(p.get('protein'), label)
    if 'sodium' in label:
        salt = p.get('salt')
        return str(round(float(salt) / 2.5, 3)) if salt is not None else 'N/A'
    if 'salt' in label:
        return nutritional(p.get('salt'), label)

    if 'inner packaging material' in label:
        return str(p.get('inner_packaging_material', '')) or 'N/A'
    if 'recyclable' in label:
        return 'Yes' if p.get('is_recyclable') else 'No'
    if 'outer packaging' in label and 'plastic' in label:
        return 'Yes' if p.get('outer_packaging_has_plastic') else 'No'
    if 'biodegradable' in label:
        return 'Yes' if p.get('is_biodegradable') else 'No'
    if 'compostable' in label:
        return 'Yes' if p.get('is_compostable') else 'No'

    if 'ingredients' in label:
        return str(p.get('ingredients', '')) or 'N/A'
    if 'usp / website' in label or 'website description' in label:
        return str(p.get('usp', '')) or 'N/A'
    if any(x in label for x in ['how will you promote', 'promotional plan', 'promotional support']):
        return str(p.get('promotion_plan', '')) or 'N/A'

    return None
