from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Any, Dict
import base64
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
from copy import copy, deepcopy
import datetime
import os

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class ProductData(BaseModel):
    data: Dict[str, Any]


class FillRequest(BaseModel):
    file_base64: str
    products: List[Dict[str, Any]]
    retailer_name: str
    fill_mode: str = 'tabs'


def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if (rng.min_row <= row <= rng.max_row and
                    rng.min_col <= col <= rng.max_col):
                ws.cell(row=rng.min_row, column=rng.min_col).value = value
                return
        return
    cell.value = value


ALLERGEN_KEYS = [
    ('gluten', ['gluten', 'cereal', 'wheat', 'rye', 'barley', 'oat']),
    ('eggs', ['egg']),
    ('fish', ['fish']),
    ('crustaceans', ['crustacean', 'shellfish', 'prawn', 'shrimp', 'crab', 'lobster']),
    ('molluscs', ['mollusc', 'mussel', 'oyster', 'clam', 'squid']),
    ('peanuts', ['peanut', 'groundnut']),
    ('soybeans', ['soy', 'soya', 'soybean']),
    ('milk', ['milk', 'dairy', 'lactose', 'whey', 'casein']),
    ('nuts', ['nut', 'almond', 'hazelnut', 'walnut', 'cashew', 'pecan', 'pistachio', 'macadamia', 'brazil']),
    ('celery', ['celery', 'celeriac']),
    ('mustard', ['mustard']),
    ('sesame', ['sesame']),
    ('sulphites', ['sulphite', 'sulphur', 'sulfite', 'sulfur', 'so2']),
    ('lupin', ['lupin', 'lupine']),
    ('royal jelly', ['royal jelly']),
    ('propolis', ['propolis']),
    ('bee pollen', ['bee pollen']),
]


def get_allergen_value(label, allergen_details):
    if not allergen_details:
        return 'Not Present'
    label_lower = label.lower()
    for category, keywords in ALLERGEN_KEYS:
        if any(kw in label_lower for kw in keywords):
            for a in allergen_details:
                allergen_name = str(a.get('allergen', '')).lower()
                if any(kw in allergen_name for kw in keywords):
                    if a.get('present_in_formulation') or a.get('present'):
                        return 'Present'
                    if a.get('may_contain_traces') or a.get('may_contain'):
                        return 'May Contain'
            return 'Not Present'
    return 'Not Present'


def is_allergen_field(label):
    label_lower = label.lower()
    for _, keywords in ALLERGEN_KEYS:
        if any(kw in label_lower for kw in keywords):
            return True
    return False


def nutritional_value(val, label, serving_size):
    if val is None:
        return None
    label_lower = label.lower()
    v = float(val)
    if 'per serving' in label_lower and serving_size:
        v = round(v * float(serving_size) / 100, 1)
    elif 'per 50g' in label_lower:
        v = round(v * 0.5, 1)
    elif 'per 30g' in label_lower:
        v = round(v * 0.3, 1)
    elif 'per 25g' in label_lower:
        v = round(v * 0.25, 1)
    else:
        v = round(v, 2)
        if v == int(v):
            v = int(v)
    return str(v)


def map_field(label_raw, product):
    label = label_raw.lower().strip().rstrip('*').strip()
    p = product
    serving = p.get('serving_size_value')

    if is_allergen_field(label):
        return get_allergen_value(label, p.get('allergen_details', []))

    if any(x in label for x in ['full product name', 'product name on pack',
                                  'product name', 'name of product',
                                  'article name', 'item name',
                                  'product title']) and 'description' not in label:
        return str(p.get('product_name', ''))

    if label in ['brand name', 'brand', 'brand / manufacturer',
                 'manufacturer', 'manufacturer name', 'brand name *']:
        return str(p.get('brand_name', ''))

    if any(x in label for x in ['supplier name', 'supplier / company',
                                  'company name', 'vendor name',
                                  'supplier company']):
        return str(p.get('supplier_name') or p.get('brand_name', ''))

    if any(x in label for x in ['sku', 'supplier code', 'supplier reference',
                                  'reference code', 'supplier ref',
                                  'your code', 'vendor code',
                                  'article number', 'item code',
                                  'product code', 'supplier item code',
                                  'commodity code']) and 'hs' not in label and 'tariff' not in label:
        return str(p.get('sku_code', ''))

    if any(x in label for x in ['case barcode', 'outer barcode',
                                  'case ean', 'shipper barcode',
                                  'carton barcode', 'case gtin',
                                  'case upc', 'outer ean']):
        return str(p.get('case_barcode', '')) or 'N/A'

    if any(x in label for x in ['ean', 'barcode', 'gtin', 'upc',
                                  'unit barcode', 'individual barcode',
                                  'product barcode', 'item barcode',
                                  'unit ean', 'unit gtin']):
        return str(p.get('ean_barcode', ''))

    if any(x in label for x in ['variant', 'pack size', 'product size',
                                  'size / format', 'format', 'pack format']):
        return str(p.get('variant', ''))

    if 'product description' in label and 'usp' not in label and 'sell' not in label:
        return str(p.get('product_description', ''))

    if any(x in label for x in ['usp', 'key claims', 'unique selling',
                                  'selling point', 'key benefit',
                                  'product claim', 'about the product',
                                  'sell copy', 'marketing description',
                                  'consumer description', 'website description']):
        return str(p.get('usp', ''))

    if 'ingredient' in label:
        return str(p.get('ingredients', '')) or 'N/A'

    if any(x in label for x in ['how will you promote', 'promotional plan',
                                  'promotional support', 'promotional activity',
                                  'marketing support', 'trade support',
                                  'marketing plan', 'consumer marketing',
                                  'promotion plan', 'how do you plan to market']):
        return str(p.get('promotion_plan', '')) or 'N/A'

    if any(x in label for x in ['rrp', 'retail price', 'recommended retail',
                                  'msrp', 'normal rrp', 'consumer price',
                                  'selling price', 'shelf price']):
        return str(p.get('rrp', ''))

    if any(x in label for x in ['wholesale price', 'trade price',
                                  'cost to', 'normal trade price',
                                  'invoice price', 'supply price',
                                  'ex-works price', 'unit cost to retailer',
                                  'buying price', 'net price']):
        return str(p.get('trade_price_per_case', ''))

    if any(x in label for x in ['cost price', 'landed cost',
                                  'our cost']) and 'wholesale' not in label:
        return str(p.get('cost_price_per_case', '')) or 'N/A'

    if any(x in label for x in ['units per case', 'units/case',
                                  'case quantity', 'qty per case',
                                  'pieces per case', 'count per case',
                                  'units per outer']):
        return str(p.get('units_per_case', ''))

    if any(x in label for x in ['case size description', 'case configuration',
                                  'pack configuration', 'pack description',
                                  'case contents', 'case format']) or \
       ('case size' in label and any(x in label for x in ['eg', 'e.g', 'example', 'x'])):
        return str(p.get('case_size_description', ''))

    if any(x in label for x in ['vat rate', 'vat', 'tax rate', 'gst rate',
                                  'hst', 'sales tax', 'tax / vat',
                                  'tax type']):
        return str(p.get('vat_rate', ''))

    if any(x in label for x in ['minimum order', 'moq', 'min order',
                                  'minimum quantity']):
        v = p.get('moq_units') or p.get('moq_value')
        return str(v) if v else 'N/A'

    if any(x in label for x in ['lead time', 'delivery time',
                                  'lead time (days)', 'lead time (weeks)',
                                  'turnaround time']):
        return str(p.get('lead_time_days', '')) or 'N/A'

    if any(x in label for x in ['payment terms', 'terms of payment',
                                  'credit terms', 'payment conditions']):
        return str(p.get('payment_terms', '')) or 'N/A'

    if any(x in label for x in ['case gross weight', 'gross weight',
                                  'gross case weight', 'total case weight',
                                  'case weight (gross)', 'weight incl']):
        return str(p.get('case_gross_weight_kg', '')) or 'N/A'

    if any(x in label for x in ['case net weight', 'net weight',
                                  'net case weight', 'product weight only',
                                  'case weight (net)', 'weight excl']):
        return str(p.get('case_net_weight_kg', '')) or 'N/A'

    if any(x in label for x in ['unit weight', 'individual unit weight',
                                  'net weight per unit', 'unit net weight',
                                  'product weight per unit']):
        w = p.get('unit_net_weight_g')
        return str(w) if w else 'N/A'

    if any(x in label for x in ['minimum shelf life', 'min shelf life',
                                  'shelf life on delivery', 'shelf life at receipt',
                                  'minimum remaining shelf life',
                                  'guaranteed shelf life']):
        v = p.get('min_shelf_life_on_delivery_weeks')
        if v:
            if 'day' in label:
                return str(int(v) * 7)
            return str(v)
        return 'N/A'

    if any(x in label for x in ['shelf life', 'best before', 'product life',
                                  'life of product', 'expiry', 'use by',
                                  'total shelf life']):
        v = p.get('shelf_life_weeks')
        if v:
            if 'day' in label:
                return str(int(v) * 7)
            return str(v)
        return 'N/A'

    if any(x in label for x in ['storage conditions', 'storage criteria',
                                  'storage requirement', 'store']):
        return str(p.get('storage_conditions', ''))

    if any(x in label for x in ['storage instructions', 'storage advice',
                                  'how to store', 'storage guidance']):
        return str(p.get('storage_instructions', '')) or 'N/A'

    if any(x in label for x in ['cases per pallet', 'pallet configuration',
                                  'cases/pallet', 'units per pallet']):
        return str(p.get('cases_per_pallet', '')) or 'N/A'

    if 'cases per layer' in label or 'layers per pallet' in label:
        v = p.get('cases_per_layer') or p.get('layers_per_pallet')
        return str(v) if v else 'N/A'

    if any(x in label for x in ['country of provenance', 'provenance',
                                  'last country of duty', 'duty paid country']):
        return str(p.get('country_of_provenance') or p.get('country_of_origin', ''))

    if any(x in label for x in ['country of origin', 'country of manufacture',
                                  'country of production', 'made in',
                                  'produced in', 'origin country',
                                  'place of manufacture']):
        return str(p.get('country_of_origin', ''))

    if any(x in label for x in ['hs code', 'hs / commodity', 'tariff code',
                                  'commodity code', 'customs code',
                                  'hts code', 'import code']) and 'sku' not in label:
        return str(p.get('hs_commodity_code', '')) or 'N/A'

    if 'meursing' in label:
        return str(p.get('meursing_code', '')) or 'N/A'

    if any(x in label for x in ['eu address', 'european address']):
        return 'Yes' if p.get('eu_address_on_pack') else 'No'

    if any(x in label for x in ['uk address', 'united kingdom address',
                                  'gb address']):
        return 'Yes' if p.get('uk_address_on_pack') else 'No'

    if 'vegan' in label and 'non' not in label:
        return 'Yes' if p.get('is_vegan') else 'No'

    if 'vegetarian' in label and 'non' not in label:
        return 'Yes' if p.get('is_vegetarian') else 'No'

    if any(x in label for x in ['gluten free', 'gluten-free', 'gf ',
                                  ' gf', 'free from gluten']):
        return 'Yes' if p.get('is_gluten_free') else 'No'

    if any(x in label for x in ['organic certification number',
                                  'organic cert number', 'cert number',
                                  'certification number', 'accreditation number',
                                  'approval number', 'organic body']):
        return str(p.get('organic_cert_number', '')) or 'N/A'

    if 'organic' in label and 'cert' not in label and 'number' not in label:
        return 'Yes' if p.get('is_organic') else 'No'

    if any(x in label for x in ['fairtrade', 'fair trade', 'fair-trade']):
        return 'Yes' if p.get('is_fairtrade') else 'No'

    if any(x in label for x in ['added sugar', 'contains sugar',
                                  'no added sugar', 'free from sugar']):
        return 'Yes' if p.get('contains_added_sugar') else 'No'

    if any(x in label for x in ['gm free', 'non-gmo', 'gmo free',
                                  'non gmo', 'genetically modified free',
                                  'not genetically modified']):
        return 'Yes' if p.get('is_gm_free') else 'No'

    if any(x in label for x in ['hfss scope', 'is product hfss',
                                  'hfss in scope']):
        return 'Yes' if p.get('hfss_scope') else 'No'

    if any(x in label for x in ['hfss score', 'nutrient profile score',
                                  'npm score', 'hfss nutrient']):
        return str(p.get('hfss_score', '')) or 'N/A'

    if any(x in label for x in ['less healthy', 'hfss less healthy',
                                  'is product less healthy']):
        return 'Yes' if p.get('hfss_less_healthy') else 'No'

    if 'biodynamic' in label:
        return 'Yes' if p.get('is_biodynamic') else 'No'

    if 'irradiated' in label:
        return 'Yes' if p.get('is_irradiated') else 'No'

    if any(x in label for x in ['contains alcohol', 'alcohol?',
                                  'is this alcoholic', 'alcoholic product']):
        return 'Yes' if p.get('contains_alcohol') else 'No'

    if any(x in label for x in ['abv', 'alcohol by volume',
                                  'alcohol %', 'alcohol content',
                                  'alcohol percentage']):
        return str(p.get('abv_percentage', '')) or 'N/A'

    if 'palm oil free' in label:
        status = str(p.get('palm_oil_status', '')).lower()
        return 'Yes' if 'not contain' in status or 'free' in status else 'No'

    if any(x in label for x in ['palm oil type', 'type of palm oil']):
        return str(p.get('palm_oil_type', '')) or 'N/A'

    if any(x in label for x in ['palm oil percentage', '% palm oil']):
        return str(p.get('palm_oil_percentage', '')) or 'N/A'

    if any(x in label for x in ['palm oil', 'rspo']):
        return str(p.get('palm_oil_status', '')) or 'N/A'

    if any(x in label for x in ['egg status', 'egg sourcing',
                                  'free range egg', 'are they free range']):
        return str(p.get('egg_status', '')) or 'N/A'

    if any(x in label for x in ['dairy free', 'dairy-free',
                                  'free from dairy', 'lactose free']):
        return 'Yes' if p.get('is_dairy_free') else 'No'

    if any(x in label for x in ['soy free', 'soya free',
                                  'free from soy', 'free from soya']):
        v = p.get('allergen_details', [])
        soy_present = any('soy' in str(a.get('allergen', '')).lower()
                         and (a.get('present') or a.get('present_in_formulation'))
                         for a in v)
        return 'No' if soy_present else 'Yes'

    if any(x in label for x in ['nut free', 'free from nuts',
                                  'tree nut free']):
        v = p.get('allergen_details', [])
        nut_present = any('nut' in str(a.get('allergen', '')).lower()
                         and (a.get('present') or a.get('present_in_formulation'))
                         for a in v)
        return 'No' if nut_present else 'Yes'

    if any(x in label for x in ['plant based', 'plant-based',
                                  'suitable for plant based']):
        return 'Yes' if p.get('is_vegan') else 'No'

    if 'halal' in label:
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('halal' in c for c in certs) else 'No'

    if 'kosher' in label:
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('kosher' in c for c in certs) else 'No'

    if any(x in label for x in ['b corp', 'bcorp', 'b-corp']):
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('b corp' in c for c in certs) else 'No'

    if any(x in label for x in ['rainforest alliance',
                                  'rainforest']):
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('rainforest' in c for c in certs) else 'No'

    if any(x in label for x in ['usda organic', 'usda certified']):
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('usda' in c for c in certs) else 'No'

    if any(x in label for x in ['non-gmo project', 'non gmo project verified']):
        certs = [str(c).lower() for c in p.get('certifications', [])]
        return 'Yes' if any('non-gmo' in c or 'non gmo' in c for c in certs) else 'No'

    if 'energy' in label and any(x in label for x in ['kcal', 'cal', 'calorie']):
        return nutritional_value(p.get('energy_kcal'), label, serving)

    if 'energy' in label and any(x in label for x in ['kj', 'kilojoule', 'kilo joule']):
        return nutritional_value(p.get('energy_kj'), label, serving)

    if any(x in label for x in ['energy value', 'energy content',
                                  'calorific value', 'energy/']):
        kcal = p.get('energy_kcal')
        kj = p.get('energy_kj')
        if kcal and kj:
            return f"{kj}kJ / {kcal}kcal"
        return nutritional_value(kcal, label, serving)

    if any(x in label for x in ['saturate', 'saturated fat',
                                  'of which saturate', 'sat fat']):
        return nutritional_value(p.get('saturates'), label, serving)

    if 'monounsaturate' in label:
        return nutritional_value(p.get('monounsaturates'), label, serving)

    if 'polyunsaturate' in label:
        return nutritional_value(p.get('polyunsaturates'), label, serving)

    if 'trans fat' in label or 'trans fatty' in label:
        return '0'

    if any(x in label for x in ['total fat', 'fat content',
                                  'fat g', 'fat per', '- fat',
                                  'of which fat']) or \
       label in ['fat', 'fat *', 'fat (g)', 'fat g per 100g *',
                 'fat g per 100g', 'fat (g per 100g)']:
        return nutritional_value(p.get('fat'), label, serving)

    if any(x in label for x in ['total carbohydrate', 'carbohydrate',
                                  'carbs', 'carbohydrates',
                                  'total carbs', 'of which carbs']):
        return nutritional_value(p.get('carbohydrates'), label, serving)

    if any(x in label for x in ['total sugar', 'of which sugar',
                                  'sugars', 'sugar content']):
        return nutritional_value(p.get('sugars'), label, serving)

    if any(x in label for x in ['dietary fibre', 'dietary fiber',
                                  'fibre', 'fiber', 'roughage']):
        return nutritional_value(p.get('fibre'), label, serving)

    if any(x in label for x in ['polyol', 'sugar alcohol']):
        return nutritional_value(p.get('polyols'), label, serving)

    if any(x in label for x in ['starch']):
        return nutritional_value(p.get('starch'), label, serving)

    if any(x in label for x in ['total protein', 'protein content',
                                  'protein per', 'of which protein']) or \
       label in ['protein', 'protein *', 'protein g per 100g *',
                 'protein (g per 100g)', 'protein g per 100g']:
        return nutritional_value(p.get('protein'), label, serving)

    if 'sodium' in label:
        salt = p.get('salt')
        if salt is not None:
            sodium = round(float(salt) / 2.5, 3)
            return str(sodium)
        return 'N/A'

    if any(x in label for x in ['salt equivalent', 'total salt',
                                  'salt content', 'salt per']) or \
       label in ['salt', 'salt *', 'salt g per 100g *',
                 'salt (g per 100g)', 'salt g per 100g']:
        return nutritional_value(p.get('salt'), label, serving)

    if any(x in label for x in ['inner packaging material',
                                  'packaging material', 'packaging type',
                                  'primary packaging']):
        return str(p.get('inner_packaging_material', '')) or 'N/A'

    if any(x in label for x in ['is packaging recyclable', 'recyclable',
                                  'can it be recycled']):
        return 'Yes' if p.get('is_recyclable') else 'No'

    if any(x in label for x in ['biodegradable']):
        return 'Yes' if p.get('is_biodegradable') else 'No'

    if any(x in label for x in ['compostable']):
        return 'Yes' if p.get('is_compostable') else 'No'

    if 'outer packaging' in label and 'plastic' in label:
        return 'Yes' if p.get('outer_packaging_has_plastic') else 'No'

    if 'inner packaging' in label and 'plastic' in label:
        return 'Yes' if p.get('inner_packaging_has_plastic') else 'No'

    if any(x in label for x in ['recycled plastic', '30% recycled',
                                  'recycled content']):
        return 'Yes' if (p.get('inner_plastic_recycled_30pct') or
                         p.get('outer_plastic_recycled_30pct')) else 'No'

    if any(x in label for x in ['fsc', 'pefc', 'paper certified',
                                  'card certified', 'certified paper']):
        return str(p.get('paper_card_certified', '')) or 'Uncertified'

    return None


def detect_layout(ws):
    label_col = 2
    value_col = 3
    header_row = None
    product_col_start = None

    for row in ws.iter_rows(min_row=1, max_row=10):
        for cell in row:
            if cell.value and not isinstance(cell, MergedCell):
                val = str(cell.value).upper()
                if 'PRODUCT 1' in val or 'PRODUCT1' in val:
                    header_row = cell.row
                    product_col_start = cell.column
                    label_col = cell.column - 1
                    value_col = cell.column
                    break
        if header_row:
            break

    return {
        'label_col': label_col,
        'value_col': value_col,
        'header_row': header_row,
        'product_col_start': product_col_start,
        'is_column_format': header_row is not None,
    }


def fill_single_sheet(ws, product, value_col=3, label_col=2):
    filled = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.column != label_col:
                continue
            if not cell.value:
                continue
            field_label = str(cell.value).strip()
            if len(field_label) < 3:
                continue
            value = map_field(field_label, product)
            if value is not None:
                safe_write(ws, cell.row, value_col, value)
                filled += 1
    return filled


def add_product_column_headers(ws, layout, product_index, product_col):
    header_row = layout['header_row']
    template_col = layout['product_col_start']
    product_num = product_index + 1

    if not header_row:
        return

    for r_offset in range(2):
        r = header_row + r_offset
        template_cell = ws.cell(row=r, column=template_col)
        new_cell = ws.cell(row=r, column=product_col)

        if r_offset == 0:
            new_cell.value = f"PRODUCT {product_num}"
        else:
            new_cell.value = f"Fill in product {product_num} details below ↓"

        if not isinstance(template_cell, MergedCell):
            if template_cell.fill and template_cell.fill.fill_type == 'solid':
                new_cell.fill = copy(template_cell.fill)
            if template_cell.font:
                new_cell.font = copy(template_cell.font)
            if template_cell.alignment:
                new_cell.alignment = copy(template_cell.alignment)


def copy_data_cell_format(ws, layout, template_col, new_col, data_start_row):
    for row in ws.iter_rows(min_row=data_start_row):
        tc = ws.cell(row=row[0].row, column=template_col)
        nc = ws.cell(row=row[0].row, column=new_col)
        if isinstance(tc, MergedCell) or isinstance(nc, MergedCell):
            continue
        if tc.fill and tc.fill.fill_type == 'solid':
            nc.fill = copy(tc.fill)


@app.post("/fill")
async def fill_nlf(req: FillRequest):
    try:
        if not req.products:
            raise HTTPException(status_code=422, detail="No products provided")

        file_bytes = base64.b64decode(req.file_base64)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

        product_sheet = None
        for name in wb.sheetnames:
            nl = name.lower()
            if (nl.startswith('product') or nl.startswith('sku') or
                    nl == 'template' or nl == 'new line' or
                    nl == 'new line form' or nl == 'nlf' or nl == 'form' or
                    'line form' in nl or 'new line' in nl or 'submission' in nl):
                product_sheet = name
                break

        if not product_sheet:
            best, best_count = None, 0
            for name in wb.sheetnames:
                count = sum(1 for row in wb[name].iter_rows()
                            for cell in row
                            if cell.value and not isinstance(cell, MergedCell))
                if count > best_count:
                    best_count, best = count, name
            product_sheet = best or wb.sheetnames[0]

        ws = wb[product_sheet]
        layout = detect_layout(ws)

        total_filled = 0

        if layout['is_column_format']:
            label_col = layout['label_col']
            template_col = layout['product_col_start']
            data_start_row = (layout['header_row'] or 1) + 2

            for i, product in enumerate(req.products):
                col = template_col + i
                add_product_column_headers(ws, layout, i, col)
                if i > 0:
                    copy_data_cell_format(ws, layout, template_col, col, data_start_row)
                total_filled += fill_single_sheet(ws, product, value_col=col, label_col=label_col)

        elif req.fill_mode == 'tabs' or len(req.products) == 1:
            template_sheet_name = product_sheet
            template_idx = wb.sheetnames.index(template_sheet_name)
            for i, product in enumerate(req.products):
                if i == 0:
                    target_ws = ws
                else:
                    target_ws = wb.copy_worksheet(wb[template_sheet_name])
                    target_ws.title = f"Product {i + 1}"
                    for dv in wb[template_sheet_name].data_validations.dataValidation:
                        target_ws.add_data_validation(deepcopy(dv))
                    current_idx = wb.sheetnames.index(target_ws.title)
                    wb.move_sheet(target_ws.title, offset=(template_idx + i) - current_idx)
                    for row in target_ws.iter_rows():
                        for cell in row:
                            if not isinstance(cell, MergedCell) and cell.column == 3:
                                cell.value = None
                total_filled += fill_single_sheet(target_ws, product)

        elif req.fill_mode == 'rows':
            header_row_num = None
            for row in ws.iter_rows(min_row=1, max_row=5):
                non_empty = [c for c in row if c.value and not isinstance(c, MergedCell)]
                if len(non_empty) > 3:
                    header_row_num = row[0].row
                    break
            if header_row_num:
                headers = {}
                for cell in ws[header_row_num]:
                    if cell.value and not isinstance(cell, MergedCell):
                        headers[cell.column] = str(cell.value).strip()
                for i, product in enumerate(req.products):
                    row_num = header_row_num + 1 + i
                    for col, field_label in headers.items():
                        value = map_field(field_label, product)
                        if value is not None:
                            safe_write(ws, row_num, col, value)
                            total_filled += 1

        elif req.fill_mode == 'columns':
            label_col = layout['label_col']
            base_col = layout['value_col']
            for i, product in enumerate(req.products):
                total_filled += fill_single_sheet(
                    ws, product, value_col=base_col + i, label_col=label_col,
                )

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        date_str = datetime.date.today().isoformat()
        name = req.products[0].get('product_name', 'Product') if req.products else 'Product'
        if len(req.products) > 1:
            filename = f"{req.retailer_name}_NLF_{len(req.products)}_products_{date_str}.xlsx".replace(" ", "_")
        else:
            filename = f"{req.retailer_name}_{name}_{date_str}.xlsx".replace(" ", "_")

        return {
            "file_base64": base64.b64encode(output.read()).decode(),
            "filename": filename,
            "fields_filled": total_filled,
            "products_filled": len(req.products),
            "fill_mode": req.fill_mode,
        }

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"{str(e)}\n{traceback.format_exc()}")


@app.get("/health")
async def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
