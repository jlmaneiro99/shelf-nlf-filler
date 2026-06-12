import json
from main import map_field, safe_str, is_allergen_field

TEST_PRODUCT = {
    'product_name': 'Test Protein Powder',
    'brand_name': 'TestBrand',
    'supplier_name': 'TestBrand Ltd',
    'sku_code': 'TB-001',
    'ean_barcode': '5060000000001',
    'case_barcode': None,
    'variant': '500g',
    'rrp': 19.99,
    'trade_price_per_case': 95.00,
    'units_per_case': 6,
    'case_size_description': '6 x 500g',
    'vat_rate': 'Zero',
    'shelf_life_weeks': 52,
    'storage_conditions': 'Ambient',
    'country_of_origin': 'United Kingdom',
    'is_vegan': True,
    'is_vegetarian': True,
    'is_gluten_free': True,
    'is_organic': False,
    'organic_cert_number': None,
    'is_recyclable': True,
    'contains_added_sugar': False,
    'energy_kcal': 400,
    'energy_kj': 1674,
    'fat': 5.0,
    'saturates': 1.0,
    'carbohydrates': 10.0,
    'sugars': 2.0,
    'fibre': 3.0,
    'protein': 75.0,
    'salt': 1.0,
    'serving_size_value': 30,
    'allergen_details': [
        {'allergen': 'Milk', 'may_contain': True}
    ],
    'lead_time_days': None,
    'certifications': ['Vegan', 'Gluten Free'],
}


def test(label, expected, product=TEST_PRODUCT):
    result = map_field(label, product)
    status = 'PASS' if str(result) == str(expected) else 'FAIL'
    print(f'{status} | "{label}" → got "{result}" expected "{expected}"')
    return status == 'PASS'


results = []

results.append(test('Full Product Name *', 'Test Protein Powder'))
results.append(test('Product Name on Pack', 'Test Protein Powder'))
results.append(test('Brand Name', 'TestBrand'))
results.append(test('Supplier Name', 'TestBrand Ltd'))
results.append(test('SKU / Supplier Code', 'TB-001'))
results.append(test('EAN Barcode (Unit)', '5060000000001'))
results.append(test('Individual unit barcode', '5060000000001'))
results.append(test('Case Barcode', 'N/A'))
results.append(test('UPC Barcode Case', 'N/A'))

results.append(test('RRP per Unit (inc VAT)', '19.99'))
results.append(test('MSRP per Unit (USD)', '19.99'))
results.append(test('Wholesale Price per Case (ex VAT)', '95.0'))
results.append(test('Trade Price per Case', '95.0'))
results.append(test('Units per Case', '6'))
results.append(test('VAT Rate', 'Zero'))
results.append(test('Lead Time (days)', 'N/A'))

results.append(test('Shelf Life (weeks from manufacture)', '52'))
results.append(test('Shelf Life from Manufacture days', '364'))
results.append(test('Storage Conditions', 'Ambient'))
results.append(test('Country of Origin', 'United Kingdom'))

results.append(test('Is product Vegan?', 'Yes'))
results.append(test('Vegan', 'Yes'))
results.append(test('Is product Gluten Free?', 'Yes'))
results.append(test('Gluten Free', 'Yes'))
results.append(test('Gluten Free?', 'Yes'))
results.append(test('Is product Organic?', 'No'))
results.append(test('Organic Certification Number', 'N/A'))
results.append(test('Contains Added Sugar?', 'No'))
results.append(test('Is Packaging Recyclable', 'Yes'))
results.append(test('Is packaging recyclable?', 'Yes'))

results.append(test('Energy kcal per 100g', '400'))
results.append(test('Fat g per 100g', '5'))
results.append(test('Protein g per 100g', '75'))
results.append(test('Salt g per 100g', '1'))

results.append(test('Total Fat g per serving', '1.5'))
results.append(test('Total Protein g per serving', '22.5'))

results.append(test('Sodium mg per 100g', '400'))
results.append(test('Sodium mg per serving', '120'))

results.append(test('Added Sugars g per serving', '0'))

results.append(test('Milk', 'May Contain'))
results.append(test('Milk / Dairy', 'May Contain'))
results.append(test('Eggs', 'Not Present'))
results.append(test('Peanuts', 'Not Present'))
results.append(test('Cereals Containing Gluten', 'Not Present'))

print()
passed = sum(results)
total = len(results)
print(f'RESULT: {passed}/{total} tests passed')
if passed < total:
    print('TESTS FAILED — DO NOT DEPLOY')
    exit(1)
else:
    print('ALL TESTS PASSED — SAFE TO DEPLOY')
