import json
import asyncio
import base64
import io

import openpyxl
from openpyxl.cell.cell import MergedCell

from main import (
    map_field,
    safe_str,
    is_allergen_field,
    safe_write,
    is_formula_cell,
    count_formula_cells,
    fill_horizontal_rows,
    fill_using_form_spec,
    FillRequest,
    fill_nlf,
)

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

# Conservative compliance — NASAA must not infer from generic Organic
nasaa_product = {**TEST_PRODUCT, 'is_organic': True, 'certifications': ['Organic', 'Vegan']}
results.append(test('NASAA Organic', 'No', product=nasaa_product))
results.append(test('Australian Certified Organic', 'No', product=nasaa_product))

# Unknown label returns None (Claude fallback candidate)
thr = map_field('THR Licensed', TEST_PRODUCT)
thr_ok = thr is None
print(f'{"PASS" if thr_ok else "FAIL"} | "THR Licensed" → got "{thr}" expected None')
results.append(thr_ok)

print()
print('--- Integration tests ---')

# Formula protection
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Data'
ws['A1'] = 5
ws['B4'] = '=A1*2'
ws['A4'] = 'Product Name'
file_bytes = io.BytesIO()
wb.save(file_bytes)
b64 = base64.b64encode(file_bytes.getvalue()).decode()

req = FillRequest(
    file_base64=b64,
    products=[{'product_name': 'Test Item', 'allergen_details': [], 'certifications': []}],
    retailer_name='Test',
    fill_mode='auto',
    form_spec={
        'data_sheet': 'Data',
        'layout': 'vertical',
        'label_column': 1,
        'value_column': 2,
        'example_rows': [],
        'other_sheets': [],
        'field_map': [],
    },
)
res = asyncio.run(fill_nlf(req))
out_wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(res['file_base64'])))
formula_val = out_wb['Data']['B4'].value
formula_ok = isinstance(formula_val, str) and formula_val.startswith('=')
print(f'{"PASS" if formula_ok else "FAIL"} | Formula cell preserved: {formula_val!r}')
results.append(formula_ok)

# Example row protection
wb2 = openpyxl.Workbook()
ws2 = wb2.active
ws2.title = 'Sheet1'
for col, hdr in enumerate(['Product Name', 'Brand', 'RRP'], start=1):
    ws2.cell(row=8, column=col).value = hdr
ws2.cell(row=9, column=1).value = 'EXAMPLE PRODUCT'
ws2.cell(row=9, column=2).value = 'Demo Brand'
ws2.cell(row=9, column=3).value = '9.99'
file_bytes2 = io.BytesIO()
wb2.save(file_bytes2)
b64_2 = base64.b64encode(file_bytes2.getvalue()).decode()

products_3 = [
    {'product_name': f'Product {i}', 'brand_name': 'BrandX', 'rrp': 10 + i,
     'allergen_details': [], 'certifications': []}
    for i in range(1, 4)
]
req2 = FillRequest(
    file_base64=b64_2,
    products=products_3,
    retailer_name='Dundeis',
    fill_mode='auto',
    form_spec={
        'data_sheet': 'Sheet1',
        'layout': 'horizontal_rows',
        'header_row': 8,
        'first_data_row': 10,
        'example_rows': [9],
        'other_sheets': [],
        'field_map': [],
    },
)
res2 = asyncio.run(fill_nlf(req2))
out2 = openpyxl.load_workbook(io.BytesIO(base64.b64decode(res2['file_base64'])))
ws_out = out2['Sheet1']
example_preserved = ws_out.cell(row=9, column=1).value == 'EXAMPLE PRODUCT'
row10 = ws_out.cell(row=10, column=1).value == 'Product 1'
row11 = ws_out.cell(row=11, column=1).value == 'Product 2'
row12 = ws_out.cell(row=12, column=1).value == 'Product 3'
example_ok = example_preserved and row10 and row11 and row12
print(f'{"PASS" if example_ok else "FAIL"} | Example row untouched + products at 10/11/12')
results.append(example_ok)

# Formula count unchanged with VLOOKUP sheet
wb3 = openpyxl.Workbook()
ws_data = wb3.active
ws_data.title = 'Product 1'
ws_data['A4'] = 'Product Name'
ws_data['B4'] = 'Brand'
ws_data['A5'] = 'Full Product Name'
ws_lookup = wb3.create_sheet('Lookup')
ws_lookup['A1'] = '=VLOOKUP("x",Product_1!A:B,2,FALSE)'
file_bytes3 = io.BytesIO()
wb3.save(file_bytes3)
before_count = count_formula_cells(openpyxl.load_workbook(io.BytesIO(file_bytes3.getvalue())))
b64_3 = base64.b64encode(file_bytes3.getvalue()).decode()
req3 = FillRequest(
    file_base64=b64_3,
    products=[{'product_name': 'Safe Fill', 'brand_name': 'Co', 'allergen_details': [], 'certifications': []}],
    retailer_name='Test',
    form_spec={
        'data_sheet': 'Product 1',
        'layout': 'vertical',
        'label_column': 1,
        'value_column': 2,
        'other_sheets': ['Lookup'],
        'example_rows': [],
        'field_map': [],
    },
)
res3 = asyncio.run(fill_nlf(req3))
after_wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(res3['file_base64'])))
after_count = count_formula_cells(after_wb)
count_ok = after_count >= before_count and 'Lookup' in after_wb.sheetnames
print(f'{"PASS" if count_ok else "FAIL"} | Formula count before={before_count} after={after_count}')
results.append(count_ok)

print()
passed = sum(1 for r in results if r is True)
total = len(results)
print(f'RESULT: {passed}/{total} tests passed')
if passed < total:
    print('TESTS FAILED — DO NOT DEPLOY')
    exit(1)
else:
    print('ALL TESTS PASSED — SAFE TO DEPLOY')
