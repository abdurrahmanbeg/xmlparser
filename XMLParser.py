import xml.etree.ElementTree as ET
import csv
import os

title = 'Type','Order_Number', 'SelRefFrac','Module','Position','Category_Name','Category_ID','Category_shortCode','Vaue_Name','Value_Unit_ID', 'Value_Shortcode'
orderNo = 0

path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
for filename in os.listdir(path):
    if not filename.endswith('.xml'): continue
    fullname = os.path.join(path, filename)
    print('Parsing: ' + fullname)
    root = ET.parse(fullname)

    for orderNumber in root.findall('Lead_Order_Number'):
        orderNo =  orderNumber.text
    with open(orderNumber.text + '.csv', 'w', newline='') as outfileCSV:
        writer = csv.writer(outfileCSV)
        writer.writerow(title)
        outfileCSV.close()
    for products in root.findall('Products'):
        for beginProducts in products.findall('Begin_Product'):
            for selection in beginProducts.findall('Selection'):
                selRefFrac = selection.get('selRefFrac')
                for selectionChild in selection.findall('Configuration_Data'):
                    for module in selectionChild.findall('Module'):
                        for category in module.findall('Category'):
                             for item in category.findall('Item'):
                                  with open(orderNumber.text + '.csv', 'a', newline='') as outfileCSV:
                                    writer = csv.writer(outfileCSV)
                                    row = 'Configuration_Data',orderNo, selRefFrac, module.get('name'), module.get('position'),category.get('nID'), category.get('name'),category.get('shortCode'),item.get('nID'),item.get('name'),item.get('shortCode')
                                    writer.writerow(row)
                                    outfileCSV.close()
                for selectionChild in selection.findall('Performance_Data'):
                    for spectrum in selectionChild.findall('Spectrum_Perf_Data'):
                        for module in spectrum.findall('Perf_Module'):
                            for category in module.findall('Perf_Category'):
                                for value in category.findall('Perf_Value_Unit'):
                                    with open(orderNumber.text + '.csv', 'a', newline='') as outfileCSV:
                                        writer = csv.writer(outfileCSV)
                                        row = 'Performance_Data',orderNo,selRefFrac,module.get('name'), module.get('position'),category.get('nID'), category.get('name'),category.get('shortCode'),value.get('unit_id'),value.get('unit_desc'),value.text
                                        writer.writerow(row)
                                        outfileCSV.close()