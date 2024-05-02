import pandas as pd

import os
import glob
import shutil
import sys

from PyPDF2 import PdfMerger

import math

def merge_pdfs(paths, output):
    merger = PdfMerger()
    for path in paths:
        merger.append(path)
    merger.write(output)
    merger.close()

def run():

    # อ่าน Excel ต้นทาง

    df = pd.read_excel('order.xlsx')

    # ดึงค่า 'Purchase Order','Bill.Doc.'
    
    data_dict = df[['Purchase Order','Bill.Doc.']].to_dict(orient='records')
    new_data_list_1 = []
    current_main = None
    current_sub = []
    for item in data_dict:
        if math.isnan(item['Purchase Order']) and math.isnan(item['Bill.Doc.']) :
            continue      
        if not math.isnan(item['Purchase Order']):
            if current_main is not None:
                new_data_list_1.append({'po': current_main, 'bill': current_sub})
                current_sub = []
            current_main = item['Purchase Order']
        current_sub.append(item['Bill.Doc.'])
    if current_main is not None:
        new_data_list_1.append({'po': current_main, 'bill': current_sub})
    for item in new_data_list_1:
        item['po'] = str(int(item['po']))
        item['bill'] = [str(int(po)) for po in item['bill']]

    # ดึงค่า 'Bill.Doc.','Del. no.'

    data_dict = df[['Bill.Doc.','Del. no.']].to_dict(orient='records')
    new_data_list_2 = []
    current_main = None
    current_sub = []
    for item in data_dict:
        if math.isnan(item['Bill.Doc.']) and math.isnan(item['Del. no.']) :
            continue     
        if not math.isnan(item['Bill.Doc.']):
            if current_main is not None:
                new_data_list_2.append({'bill': current_main, 'del': current_sub})
                current_sub = []
            current_main = item['Bill.Doc.']
        current_sub.append(item['Del. no.'])
    if current_main is not None:
        new_data_list_2.append({'bill': current_main, 'del': current_sub})
    for item in new_data_list_2:
        item['bill'] = str(int(item['bill']))
        item['del'] = [str(int(del_no)) for del_no in item['del']]

    # ดึงค่า 'Del. no.','Inv.list'

    data_dict = df[['Del. no.','Inv.list']].to_dict(orient='records')
    new_data_list_3 = []
    current_main = None
    current_sub = []
    for item in data_dict:
        if not math.isnan(item['Del. no.']):
            if current_main is not None:
                new_data_list_3.append({'del': current_main, 'inv': current_sub})
                current_sub = []
            current_main = item['Del. no.']
        current_sub.append(item['Inv.list'])
    if current_main is not None:
        new_data_list_3.append({'del': current_main, 'inv': current_sub})
    for item in new_data_list_3:
        item['del'] = str(int(item['del']))
        # item['inv'] = [str(int(inv)) for inv in item['inv']]
        item['inv'] = [str(int(inv)) if not pd.isna(inv) else inv for inv in item['inv']]

    delivery_invoice_dict = {}
    for item in new_data_list_3:
        delivery_invoice_dict[item['del']] = item['inv']
    merged_list = []
    for item in new_data_list_2:
        invoices = []
        for delivery in item['del']:
            invoices.extend(delivery_invoice_dict.get(delivery, []))
        merged_item = {
            'bill': item['bill'],
            'del': item['del'],
            'inv': invoices
        }
        merged_list.append(merged_item)
    
    # Merge Data 3 ชุดเข้าด้วยกัน

    bill_po_dict = {}
    for item in new_data_list_1:
        for bill in item['bill']:
            bill_po_dict[bill] = item['po']
    final_list = []
    for item in merged_list:
        final_item = {
            'bill': item['bill'],
            'del': item['del'],
            'inv': item['inv'],
            'po': bill_po_dict.get(item['bill'], None)
        }
        final_list.append(final_item)
    
    # ตรวจสอบค่า NaN

    for item in final_list:
        for key, value in item.items():
            if isinstance(value, list):
                item[key] = [val for val in value if not pd.isna(val)]

    # เติมหลักเอกสาร

    for item in final_list:
        item['po'] = item['po'].zfill(7)
        item['bill'] = item['bill'].zfill(10)
        for i in range(len(item['del'])):
            item['del'][i] = item['del'][i].zfill(10)
        for i in range(len(item['inv'])):
            item['inv'][i] = item['inv'][i].zfill(10)
        
    print('final_list =',final_list)

    # Merge File

    # df = pd.read_excel('order.xlsx')

    df['po'] = df['Purchase Order'].fillna(0).apply(lambda x: '{:.0f}'.format(x).zfill(7))
    df['bill'] = df['Bill.Doc.'].fillna(0).apply(lambda x: '{:.0f}'.format(x).zfill(10))
    df['del'] = df['Del. no.'].fillna(0).apply(lambda x: '{:.0f}'.format(x).zfill(10))
    df['inv'] = df['Inv.list'].fillna(0).apply(lambda x: '{:.0f}'.format(x).zfill(10))

    for list_bill in final_list:
        print('----------------------------------------------------------------')

        merge_list = []
        error_list = []
        
        files = glob.glob(f'doc_bill/{list_bill['bill']}.pdf')
        if len(files) == 0 :
            print(list_bill['bill'] , 'dont find')
            error_list.append(list_bill['bill'])
            df.loc[df['bill'].astype(str) == list_bill['bill'], 'bill_note'] = 'Not Found'
        else:
            print(list_bill['bill'])
            merge_list+=files

        for list_del in list_bill['del']:
            files = glob.glob(f'doc_del/{list_del}.pdf')
            # print(files)
            if len(files) == 0 :
                print(list_del , 'dont find')
                error_list.append(list_del)
                df.loc[df['del'].astype(str) == list_del, 'del_note'] = 'Not Found'
            else:
                print(list_del)
                merge_list+=files

        for list_inv in list_bill['inv']:
            files = glob.glob(f'doc_inv/{list_inv}.pdf')
            # print(files)
            if len(files) == 0 :
                print(list_inv , 'dont find')
                error_list.append(list_inv)
                df.loc[df['inv'].astype(str) == list_inv, 'inv_note'] = 'Not Found'
            else:
                print(list_inv)
                merge_list+=files

        files = glob.glob(f'doc_po/{list_bill['po']}.pdf')
        if len(files) == 0 :
            print(list_bill['po'] , 'dont find')
            error_list.append(list_bill['po'])
            df.loc[df['po'].astype(str) == list_bill['po'], 'po_note'] = 'Not Found'
        else:
            print(list_bill['po'])
            merge_list+=files

        print('merge_list =',merge_list)
        print('error_list =',error_list)

        if len(error_list) == 0 :
            
            destination_directory = 'merge_files'

            if not os.path.exists(destination_directory):
                os.makedirs(destination_directory)
                
            output_file = f'merge_files/{list_bill['bill']}-merge.pdf'  # ระบุชื่อไฟล์ที่ต้องการบันทึกผลลัพธ์
            merge_pdfs(merge_list, output_file)
            print('Create',f'merge_files/{list_bill['bill']}-merge.pdf')
        else:
            print(f'{list_bill['bill']} เอกสารไม่ครบ')
    
    df = df.drop(columns=['po','bill','del','inv'])

    df.to_excel('output_merge/order-merge.xlsx', index=False)

if __name__ == "__main__":
    
    run()