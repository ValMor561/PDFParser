import pandas as pd
import camelot
import fitz
import re
import os
import shutil
from datetime import datetime

#Чтение страниц
class ConversionBackend(object):
    def convert(self, pdf_path, png_path):
        doc = fitz.open(pdf_path) 
        for page in doc.pages():
            pix = page.get_pixmap()  
            pix.save(png_path)

def check_status(value, tmp):
    if "$" in value:
        parts = value.split(" - ")
        tmp['Status'] = parts[0]
        tmp['Price'] = float(parts[1].replace('$','').replace(',',''))
    elif " to " in value:
        parts = value.split(" to ")
        tmp['Status'] = parts[0]
        tmp['Date'] = pd.to_datetime(parts[1])
    else:
        tmp['Status'] = value
    return tmp
#Обработка конкретной таблицы
def ParseTable(table):
    res = []
    m = len(table.data)
    i = 0
    while i != m:
        tmp = {}
        sale = table.data[i][0].replace("\n", " ")
        match = re.search(r'Sale(.*)', sale)
        if match:
            sale_value = match.group(1)
            tmp["Sale"] = sale_value

        case = table.data[i][1].replace("\n", " ")
        match = re.search(r'Case Number(.*)', case)
        if match:
            case_value = match.group(1)
            tmp["Case Number"] = case_value

        sale_type = table.data[i][2].replace("\n", " ")
        match = re.search(r'Sale Type(.*)', sale_type)
        if match:
            sale_type_value = match.group(1)
            if sale_type_value == '':
                sale_type_value = table.data[i + 1][2].replace("\n", " ")
            tmp["Sale Type"] = sale_type_value

        tmp['Status'] = ''
        tmp['Date'] = ''
        tmp['Price'] = ''
        status = table.data[i][5].replace("\n", " ")
        match = re.search(r'Status(.*)', status)
        if match:
            status_value = match.group(1)
            if status_value == '':
                status_value = table.data[i + 1][5].replace("\n", " ")
            tmp = check_status(status_value, tmp)
        
        tracts = table.data[i][7].replace("\n", " ")
        match = re.search(r'Tracts(.*)', tracts)
        if match:
            tracts_value = match.group(1)
            tmp["Tracts"] = tracts_value
        
        cost_value = table.data[i][8].replace("\n", " ") if table.data[i][8].replace("\n", " ") != "" else table.data[i + 1][9].replace("\n", " ")
        tmp["Cost & Tax Bid"] = float(cost_value.replace("$", "").replace(",", ""))
        
        svs_value = table.data[i+1][11].replace("\n", " ")
        tmp["SVS"] = True if svs_value == "X" else False
        
        nine_two_value = table.data[i+1][13].replace("\n", " ")
        tmp["3129.2"] = True if nine_two_value == "X" else False
        
        nine_three_value = table.data[i+1][15].replace("\n", " ")
        tmp["3129.3"] = True if nine_three_value == "X" else False
        
        ok_value = table.data[i+1][17].replace("\n", " ")
        tmp["OK"] = True if ok_value == "X" else False

        plain = table.data[i+2][1].replace("\n", " ")
        match = re.search(r'Plaintiff\(s\):(.*)', plain)
        if match:
            plain_value = match.group(1)
            tmp["Plaintiff(s)"] = plain_value
        
        attorney = table.data[i][3].replace("\n", " ")
        match = re.search(r'Attorney for the Plaintiff:(.*)', attorney)
        if match:
            attorney_value = match.group(1)
            tmp["Attorney for the Plaintiff"] = attorney_value
        else:
            attorney = table.data[i+2][3].replace("\n", " ")
            match = re.search(r'Attorney for the Plaintiff:(.*)', attorney)
            if match:
                attorney_value = match.group(1)
                tmp["Attorney for the Plaintiff"] = attorney_value

        dependatnts = table.data[i+2][4].replace("\n", " ")
        match = re.search(r'Defendant\(s\):(.*)', dependatnts)
        if match:
            dependatnts_value = match.group(1)
            if dependatnts_value == '':
                dependatnts_value = table.data[i+2][5].replace("\n", " ")
            tmp["Defendant(s)"] = dependatnts_value

        properties = table.data[i+2][6].replace("\n", " ")
        match = re.search(r'Property (.*)', properties)
        if match:
            propeties_value = match.group(1)
            properies_split = re.split(r'(?<=PA\s\d{5})', propeties_value)
            if properies_split[len(properies_split) - 1] == '':
                properies_split.pop()
            if len(properies_split) == 0:
                properies_split = ["нет адреса"]
            tmp["Property"] = properies_split
        else:
            tmp["Property"] = ["нет адреса"]

        municipality = table.data[i+2][9].replace("\n", " ")
        match = re.search(r'Municipality(.*)', municipality)
        if match:
            municipality_value = match.group(1)
            tmp["Municipality"] = municipality_value

        parcel = table.data[i+2][11].replace("\n", " ")
        match = re.search(r'Parcel/Tax ID:(.*)', parcel)
        if match:
            parcel_value = match.group(1)
            tmp["Parcel/Tax ID"] = parcel_value

        i += 3
        res.append(tmp)
    return res        

#Добавление значения в xlsx файл
def AddToExel(data, outfile):
    df = pd.DataFrame(data)
    df_expanded = df.explode('Property')

    try:
        existing_data = pd.read_excel(outfile)
        df_to_add = df_expanded[~df_expanded['Property'].isin(existing_data['Property'])]
        df_combined = pd.concat([existing_data, df_to_add], ignore_index=True)
    except FileNotFoundError:
        df_combined = df_expanded

    df_combined.to_excel(outfile, index=False)

def CheckPDF(pdf_file):
    tables = camelot.read_pdf(pdf_file, 
                                backend=ConversionBackend(), 
                                line_scale=60, 
                                pages='all',
                                copy_text=['h'],)           
    count_tables = tables.n
    for i in range(0, count_tables):
        datas = ParseTable(tables[i])
        for data in datas:
            filename = os.path.splitext(os.path.basename(pdf_file))[0]
            if not os.path.exists('result'):
                os.makedirs('result')
            AddToExel(data, f"result/{filename}.xlsx")

def main():
    files = os.listdir(f"pdf")
    pdf_files = []
    for file in files:
        if ".pdf" in file:
            pdf_files.append(f"pdf/{file}")

    for pdf_file in pdf_files:
        CheckPDF(pdf_file)
        if not os.path.exists('checked'):
            os.makedirs('checked')
        filename = os.path.splitext(os.path.basename(pdf_file))[0]
        shutil.move(pdf_file, f"checked/{filename}.pdf")
        print(f"[+]Файл {filename}.pdf был проверен\n")
    print("[+] Все файлы проверены")

if __name__ == "__main__":
    main()