import pandas as pd
import camelot
import fitz
import re
import argparse

#Чтение страниц
class ConversionBackend(object):
    def convert(self, pdf_path, png_path):
        doc = fitz.open(pdf_path) 
        for page in doc.pages():
            pix = page.get_pixmap()  
            pix.save(png_path)

#Обработка конкретной таблицы
def ParseTable(table):
    res = []
    m = len(table.data)
    i = 0
    while i != m:
        tmp = {}
        case = table.data[i][1]
        match = re.search(r'Case Number(.*)', case)
        if match:
            case_value = match.group(1)
            tmp["Case Number"] = case_value

        properties = table.data[i + 1][6]
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

        cost = table.data[i][9]
        match = re.search(r'Cost & Tax Bid(.*)', cost)
        if match:
            cost_value = match.group(1)
            tmp["Cost & Tax Bid"] = cost_value
        i += 2
        res.append(tmp)
    return res        

#Добавление значения в xlsx файл
def AddToExel(data, outfile):
    df = pd.DataFrame(data)
    df_expanded = df.explode('Property')
    df_expanded['Case Number'] = df_expanded['Case Number'].ffill()
    df_expanded['Cost & Tax Bid'] = df_expanded['Cost & Tax Bid'].ffill()

    try:
        existing_data = pd.read_excel(outfile)
        df_to_add = df_expanded[~df_expanded['Property'].isin(existing_data['Property'])]
        df_combined = pd.concat([existing_data, df_to_add], ignore_index=True)
    except FileNotFoundError:
        df_combined = df_expanded

    df_combined.to_excel(outfile, index=False)

def main(input_filename, output_filename):
    tables = camelot.read_pdf(input_filename, 
                            backend=ConversionBackend(), 
                            strip_text='\n', 
                            line_scale=40, 
                            pages='4',
                            copy_text=['h'],)
    
    count_tables = tables.n
    for i in range(0, count_tables):
        datas = ParseTable(tables[i])
        print(datas)
        for data in datas:
            AddToExel(data, output_filename)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Input PDF file")
    parser.add_argument("outfile", help="Output Excel file")
    args = parser.parse_args()
    main(args.file, args.outfile)
