"""
This file is for our new theme: gets the date creates an Excel file writes it to Excel this is the date
Create by: Miqayel Postoyan
Date: 15 April
"""
import argparse
import xlsxwriter

def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--filename", required=True, help="filename")
    parser.add_argument("-o", "--output", required=True, help="output elsx file")
    parser.add_argument("-s", "--sheet", help="sheet")

    args = parser.parse_args()
    fname = args.filename
    output = args.output
    sheet = args.sheet
    return fname, output, sheet



def get_content(filename):
    with open(filename) as f:
        return f.readlines()


def get_info_dic(n, s, a, p):
    d = {}
    d["name"] = n
    d["surname"] = s
    d["age"] = a
    d["profession"] = p
    return d



def get_info_list(cnt):
    ml = []
    for line in cnt:
        name, surname, age, profession = line.split()
        ml.append(get_info_dic(name, surname, age, profession))
    return ml


def writer_excel(output, info_list, sheet):
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format()
    cell_format2 = workbook.add_format()
    cell_format.set_bold()
    cell_format.set_bg_color('green')
    cell_format2.set_bg_color('yellow')
    worksheet.write(0, 0, "Name", cell_format)
    worksheet.write(0, 1, "Surname", cell_format)
    worksheet.write(0, 2, "Age", cell_format)
    worksheet.write(0, 3, "Profession", cell_format)
    row = 1
    for i in info_list:
        worksheet.write(row, 0, i["name"])
        worksheet.write(row, 1, i["surname"])
        if i["age"] > str(35):
            worksheet.write(row, 2, i["age"], cell_format2)
        else:
            worksheet.write(row, 2, i["age"])
        worksheet.write(row, 3, i["profession"])
        row += 1
    if sheet:
        row2 = 0
        worksheet2 = workbook.add_worksheet(sheet)
        for i in info_list:
            if i["profession"] == sheet:
                worksheet2.write(row2, 0, i["name"])
                worksheet2.write(row2, 1, i["surname"])
                worksheet2.write(row2, 2, i["age"])
                worksheet2.write(row2, 3, i["profession"])
                row2 += 1
    workbook.close()



def main():
    fname, output, sheet = get_arguments()
    cnt = get_content(fname)
    info_name = get_info_list(cnt)
    writer_excel(output, info_name, sheet)




if __name__ == "__main__":
    main()
