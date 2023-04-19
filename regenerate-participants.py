from openpyxl import load_workbook
import io

try:
    wb = load_workbook(filename = '2023_Homotopy_Type_Theory_Conference.xlsx')
except:
    print("Ensure that the file 2023_Homotopy_Type_Theory_Conference.xlsx is present in the current directory (do not push the spreadsheet to git)")
    exit()

md = open("participants.md", "w")

sh = wb.active

sh.delete_rows(1, 4)

md.write("---\nlayout: page\npermalink: /participants/\ntitle: 'Participants'\n---\n\n")
md.write("Name | Affiliation\n---|---\n")

l = []

for r in sh.rows:
    l.append([r[1].value, r[2].value, r[15].value])

l.sort(key = lambda person : person[1].lower())

for person in l:
    md.write(person[0] + ' ' + person[1] + ' | ' + person[2] + '\n')

md.close()
