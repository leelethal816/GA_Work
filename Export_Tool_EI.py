# Import xlwt library (This library only read and write to xls)
import xlwt

# Open the txt file, create a workbook, and rename it
txt_file = open("Emerald Insights.txt", "r")
wb = xlwt.Workbook()
sheet1 = wb.add_sheet("Sheet 1")

# Write header to Excel file
sheet1.write(0, 0, "Year")
sheet1.write(0, 1, "Article Name")
sheet1.write(0, 2, "DOI")
sheet1.write(0, 3, "Journal Name")

# Set accumulator m to 1
m = 1

# Loop through txt file, extract year, article name, doi, and journal name, strip them, and write them to Excel file
for line in txt_file:
    if "UR  - " in line:
        line = line.replace("UR  - ", "")
        sheet1.write(m, 2, line)
    elif "PY  - " in line:
        line = line.replace("PY  - ", "")
        sheet1.write(m, 0, line)
    elif "TI  - " in line:
        line = line.replace("TI  - ", "")
        sheet1.write(m, 1, line)
    elif "T2  - " in line:
        line = line.replace("T2  - ", "")
        sheet1.write(m, 3, line)
    elif "ER  - " in line:
        txt_file.readline()
        txt_file.readline()
        m += 1

# Save to Excel file
wb.save("Emerald Insights.xls")

# Close txt file
txt_file.close()
