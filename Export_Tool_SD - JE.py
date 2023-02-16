# Import xlwt library (This library only read and write to xls)
import xlwt

# Open the txt file, create a workbook, and rename it
txt_file = open("ScienceDirect 1-100 - JE.txt", "r")
wb = xlwt.Workbook()
sheet1 = wb.add_sheet("Sheet 1")

# Write header to Excel file
sheet1.write(0, 0, "Year")
sheet1.write(0, 1, "Article Name")
sheet1.write(0, 2, "Abstract")
sheet1.write(0, 3, "DOI")
sheet1.write(0, 4, "Journal Name")

for i in range(15):
    sheet1.write(0, (5 + i), f"Keyword {i + 1}")

# Set accumulator m,n to 1
n = 1
m = 1

# Loop through txt file, extract year, article name, doi, and journal name, strip them, and write them to Excel file
for line in txt_file:
    if n == 2:
        line = line.rstrip(",\n")
        sheet1.write(m, 1, line)
    elif n == 3:
        line = line.rstrip(",\n")
        sheet1.write(m, 4, line)
    elif n == 5:
        line = line.rstrip(",\n")
        sheet1.write(m, 0, line)
    elif n == 8:
        line = line.rstrip(".\n")
        sheet1.write(m, 3, line)
    elif n == 10:
        line = line.replace("Abstract: ", "")
        line = line.rstrip(".\n")
        sheet1.write(m, 2, line)
    elif n == 11 and line != "":
        line = line.replace("Keywords: ", "")
        line_list = line.split("; ")
        for num in range(len(line_list)):
            sheet1.write(m, (5 + num), line_list[num])

    n += 1

    if n == 12:
        txt_file.readline()
        m += 1
        n = 1

# Save to Excel file
wb.save("ScienceDirect 1-100 - JE.xls")

# Close txt file
txt_file.close()
