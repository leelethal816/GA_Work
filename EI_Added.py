# Import openpyxl library
import openpyxl

# Read txt file and open Excel
txt_file = open("Emerald Insights.txt", "r")
wb = openpyxl.Workbook()

# Make Excel worksheet active
sheet = wb.active

# Assign name to header
sheet["A1"] = "Abstract"

# Set accumulator
i = 2

# Read from txt file for abstract and write it into Excel worksheet
for line in txt_file:
    if "AB  - " in line:
        line = line.replace("AB  - Purpose ", "")
        sheet["A" + str(i)] = line
        i += 1

# Save Excel file and close txt file
wb.save("Added EI.xlsx")
txt_file.close()
