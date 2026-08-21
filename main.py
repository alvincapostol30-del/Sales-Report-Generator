import csv
from openpyxl import Workbook
from datetime import datetime

#read csv file and print values
with open("Input/sales.csv", newline="") as file:
    reader = csv.DictReader(file, delimiter="\t")
    for row in reader:
        print(f"Product: {row['Product']}")
        print(f"Salesperson: {row['Salesperson']}")
        print(f"Order ID: {row['Quantity']}")
        print(f"Unit Price: {row['Unit Price']}")
        print("-" * 40)

#This block of code will create/format an Excel file and save to output folder 
workbook = Workbook()
sheet = workbook.active
sheet.title = "Sales Report"

#Define headers
headers = [
    "Order ID",
    "Date",
    "Salesperson",
    "Department",
    "Product",
    "Quantity",
    "Unit Price",
    "Total"

]
#get list length
max_item = len(headers)

#Loop to list
for r in range(1, max_item+1):
    print(headers[r-1])
    sheet.cell(5,r).value = headers[r-1]


#Extract todays date with hours/mm/ss
generated_on = datetime.now().strftime("%d-%b-%Y %I:%M %p")


#Create Sheet description
sheet.cell(1,1).value = "Sales Report" 
sheet.cell(2,1).value ="Company"
sheet.cell(2,2).value = "ABC Corporation"
sheet.cell(3,1).value = "Generated On"
sheet.cell(3,2).value = generated_on
