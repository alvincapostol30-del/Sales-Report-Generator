import csv
with open("Input/sales.csv", newline="") as file:
    reader = csv.DictReader(file, delimiter="\t")
    for row in reader:
        print(f"Product: {row['Product']}")
        print(f"Salesperson: {row['Salesperson']}")
        print(f"Order ID: {row['Quantity']}")
        print(f"Unit Price: {row['Unit Price']}")
        print("-" * 40)