import csv
with open("Input/sales.csv", newline="") as file:
    reader = csv.DictReader(file, delimiter="\t")
    for row in reader:
        print(row)