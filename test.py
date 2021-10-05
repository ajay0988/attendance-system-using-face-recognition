import csv
"""with open('Attendance.csv', 'r') as file:
    reader = csv.reader(file ,delimiter = '\t')
    for row in reader:
        print(row)"""
import csv
with open('Attendance.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["SN", "Movie", "Protagonist"])
    writer.writerow([1, "Lord of the Rings", "Frodo Baggins"])
    writer.writerow([2, "Harry Potter", "Harry Potter"])