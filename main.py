from openpyxl import Workbook, load_workbook
import os

filename = "student_results.xlsx"

# Check if file already exists
if os.path.exists(filename):
    workbook = load_workbook(filename)
    sheet = workbook.active
else:
    workbook = Workbook()
    sheet = workbook.active

    # Create heading row only once
    sheet.append([
        "Name",
        "Tamil",
        "English",
        "Maths",
        "Science",
        "Total",
        "Average",
        "Grade"
    ])

# Number of students
num = int(input("Enter number of students: "))

for i in range(num):

    print("\nStudent", i + 1)

    name = input("Enter student name: ")

    tamil = int(input("Tamil marks: "))
    english = int(input("English marks: "))
    maths = int(input("Maths marks: "))
    science = int(input("Science marks: "))

    total = tamil + english + maths + science
    average = total / 4

    if average >= 90:
        grade = "A+"
    elif average >= 75:
        grade = "A"
    elif average >= 60:
        grade = "B"
    elif average >= 40:
        grade = "C"
    else:
        grade = "Fail"

    sheet.append([
        name,
        tamil,
        english,
        maths,
        science,
        total,
        average,
        grade
    ])

workbook.save(filename)

print("\nData added successfully!")