from openpyxl import Workbook, load_workbook
import os

# Function to validate marks
def get_marks(subject):

    while True:

        marks = int(input(f"{subject} marks: "))

        if 0 <= marks <= 100:
            return marks

        print("Invalid marks! Enter marks between 0 and 100.")


filename = "student_results.xlsx"

# Check if file exists
if os.path.exists(filename):
    workbook = load_workbook(filename)
    sheet = workbook.active
else:
    workbook = Workbook()
    sheet = workbook.active

    # Heading row
    sheet.append([
        "Name",
        "Tamil",
        "English",
        "Maths",
        "Science",
        "Total",
        "Average",
        "Grade",
        "Status"
    ])

# Number of students
num = int(input("Enter number of students: "))

for i in range(num):

    print(f"\nStudent {i + 1}")

    name = input("Enter student name: ")

    tamil = get_marks("Tamil")
    english = get_marks("English")
    maths = get_marks("Maths")
    science = get_marks("Science")

    total = tamil + english + maths + science
    average = total / 4

    # Grade Calculation
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

    # Pass/Fail Status
    if tamil < 35 or english < 35 or maths < 35 or science < 35:
        status = "Fail"
    else:
        status = "Pass"

    # Add data to Excel
    sheet.append([
        name,
        tamil,
        english,
        maths,
        science,
        total,
        average,
        grade,
        status
    ])

# Save file
workbook.save(filename)

print("\nData added successfully!")