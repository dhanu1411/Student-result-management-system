from openpyxl import Workbook

# Create Excel workbook
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

    # Grade calculation
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

    # Save data into Excel
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

# Save Excel file
workbook.save("student_results.xlsx")

print("\nData saved successfully into student_results.xlsx")