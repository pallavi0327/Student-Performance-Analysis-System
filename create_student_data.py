import pandas as pd
import numpy as np
import random

# Sample first names
first_names = ["James", "John", "Robert", "Michael", "William", "David", "Richard", 
               "Joseph", "Thomas", "Charles", "Mary", "Patricia", "Jennifer", 
               "Linda", "Elizabeth", "Barbara", "Susan", "Jessica", "Sarah", "Karen",
               "Emily", "Daniel", "Matthew", "Christopher", "Andrew", "Joshua"]

# Sample last names
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
              "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez",
              "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin"]

# Classes
classes = ["Grade 10(1)", "Grade 10(2)", "Grade 10(3)", 
           "Grade 11(1)", "Grade 11(2)", "Grade 11(3)",
           "Grade 12(1)", "Grade 12(2)", "Grade 12(3)"]

# Generate 30 student records
students = []
for i in range(1, 31):
    student = {
        "Student ID": f"2023{i:03d}",
        "Name": f"{random.choice(first_names)} {random.choice(last_names)}",
        "Class": random.choice(classes),
        "Mathematics": random.randint(55, 98),
        "Physics": random.randint(50, 96),
        "Chemistry": random.randint(60, 97),
        "Biology": random.randint(58, 95),
        "English": random.randint(65, 99),
        "History": random.randint(45, 92),
        "Geography": random.randint(50, 90)
    }
    students.append(student)

# Create DataFrame
df = pd.DataFrame(students)

# Calculate additional columns
subjects = ["Mathematics", "Physics", "Chemistry", "Biology", "English", "History", "Geography"]
df["Total Score"] = df[subjects].sum(axis=1)
df["Average Score"] = df[subjects].mean(axis=1).round(2)

# Add ranking
df["Rank"] = df["Total Score"].rank(ascending=False, method="min").astype(int)

# Sort by rank
df = df.sort_values("Rank").reset_index(drop=True)

# Save to files
df.to_excel("student_grades.xlsx", index=False)
df.to_csv("student_grades.csv", index=False)

# Print summary
print("=" * 50)
print("STUDENT GRADE DATA FILES CREATED SUCCESSFULLY!")
print("=" * 50)
print(f"\n✅ Created 2 files in current directory:")
print(f"   1. student_grades.xlsx (Excel format)")
print(f"   2. student_grades.csv (CSV format)")
print(f"\n📊 File contains {len(df)} student records")
print(f"\n📋 Sample of the data:")
print("-" * 40)
print(df.head(5).to_string(index=False))
print(f"\n📈 Statistics:")
print(f"   Total Students: {len(df)}")
print(f"   Average Total Score: {df['Total Score'].mean():.1f}")
print(f"   Highest Score: {df['Total Score'].max()}")
print(f"   Lowest Score: {df['Total Score'].min()}")
print("\n🎯 To use these files:")
print("   1. Run the Student Grade Analysis System")
print("   2. Click 'File → Import Data'")
print("   3. Select student_grades.xlsx or student_grades.csv")
print("\n📁 Files will be saved in:", pd.__file__[:pd.__file__.rfind('/')])
print("=" * 50)