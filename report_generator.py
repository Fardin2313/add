from docx import Document

print("------ Report Generator ------")

# Taking user input
name = input("Enter your name: ")
topic = input("Enter report topic: ")
details = input("Write a short report description: ")

# Creating a new Word file
doc = Document()

doc.add_heading("Student Report", level=1)

doc.add_paragraph(f"Name: {name}")
doc.add_paragraph(f"Topic: {topic}")
doc.add_paragraph("\nReport Details:")
doc.add_paragraph(details)

# Saving file
filename = topic.replace(" ", "_") + "_Report.docx"
doc.save(filename)

print(f"\nReport Created Successfully!\nFile saved as: {filename}")
