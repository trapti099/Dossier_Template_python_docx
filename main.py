from data_from_api import *
from functions import *
from constants import *

# Create a new Document
doc = Document()

# Add footer image
add_image_to_footer(doc,footer_img_path)
doc.add_picture(heading_img_path,width=Cm(4.63),height = Cm(1.05))  # Add heading image
profile_text_paragraph = doc.add_paragraph()
run = profile_text_paragraph.add_run()
inline_shape = run.add_picture(profile_img_path,width=Cm(4.61),height = Cm(4.61)) # Add profile photo
new_paragraph(obj=doc,text=f"{person_name}\n",boolean_bold=True,r=128,g=0,b=128)
new_paragraph(obj=doc,text=f"{designation},\n{company}\n\nLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

# Add a new page
doc.add_page_break()

for table in doc.tables:
    format_table(table)

# Add profile photo for the second page

profile_picture = doc.add_picture(candidate_profile_img_path, width=Inches(2))

# Add the "Personal Data" header as a paragraph
new_paragraph(obj=doc,text="Personal Data:",boolean_bold=True,r=128,g=0,b=128)

# Add a table with three columns
table = doc.add_table(rows=0, cols=3)

# Add rows for personal data
personal_data = ["Date Of Birth","Nationality","Location","Languages"]
personal_data_api = [dob,nationality,location,language]
for i in range(len(personal_data)):
    add_row(table,personal_data[i],":",personal_data_api[i])

for cell in table.columns[2].cells:
    cell.width = Cm(11.0)

# Add Educational Data
new_paragraph(obj=doc,text="Educational Data:",text_size=12,boolean_bold=True,r=128,g=0,b=128)

# Add a table with three columns
table = doc.add_table(rows=0, cols=3)

add_educational_row(table, "1997(N/A)", "Lorem ipsum dolor sit amet, Mumbai(N/A)", "Lorem ipsum dolor sit amet(N/A)")
add_educational_row(table, "1992(N/A)", "Lorem ipsum dolor sit amet, Mumbai(N/A)", "Lorem ipsum dolor sit amet(N/A)")
for cell in table.columns[2].cells:
    cell.width = Cm(11.0)

# Add image
doc.add_picture(employee_data_img_path, width=Inches(7))

exp_count = 1
desig_count = 0
for i in arr_res:
    Key_exp_para=new_paragraph(obj=doc,text=f"Experience {exp_count}:",text_size=16,boolean_bold=True,boolean_italic=True,r=128,g=0,b=128)
    designation = data['data']['profile']['positions'][desig_count]['title']
    company = data['data']['profile']['positions'][desig_count]['company']['name']
    company_profile_paragraph = new_paragraph(obj=doc,text="Company Profile:",text_size=11,boolean_bold=True,boolean_italic=True,r=128,g=0,b=128)
    add_text_to_para(obj=company_profile_paragraph,text=f" {designation},{company}",text_size=10,boolean_italic=True)
    Key_resp_para=new_paragraph(obj=doc,text="Position Summary:",text_size=12,boolean_bold=True,boolean_italic=True,r=128,g=0,b=128)
    add_text_to_para(obj=Key_resp_para,text=f"\n{i}",text_size=10,boolean_italic=False)
    exp_count+=1
    desig_count +=1

# Add image
profile_picture = doc.add_picture(additional_data_img_path, width=Inches(2))

# Add a table with three columns
table = doc.add_table(rows=0, cols=3)

# Add rows for additional data with bullet points
add_row(table, "Location Preference", ":", pref_loc)
add_row(table, "CTC*", ":", ctc)
add_row(table, "Contact Details", ":", f"{email}, {phone_number}")
add_additional_row(table, "2019(N/A)", ":","Lorem ipsum dolor sit amet, Mumbai(N/A)", "Lorem ipsum dolor sit amet(N/A)\nLorem ipsum dolor sit amet(N/A)\n •  Lorem ipsum dolor sit amet(N/A)")
add_additional_row(table, "1992-1994(N/A)", ":", "Lorem ipsum dolor sit amet, Mumbai(N/A)", "Lorem ipsum dolor sit amet\n •  Lorem ipsum dolor sit amet")
for cell in table.columns[2].cells:
    cell.width = Cm(11.0)

# Add a new page
doc.add_page_break()
doc.add_picture(background_image_path, width=Cm(20),height = Cm(23.14))
# Save the document
file_name = person_name + '_' + personal_ID
doc.save(f"{file_name}.docx")
print(f"Your document is saved with the name {file_name}.docx")
