import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
import datetime
import os

def create_paragraph_with_spacing(document, text, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """Adds a paragraph with single line spacing, no space after, and specified alignment."""
    p = document.add_paragraph(text)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.alignment = alignment
    return p

def generate_certificate():
    """Generates the Word document (.docx) with aligned text and prompts for save location."""
    
    # Get values from the form fields
    office_name= entry_office_name.get()
    contractor_name = entry_contractor_name.get()
    contractor_address = entry_contractor_address.get()
    project_name = entry_project_name.get()
    contract_id = entry_contract_id.get()
    project_location = entry_project_location.get()
    agreement_date = entry_agreement_date.get()
    completion_date = entry_completion_date.get()
    contract_amount = entry_contract_amount.get()
    final_contract_amount = entry_final_contract_amount.get()
    authorized_signatory = entry_authorized_signatory.get()
    designation = entry_designation.get()

    if not all([office_name,contract_id,contractor_name, project_name, project_location, authorized_signatory, designation]):
        messagebox.showerror("Error", "Please fill in all mandatory fields.")
        return

    # Create a default filename based on the contractor's name
    default_filename = f"Experience_Certificate_{contractor_name.replace(' ', '_')}.docx"

    # Open a file dialog to ask the user for a save location
    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        initialfile=default_filename,
        title="Save Experience Certificate",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )

    # If the user cancels the dialog, the function stops
    if not save_path:
        return

    # Create a new Word document
    document = Document()

    # Add centered header information
    create_paragraph_with_spacing(document, "Government of Nepal", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "Ministry of Urban Development", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "Department of Urban Development and Building Construction", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "Project Office of the Urban Development & Building Construction", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "{office_name}", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, f"Date: {datetime.date.today().strftime('%B %d, %Y')}", alignment=WD_ALIGN_PARAGRAPH.LRFT)
    create_paragraph_with_spacing(document, "Ref. No.: ", alignment=WD_ALIGN_PARAGRAPH.LEFT)
    create_paragraph_with_spacing(document, "")

    # Add the "TO WHOM IT MAY CONCERN" heading
    create_paragraph_with_spacing(document, "TO WHOM IT MAY CONCERN", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "")

    # Add subject
    create_paragraph_with_spacing(document, "Subject: Experience Certificate", alignment=WD_ALIGN_PARAGRAPH.CENTER)
    create_paragraph_with_spacing(document, "")

    # Add the main body of the certificate (left-aligned)
    body_text = (
        f"This is to certify that {contractor_name}, based in {contractor_address}, has "
        "successfully completed the following construction work for our organization, "
        "Project Office of the Urban Development & Building Construction, {office_name}."
    )
    create_paragraph_with_spacing(document, body_text)
    create_paragraph_with_spacing(document, "")

    # Add project details (left-aligned)
    create_paragraph_with_spacing(document, "Project Details:")
    create_paragraph_with_spacing(document, f"Project Name: {project_name}")
    create_paragraph_with_spacing(document, f"Contract ID: {contract_id}")
    create_paragraph_with_spacing(document, f"Project Location: {project_location}")
    create_paragraph_with_spacing(document, f"Contract Agreement Date: {agreement_date}")
    create_paragraph_with_spacing(document, f"Work Completion Date: {completion_date}")
    create_paragraph_with_spacing(document, f"Contract Amount: {contract_amount}")
    create_paragraph_with_spacing(document, f"Final Contract Amount: {final_contract_amount}")

    # Add major works and quantities (left-aligned)
    create_paragraph_with_spacing(document, "The following quantities represent the major works executed by the contractor under the project:")
    major_works_text = text_quantities.get("1.0", tk.END).strip()
    if major_works_text:
        for line in major_works_text.splitlines():
            if line.strip():
                create_paragraph_with_spacing(document, f"• {line.strip()}")
    create_paragraph_with_spacing(document, "")

    # Add closing remarks (left-aligned)
    create_paragraph_with_spacing(document, f"We wish {contractor_name} every success in their future endeavours.")
    create_paragraph_with_spacing(document, "")
    
    # Add signatory information (right-aligned)
    create_paragraph_with_spacing(document, "Sincerely,", alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    create_paragraph_with_spacing(document, authorized_signatory, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    create_paragraph_with_spacing(document, designation, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # Save the document to the user-specified path
    try:
        document.save(save_path)
        messagebox.showinfo("Success", f"Certificate generated successfully!\nFile saved to: {os.path.abspath(save_path)}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the file: {e}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Experience Certificate Generator")
root.geometry("800x600")

frame = tk.Frame(root, padx=12, pady=12)
frame.pack(fill="both", expand=True)

def create_label_entry(parent, label_text, row, col=0, default_text=""):
    label = tk.Label(parent, text=label_text, anchor="w")
    label.grid(row=row, column=col, sticky="ew", padx=7, pady=2)
    entry = tk.Entry(parent)
    entry.grid(row=row, column=col+1, sticky="ew", padx=7, pady=2)
    if default_text:
        entry.insert(0, default_text)
    return entry

label_heading = tk.Label(frame, text="Experience Certificate Details", font=("Arial", 16, "bold"))
label_heading.grid(row=0, column=0, columnspan=2, pady=10)
entry_office_name = create_label_entry(frame, "Office Name and district:", 1)
entry_contractor_name = create_label_entry(frame, "Contractor Name/Company Name:", 2)
entry_contractor_address = create_label_entry(frame, "Contractor's Address:", 3)
entry_contract_id = create_label_entry(frame, "Contract id:", 4)
entry_project_name = create_label_entry(frame, "Project Name:",5 )
entry_project_location = create_label_entry(frame, "Project Location:", 6)
entry_agreement_date = create_label_entry(frame, "Contract Agreement Date:", 7)
entry_completion_date = create_label_entry(frame, "Work Completion Date:", 8)
entry_contract_amount = create_label_entry(frame, "Contract Amount:", 9)
entry_final_contract_amount = create_label_entry(frame, "Final Contract Amount:", 10)

# These two fields now have default values
entry_authorized_signatory = create_label_entry(frame, "Authorized Signatory Name:", 11, default_text="Sushil Acharya")
entry_designation = create_label_entry(frame, "Designation:", 12, default_text="Engineer")

label_quantities = tk.Label(frame, text="Major Works and Quantities (up to 15 items, one per line):", anchor="w")
label_quantities.grid(row=11, column=0, sticky="ew", padx=5, pady=5)
text_quantities = scrolledtext.ScrolledText(frame, width=50, height=10, wrap=tk.WORD)
text_quantities.grid(row=12, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)

text_quantities.insert(tk.END, "Earthwork Excavation: [Quantity] [cubic meters]\n"
                                "Stone Soling: [Quantity] [cubic meters]\n"
                               "Stone Masonry : [Quantity] [meters]\n"
                               "Reinforced Cement Concrete (RCC): [Quantity] [cubic meters]\n"
                               "Reinforcement Steel: [Quantity] [kilograms]\n"
                               "Brickwork: [Quantity] [ cubic meters/square meters]")

generate_button = tk.Button(frame, text="Generate Certificate", command=generate_certificate)
generate_button.grid(row=13, column=0, columnspan=2, pady=10)

frame.columnconfigure(1, weight=1)
frame.rowconfigure(14, weight=1)

root.mainloop()
