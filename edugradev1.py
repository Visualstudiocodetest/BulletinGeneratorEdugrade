import os
import sys
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import sendmail as sm
import asyncio
import getpass
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

print("Lancement du programme...")

classename = str(input("Entrer le nom de la classe : "))

sessionnumber = str(input("Entrer le numéro de la session : "))

semester = str(input("Entrer le numéro du semestre 1/2: "))

# ask user if he wants to send the email
send_email_bool = (input("Voulez vous envoyer les emails ? (Oui/Non) ")).upper()
if send_email_bool in ['OUI', 'O', 'OU', 'UI']:
    send_email_bool = True
    # Get the email address and password from the user
    from_addr = input("entrer votre adresse mail Satom : ")
    emailpassword = getpass.getpass("Entrer votre mot de passe : ")
else :
    send_email_bool = False
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def generate_student_pdf(datagrades, pdf_file, datastudent, image_path):
    # Create PDF document
    pdf = SimpleDocTemplate(pdf_file, pagesize=letter)

    # Create table with student grades
    table = Table(datagrades)

    # Add style to the table
    style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)])
    table.setStyle(style)

    # Define different styles for different sections
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle('name_style', parent=styles['Heading1'], fontSize=14)
    adress_style = ParagraphStyle('adress_style', parent=styles['BodyText'], fontSize=12, borderPadding=(10, 10, 10, -20),
                                  borderColor='black', borderWidth=1, backColor='lightgrey', alignment=TA_RIGHT)
    academic_year_text_style = ParagraphStyle('academic_year_text__style', parent=styles['BodyText'], fontSize=12,
                                              borderPadding=(5, -30, 5, -30),
                                              borderColor='black', borderWidth=1, backColor='lightgrey',
                                              alignment=TA_CENTER)

    ects_style = ParagraphStyle('ects_style', parent=styles['BodyText'], fontSize=15, borderPadding=(5, 5, 5, 5),
                                alignment=TA_CENTER)

    title_info_style = ParagraphStyle('title_info_style', fontSize=15, borderPadding=(5, 5, 5, 5),
                                      alignment=TA_LEFT, underline=True, underlineColor='black')

    info_style = ParagraphStyle('info_style', parent=styles['BodyText'], fontSize=12, borderPadding=(5, 5, 5, 5),
                                alignment=TA_LEFT)

    # Add student image in the top left corner
    student_image = Image(image_path, width=150, height=75)
    student_image.hAlign = 'LEFT'

    schooladress = Paragraph("ESTIAM GENEVE <br/><br/>Quai du Seujet 10 CH-1201 Genève ", adress_style)

    # Create a table with one row and two columns
    header_table = Table([[student_image, schooladress]])

    # Add text for academic year
    academic_year_text = Paragraph(f"Bulletins de notes Année académique {int(datetime.now().year) - 1}-"
                                   f"{datetime.now().year}", academic_year_text_style)

    # Convert the birthdate to a date object and format it as a string
    birthdate = datetime.strptime(datastudent[0][1], "%Y-%m-%d %H:%M:%S").date()

    datastudent[3][1] = datastudent[3][1].capitalize()
    datastudent[2][1] = datastudent[2][1].upper()
    student_name = Paragraph(f"Étudiant : {datastudent[3][1]} {datastudent[2][1]} <br/>"
                             f" Date de naissance : {birthdate} <br/>"
                             f"Adresse : {datastudent[1][1]}", name_style)

    # Create a table with one row and two columns
    # Add student campus, class, semester, session block modify for you need
    class_info_block = Paragraph(
        f"Campus: Geneve <br/>Classe: {classename}<br/> Semestre: {semester}<br/>Session: {sessionnumber}", info_style)
    infos_table = Table([[student_name, class_info_block]])

    total_ects = Paragraph(f"Total de points ECTS obtenus :<br/><br/>{datastudent[5][1]}/{datastudent[4][1]}",
                           ects_style)

    title_feedback_ects = Paragraph(f"<u>Décision du jury :</u><br/><br/>", title_info_style)
    feedback_ects = Paragraph(f"{datastudent[6][1]}", info_style)
    title_commentaire_semestre = Paragraph(f"<u>Commentaire du semestre :</u><br/><br/>", title_info_style)
    commentaire_semestre = Paragraph(f"{datastudent[7][1]}", info_style)

    # Add elements to the PDF
    elements = [header_table, Paragraph('<br/><br/>'), academic_year_text, Paragraph('<br/><br/>'), infos_table,
                Paragraph('<br/><br/>'), table, Paragraph('<br/><br/>'), total_ects, Paragraph('<br/><br/>'),
                title_feedback_ects, feedback_ects, Paragraph('<br/><br/>'), title_commentaire_semestre,
                commentaire_semestre]

    # Build PDF
    pdf.build(elements)

    print(f"PDF file '{pdf_file}' generated successfully.")


async def generate_individual_pdfs(excel_file):
    # Read Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Iterate through each row and generate PDF for each student
    for index, row in df.iterrows():
        # Extract student information and grades
        grades = [[key, str(row[key]), "Acquis" if row[key] >= 10 else "Non Acquis"] for key in row.keys() if
                  key not in ['Nom', 'Prénom', 'Adresse', "Année de naissance", "Décision du Jury", "Commentaire",
                              "Crédits ECTS Attendus", "Crédits ECTS Total"]]
        datastudent = [[key, str(row[key])] for key in row.keys() if
                       key in ['Nom', 'Prénom', 'Adresse', "Année de naissance", "Décision du Jury", "Commentaire",
                               "Crédits ECTS Attendus", "Crédits ECTS Total"]]

        # Combine student information and grades
        datagrades = [["Sujet", "Moyenne", "Crédits ECTS"]] + grades

        # Modify the PDF name as you wish
        pdf_file = f"{row['Nom']}_{row['Prénom']}_bulletins.pdf"

        # Generate PDF for the student
        generate_student_pdf(datagrades, pdf_file, datastudent,
                             image_path=resource_path("logoestiam.png"))

        if send_email_bool:
            try:
                sm.send_email(datastudent, emailpassword, from_addr)  # call the send email function from sendmail.py
            except Exception as e:
                print(
                    "L'envoi des emails n'a pas eu lieu veuillez verifier"
                    " votre connexion internet ou vos identifiants \n\n"
                    "Motif de l'erreur :", e)
        else:
            print("Emails non envoyés")


async def main():
    # Create a root Tk window and hide it
    root = tk.Tk()
    root.withdraw()
    print("Selectionner le fichier excel avec les notes")
    # Open a file dialog and get the selected file path
    excel_file = filedialog.askopenfilename(
        title="Selectionner le fichier excel",
        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*"))
    )
    # Check if a file was selected
    if excel_file:
        await generate_individual_pdfs(excel_file)
    else:
        print("No file selected.")


# Run the asynchronous function within an event loop
asyncio.run(main())
