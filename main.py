import random

import psycopg2

from urllib.request import urlopen
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import cm
import pandas as pd
from PyPDF2 import PdfReader
from math import ceil
import json
from pdf2image import convert_from_path
from reportlab.lib.colors import purple
attest_path = "source/bridion.pdf"
images = convert_from_path(attest_path)

for i in range(len(images)):
    # Save pages as images in the pdf
    images[i].save('page' + str(i) + '.jpg', 'JPEG')

conn = psycopg2.connect(
    host="localhost",
    database="ANASTHASIA",
    user="postgres",
    password="postgres")

cursor = conn.cursor()


with urlopen("https://webappsa.riziv-inami.fgov.be/silverpages/Home/SearchHcw/?Form.Name=Van%20Overmeire&Form.FirstName=Henri&Form.Profession=&Form.Specialisation=&Form.ConventionState=&Form.Location=&Form.NihdiNumber=&Form.Qualification=&Form.NorthEastLat=&Form.NorthEastLng=&Form.SouthWestLat=&Form.SouthWestLng=&Form.LocationLng=&Form.LocationLat=") as response:
    print(response.headers.items())
    print(response.read())
def generate_sugammadex_attestation(attestnr, patient_name, prescriber_first_name, prescriber_last_name, riziv, datum, dosage, flacons, reason):
#creating a new canvas
    DOSAGE = dosage

    #theoretical maximum weight corresponding to number of flacons distributed.
    patient_weight = int(flacons * 200 / DOSAGE)

    #hack to conform
    succesful_weight_hack = False

    while not succesful_weight_hack:
        r = abs(random.gauss(mu=0, sigma=30))
        modified_patient_weight = patient_weight - r
        if ceil(modified_patient_weight*DOSAGE/200) == flacons:
            succesful_weight_hack = True

    patient_weight = int(modified_patient_weight)

    riziv_case = riziv[0]
    qualcode_case = riziv[1]

    canvas = Canvas(f"{prescriber_last_name}{prescriber_first_name}_{attestnr}.pdf", pagesize=A4)
    page_width, page_height = canvas._pagesize
    canvas.setFont("Courier", 13)
    canvas.drawImage('page0.jpg', x=0, y=0, width=page_width, height=page_height)
    canvas.drawString(0.8 * cm, 25 * cm, str(patient_name), charSpace=2.8)
    canvas.drawString(0.8 * cm, 24.3 * cm, str(patient_INSZ), charSpace=2.8)
    canvas.drawString(0.95 * cm, 12.55 * cm, "X", charSpace=2.8)
    canvas.setFont("Courier", 9)
    operation_date_day = datum[8:10]
    operation_date_month = datum[5:7]
    operation_date_year = datum[0:4]
    canvas.drawString(8.22 * cm, 10.33 * cm, f"{operation_date_day} {operation_date_month} {operation_date_year}", charSpace=1.7)
    canvas.setFont("Courier", 12)
    canvas.drawString(1.5 * cm, 8.95 * cm, f"{reason}", charSpace=1.7)
    amount_substance = dosage * patient_weight
    amount_bottles = ceil(amount_substance / 200)
    canvas.drawString(5 * cm, 6.10 * cm, str(patient_weight), charSpace=1.7)
    canvas.drawString(13 * cm, 6.10 * cm, str(amount_substance), charSpace=1.7)
    canvas.drawString(4.2 * cm, 5 * cm, str(amount_bottles), charSpace=1.7)


    canvas.showPage()
    canvas.setFont("Courier", 13)
    canvas.drawImage('page1.jpg', x=0, y=0, width=page_width, height=page_height)
    canvas.drawString(0.8 * cm, 16.12 * cm, f"{prescriber_last_name}", charSpace=2.8)
    canvas.drawString(0.8 * cm, 15.45 * cm, f"{prescriber_first_name}", charSpace=2.8)
    canvas.drawString(2.15 * cm, 14.7 * cm, f"{riziv_case[0:5]}", charSpace=2.8)
    canvas.drawString(4.35 * cm, 14.7 * cm, f"{riziv_case[5:]}", charSpace=2.8)
    canvas.drawString(5.40 * cm, 14.7 * cm, f"{qualcode_case}", charSpace=2.8)
    canvas.drawString(0.8 * cm, 14.0 * cm, operation_date_day, charSpace=2.8)
    canvas.drawString(2.2 * cm, 14.0 * cm, operation_date_month, charSpace=2.8)
    canvas.drawString(3.3 * cm, 14.0 * cm, operation_date_year, charSpace=2.8)
    canvas.save()


worksheet = pd.read_excel("source/februari 2022.xlsx", engine='openpyxl', sheet_name="Sheet1")
worksheet=worksheet[worksheet['DESCRIPTION'] == "BRIDION 200 MG/2 ML FLAC"]

for index, row in worksheet.iterrows():
    attestnr = row['ATTESTREFERENTIE']
    patient_name = row['NAAM']
    patient_INSZ = int(row['RIJKSREGNR'])
    aantal_flacons = int(row['AANTAL'])
    operation_date = str(row['DATUM'].date())
    prescriber = row['TOEGEWEZEN_VOORSCHRIJVER']
    prescriber_names = prescriber.split(' ')
    prescriber_first_name = prescriber_names[-1]
    prescriber_names.pop(-1)
    prescriber_last_name = " ".join(prescriber_names)

    fields = "riziv_nr, qualification_code"
    table = "anesthesie_uzgent"
    conditions = f"first_name = '{prescriber_first_name}' AND last_name = '{prescriber_last_name}'"

    query = (f"SELECT {fields} "
           f"FROM {table} "
           f"WHERE {conditions};")

    cursor.execute(query)
    riziv = cursor.fetchone()

    reason_query = """ 
    SELECT reason FROM sugammadex_attestation  
    ORDER BY random()
    LIMIT 1;
    """

    cursor.execute(reason_query)
    reason = cursor.fetchone()[0]

#   try:
    generate_sugammadex_attestation(attestnr, patient_name=patient_name, prescriber_first_name=prescriber_first_name, prescriber_last_name=prescriber_last_name, riziv=riziv, datum=operation_date, dosage=4, flacons=aantal_flacons, reason=reason)
    #except TypeError:
    #    print(f"Missing database entry for {prescriber_last_name} {prescriber_first_name}")