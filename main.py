import random

import psycopg2

from urllib.request import urlopen
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import cm
import pandas as pd
import datetime
import PyPDF2
from math import ceil
import json
from pdf2image import convert_from_path
from reportlab.lib.colors import purple
attest_path = "source/bridion.pdf"
images = convert_from_path(attest_path)

for i in range(len(images)):
    # Save pages as images in the pdf
    images[i].save('page' + str(i) + '.jpg', 'JPEG')

#conn = psycopg2.connect(
#    host="localhost",
#    database="ANASTHASIA",
#    user="postgres",
#    password="postgres")

#cursor = conn.cursor()
def generate_sugammadex_attestation(attestnr, patient_name, patient_id, prescriber_first_name, prescriber_last_name, riziv, datum, weight, dosage_schema,flacons, reason):
#creating a new canvas
    #theoretical maximum weight corresponding to number of flacons distributed.
    #hack to conform

    patient_weight = int(weight)
    riziv_case = riziv[1:8]
    qualcode_case = riziv[8:12]

    canvas = Canvas(f"{prescriber_last_name}{prescriber_first_name}_{attestnr}.pdf", pagesize=A4)
    page_width, page_height = canvas._pagesize
    canvas.setFont("Courier", 13)
    canvas.drawImage('page0.jpg', x=0, y=0, width=page_width, height=page_height)
    canvas.drawString(0.8 * cm, 25 * cm, str(patient_name), charSpace=2.8)
    canvas.drawString(0.8 * cm, 24.3 * cm, str(patient_id), charSpace=2.8)

    if dosage_schema == 4:
        canvas.drawString(0.95 * cm, 12.55 * cm, "X", charSpace=2.8)
        canvas.setFont("Courier", 9)
        operation_date_day = datum[8:10]
        operation_date_month = datum[5:7]
        operation_date_year = datum[0:4]
        canvas.drawString(8.22 * cm, 10.33 * cm, f"{operation_date_day} {operation_date_month} {operation_date_year}", charSpace=1.7)
        canvas.setFont("Courier", 12)
        canvas.drawString(1.5 * cm, 8.95 * cm, f"{reason}", charSpace=1.7)
        amount_substance = dosage_schema * patient_weight
        amount_bottles = ceil(amount_substance / 200)
        canvas.drawString(5 * cm, 6.10 * cm, str(patient_weight), charSpace=1.7)
        canvas.drawString(13 * cm, 6.10 * cm, str(amount_substance), charSpace=1.7)
        canvas.drawString(4.2 * cm, 5 * cm, str(amount_bottles), charSpace=1.7)

    elif dosage_schema == 16:

        canvas.drawString(0.95 * cm, 20.6 * cm, "X", charSpace=2.8)
        canvas.setFont("Courier", 9)
        operation_date_day = datum[8:10]
        operation_date_month = datum[5:7]
        operation_date_year = datum[0:4]
        canvas.drawString(8.22 * cm, 10.33 * cm, f"{operation_date_day} {operation_date_month} {operation_date_year}",
                          charSpace=1.7)
        canvas.setFont("Courier", 12)
        amount_substance = dosage_schema * patient_weight
        amount_bottles = ceil(amount_substance / 200)
        if amount_bottles > 7:
            amount_bottles = 7
        canvas.drawString(5 * cm, 16.25 * cm, str(patient_weight), charSpace=1.7)
        canvas.drawString(12.9 * cm, 16.25 * cm, str(amount_substance), charSpace=0.5)
        canvas.drawString(3.9 * cm, 15.3 * cm, str(amount_bottles), charSpace=1.7)
        canvas.drawString(17.3 * cm, 18.35 * cm, operation_date_day, charSpace=1.4)
        canvas.drawString(18.15 * cm, 18.35 * cm, operation_date_month, charSpace=1.4)
        canvas.drawString(19.05 * cm, 18.35 * cm, operation_date_year, charSpace=1.3)


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

worksheet = pd.read_excel("source/juni 2023.xlsx", engine='openpyxl', sheet_name="Sheet1")
worksheet=worksheet[worksheet['DESCRIPTION'] == "BRIDION 200 MG/2 ML FLAC"]

user_first_name = input("What is your first name?: ")
user_last_name = input("What is your last name?: ")
user_RIZIV = input("Type your full riziv number: ")
pdfMerge = PyPDF2.PdfMerger()

for index, row in worksheet.iterrows():
    DOSAGE = 4
    if f"{user_last_name} {user_first_name}" == row["TOEGEWEZEN_VOORSCHRIJVER"]:
        attestnr = row['ATTESTREFERENTIE']
        patient_name = row['NAAM']
        patient_INSZ = str(int(row['RIJKSREGNR']))
        print(f"len ISNZ {len(patient_INSZ)}")
        if len(patient_INSZ) == 10:
            patient_INSZ = "0" + patient_INSZ
        print(f"{patient_name} {patient_INSZ}")
        INSZ_sum = 0
        for digit in str(int(float(patient_INSZ))):
            INSZ_sum += int(digit)
        random.seed(INSZ_sum)
        patient_year_of_birth = int(patient_INSZ[0:2])
        aantal_flacons = int(row['AANTAL'])

        if patient_year_of_birth == 23:
            estimated_weight = 5
            estimated_sigma = 1

        elif 0 <= patient_year_of_birth and patient_year_of_birth <= 22:
            estimated_age = int(datetime.date.today().year - (2000 + patient_year_of_birth))

        elif patient_year_of_birth > 22:
            estimated_age = int(datetime.date.today().year - (1900 + patient_year_of_birth))

        print(f"estimated age: {estimated_age}")
        if 1 <= estimated_age <= 5:
            estimated_weight = 2 * (estimated_age + 5)
            estimated_sigma = 1
        elif 6 <= estimated_age <= 14:
            estimated_weight = 4 * estimated_age
            estimated_sigma = (estimated_age - 6) + 2
        elif 14 < estimated_age and estimated_age < 80:
            estimated_weight = 72
            estimated_sigma = 10
        elif estimated_age >= 80:
            estimated_weight = 70
            estimated_sigma = 10

        else:
            estimated_weight = 10
            estimated_sigma = 10

        succesful_weight_hack = False


        while not succesful_weight_hack:
            modified_patient_weight = random.gauss(mu=estimated_weight, sigma=estimated_sigma)
            weight_based_calculated_dosage = aantal_flacons *  200 / estimated_weight
            number_of_allowed_flacons = ceil(modified_patient_weight * DOSAGE / 200)
            if aantal_flacons > number_of_allowed_flacons and weight_based_calculated_dosage > 8:
                DOSAGE = 16
            if modified_patient_weight > 0 and ceil(modified_patient_weight * DOSAGE / 200) >= aantal_flacons:
                succesful_weight_hack = True
                patient_weight = modified_patient_weight

        #hack to conform
        print(f"{estimated_age} {patient_weight}")
        operation_date = str(row['DATUM'].date())
        prescriber = row['TOEGEWEZEN_VOORSCHRIJVER']
        prescriber_names = prescriber.split(' ')
        prescriber_first_name = prescriber_names[-1]
        prescriber_names.pop(-1)
        prescriber_last_name = " ".join(prescriber_names)

        #fields = "riziv_nr, qualification_code"
        #table = "anesthesie_uzgent"
        #conditions = f"first_name = '{prescriber_first_name}' AND last_name = '{prescriber_last_name}'"

#       query = (f"SELECT {fields} "
#               f"FROM {table} "
#               f"WHERE {conditions};")
#
#       cursor.execute(query)
#       riziv = cursor.fetchone()
        riziv = user_RIZIV
#       reason_query = """
#       SELECT reason FROM sugammadex_attestation
#       ORDER BY random()
#       LIMIT 1;
#       """
#
#       cursor.execute(reason_query)
#       reason = cursor.fetchone()[0]
        reason = "Restcurarisatie"
#       try:
        generate_sugammadex_attestation(attestnr, patient_name=patient_name, patient_id=patient_INSZ, prescriber_first_name=prescriber_first_name, prescriber_last_name=prescriber_last_name, riziv=riziv, datum=operation_date, dosage_schema=DOSAGE, weight= patient_weight, flacons = aantal_flacons, reason=reason)
        #except TypeError:
        #    print(f"Missing database entry for {prescriber_last_name} {prescriber_first_name}")
        print("GENERATED")
        pdfMerge.append(f"{prescriber_last_name}{prescriber_first_name}_{attestnr}.pdf")
    else:
        pass

pdfMerge.write(f"{prescriber_first_name}_{prescriber_last_name}_merged.pdf")


