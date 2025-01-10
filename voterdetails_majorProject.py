from PIL.ImageOps import grayscale
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
import numpy as np
import re
# from openpyxl import load_workbook

# Convert PDF to images
pdf_file = 'E:\\PYTHON2024\\PythonProjects\\MAJOR_PROJECT\\ImagestoTexts\\PS_NO_3_45.pdf'
images = convert_from_path(pdf_file)
data = []
# Loop through each image (page)
sno_text, voter_name, relative_name, house_no, age, gender, code, code_text,relation = '','','','','','','','',''
for image in images:
    # Convert image to grayscale
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    #
    # # Apply thresholding to segment out the boxes
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # Find contours of the boxes
    contours, _ = cv2.findContours(thresh,cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    sno_list =[]
    # Loop through each contour (box)
    for contour in contours:
        # Get the bounding rectangle of the contour
        x, y, w, h = cv2.boundingRect(contour)

        # Check if the contour is a box (aspect ratio between 2 and 10)
        aspect_ratio = float(w) / h
        if aspect_ratio > 2 and aspect_ratio < 10:
            # Extract the text from the box
            roi = gray[y:y + h, x:x + w]
            left_width = int(w/2)
            right_width = w-left_width
            if left_width>0 and y<gray.shape[0] and x<gray.shape[1]:
                sno_text = pytesseract.image_to_string(roi[:,:left_width], lang='eng', config=r'--psm 6')
            if right_width > 0 and y < gray.shape[0] and x+left_width < gray.shape[1]:
            # Extract the code from the second box
                code_text = pytesseract.image_to_string(roi[:,left_width:], lang='eng', config=r'--psm 6')
            sno_list = sno_text.split('\n')
            for item in sno_list:
                if item.startswith("Name :") or item.startswith("Name +") or item.startswith("Name ?"):
                    voter_name= re.split(r"Name :|Name +|Name ?|Name !|Name =|Name \++|Name +:",item)[-1]
                    voter_name = re.sub(r'[^a-zA-Z\s]','',voter_name)
            for item in sno_list:
                if item.startswith("Husbands Name:") or item.startswith("Fathers Name:") or item.startswith("Mothers Name:") or item.startswith("Others:"):
                    relative_name = re.split("Husbands Name:|Fathers Name:|Mothers Name:|Others:",item)[-1]
                    relative_name = re.sub(r'[^a-zA-Z\s]', '', relative_name)
            for item in sno_list:
                if item.startswith("House Number :"):
                    house_no= item.split("House Number :")[-1]
            for item in sno_list:
                if item.startswith("Age :"):
                    age_list = item.split("Age :")[-1]
                    age_list = re.findall(r"\d+", age_list)
                    age = ''.join(age_list)
            for item in sno_list:
                if item.startswith("Age :"):
                    gender_list = item.split("Age :")[-1]
                    gender_list = re.findall(r"\b(Female|Male)\b", gender_list)
                    gender = ''.join(gender_list)
                    if gender == "Male":
                        relation = 'Father'
                    else:
                        relation = ''
            code_list = code_text.split('\n')
            for item in code_list:
                if item.isalnum():
                    # code = code_list[0].strip(r':,]*\\')
                    code = re.sub(r'[^a-zA-Z0-9]','',code_list[0])

            data.append({"EPIC_ID":code ,"Name": voter_name,"Relative Name":relative_name, "Relation": relation, "House Number":house_no,"Age":age ,"Gender":gender})

df = pd.DataFrame(data)
# df.read_excel('Major_project.xlsx')
df = df.drop_duplicates()
df.to_excel('Major_project2.xlsx')





