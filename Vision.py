from google.cloud import vision_v1p3beta1 as vision
client = vision.ImageAnnotatorClient()
import io
import os
import xlwt
import pandas
from openpyxl import load_workbook
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
import xlsxwriter
import shutil
profile_count1 = 1
for filename in os.listdir(r'C:\Users\Sauhard\Downloads\MyFolder\Not Done'):
    with io.open(os.path.join(r'C:\Users\Sauhard\Downloads\MyFolder\Not Done', filename), 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    # Language hint codes for handwritten OCR:
    # en-t-i0-handwrit, mul-Latn-t-i0-handwrit
    # Note: Use only one language hint code per request for handwritten OCR.
    image_context = vision.ImageContext(
        language_hints=['en-t-i0-handwrit'])

    response = client.document_text_detection(image=image,
                                              image_context=image_context)
    print(response.full_text_annotation.text)
    w = str(response.full_text_annotation.text).split('Name')
    advisor = w[0].split('\n')[0]
    temp_string = w[1].split('Address')
    name  = temp_string[0]
    print(name) 
    temp_string = temp_string[1].split('City')
    address = temp_string[0]
    temp_string = temp_string[1].split('Postal Code')
    city = temp_string[0]
    temp_string = temp_string[1].split('Telephone 1')
    postal_code = temp_string[0]
    print(postal_code)
    temp_string = temp_string[1].split('Telephone 2')
    tel1 = temp_string[0]
    print(tel1)
    temp_string = temp_string[1].split('Email')
    tel2 = temp_string[0]
    print(tel2)
    temp_string = temp_string[1].split('Carrier')
    email = temp_string[0]
    print(email)
    temp_string = temp_string[1].split('Policy #')
    carrier = temp_string[0]
    print(carrier)
    temp_string = temp_string[1].split('Type')
    policy = temp_string[0]
    print(policy)
    temp_string = temp_string[1].split('Issue Date')
    type1 = temp_string[0]
    print(type1)
    temp_string = temp_string[1].split('Sex')
    issue_date = temp_string[0]
    print(issue_date)
    temp_string = temp_string[1].split('DOB')
    sex = temp_string[0]
    print(sex)
    temp_string = temp_string[1].split('Smoker')
    dob = temp_string[0]
    print(dob)
    temp_string = temp_string[1].split("""No's""")
    smoker = temp_string[0]
    print(smoker)


    
    book_ro = open_workbook('Client.xls')
    book = copy(book_ro)
    sheet1 = book.get_sheet(0)  
    sheet1.write(profile_count1, 1, name)
    sheet1.write(profile_count1, 2, sex)
    sheet1.write(profile_count1, 3, dob)
    sheet1.write(profile_count1, 4, address)
    sheet1.write(profile_count1, 5, postal_code)
    sheet1.write(profile_count1, 6, tel1)
    sheet1.write(profile_count1, 7, tel2)
    sheet1.write(profile_count1, 8, email)
    sheet1.write(profile_count1, 9, advisor)
    sheet1.write(profile_count1, 10, carrier)
    sheet1.write(profile_count1, 11, policy)
    sheet1.write(profile_count1, 12, type1)
    sheet1.write(profile_count1, 13, issue_date)
    sheet1.write(profile_count1, 14, smoker)
    book.save("Client.xls")
     
    profile_count1 = profile_count1+1
    
    name=""
    sex=""
    dob=""
    address= ""
    postal_code=""
    tel1=""
    tel2=""
    email=""
    advisor=""
    carrier=""
    policy=""
    type1=""
    issue_date=""
    smoker=""

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))
    else:
        pass
        path = r'C:\Users\Sauhard\Downloads\MyFolder\Not Done'
        finDir = r'C:\Users\Sauhard\Downloads\MyFolder\Done'
        shutil.move(os.path.join(path, filename), finDir) 