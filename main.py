import sys

from PIL import Image, ImageFont, ImageDraw  
import xlrd

#fontstyle


#Help function
def help():
    print("""
                This is Python3 script that automatically creates certificates for the names given.

                1. Replace(Keep your filename same as one in folder) your Excel sheet that contains names of the participants with data/mydata.xlsx
                2. Name your certificate image file as template.jpg and replace with data/template.jpg
                3. You can choose a font among 4 available fonts
                4. Please keep ready the (pixels X pixels) at which the text is to be installed.
                    a. You can find the dimensions in the paint application
                5. To test with one certificate use this command "python3 main.py test"

            """)

try:
    if sys.argv[1] == "help":
        help()
        sys.exit()
except:
    pass



#Choose your font
available_fonts = ["Roboto-Medium", "Roboto-Black", "Louis George Cafe Light", "Baby Bunny"]

for i in range(len(available_fonts)):
    print(f"\n{i+1}. {available_fonts[i]} - To Choose this font press {i+1} and Enter")


font_to_be_used = available_fonts[int(input("\n")) - 1]

print(f"\nYour font is set to {font_to_be_used}")
  
#Location of the Data file - Excel sheet
excelFileLocation = ("data/mydata.xlsx")
  
# To open Workbook 
wb = xlrd.open_workbook(excelFileLocation)

sheet = wb.sheet_by_index(0) 

print(f"\nThere are {sheet.nrows} names in the Excel Sheet provided.")

try:
    nameCount = 1 if sys.argv[1] == "test" else sheet.nrows
except:
    nameCount = sheet.nrows

print("\nEnter the (pixel X pixel) where the text needs to be inserted. Note: Enter in dimensions by giving space one after the other")
textPosition = tuple(int(i) for i in (input().split(" ")))

for i in range(nameCount):
    nameOfTheParticipant = str(sheet.cell_value(i, 0)).capitalize()
    # nameInTheCertificate = nameOfTheParticipant.center(0, " ")
    nameInTheCertificate = nameOfTheParticipant
    
    image = Image.open(r"data/template.jpg")
      
    drawImage = ImageDraw.Draw(image)  
      
    available_fonts = ["Roboto-Medium.ttf", "Roboto-Black.ttf", "Louis George Cafe Light.ttf", "Baby Bunny.ttf"]
    font = ImageFont.truetype(r'data/fonts/{}.ttf'.format(font_to_be_used), 140)
      

    # drawing text size
    drawImage.text(textPosition, nameInTheCertificate, fill ="white", align ="center")
      
    image.save(f"certificates/{nameOfTheParticipant}.jpg")

print("""
        ************************************
        Your Certificates are ready!
        Thanks for Using!
        ************************************
    """)



