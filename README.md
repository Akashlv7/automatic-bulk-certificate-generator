# automatic-bulk-certificate-generator

This is Python3 script that automatically creates certificates for the names given.
    1. Replace(Keep your filename same as one in folder) your Excel sheet that contains names of the participants with data/mydata.xlsx
    2. Name your certificate image file as template.jpg and replace with data/template.jpg
    3. You can choose a font among 4 available fonts
    4. Please keep ready the (pixels X pixels) at which the text is to be installed.
        a. You can find the dimensions in the paint application 5. To test with one certificate use this command "python3 main.py test"


**#Executing the Script
**
1. Clone/Download the repo
2. Navigate into "automatic-bulk-certificate-generator" folder
3. Create virtual environment by using - "python3 -m venv <environment_name>"
4. Activate the enviroment using the source command - "source <environment_name>/bin/activate"
5. Install the dependencies using - "pip install -r requirements.txt"
6. Now you are good to go!
7. Make sure you placed the Excel and Image file in proper folder as explained above.
8. Use the following command to run the script - "python3 main.py"
9. To test the script with one certicate use the following command - "python3 main.py test"
10. To See help use - "python3 main.py help"
