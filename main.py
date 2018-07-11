# Keldon Fischer
# July 10, 2018
#
# linkedin: https://www.linkedin.com/in/keldon-fischer-4437b4116/
# github: https://github.com/graetkel


from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
import xlsxwriter
import os, sys
from openpyxl import load_workbook
import re
import requests
from bs4 import BeautifulSoup


fileLocation = ""

# companyTotalResults:
# This fuction takes in a string of a company name. Then it goes to google and
# searches "company_name" "prison labor", after searching it grabs the number of
# total results. This method then returns the total results.
def companyTotalResults(theCompanyName):
    companyName = theCompanyName

    companyNamePlus = re.sub(" ", "+", companyName)
    companyNamePlus = re.sub("\n", "", companyNamePlus)
    companyNamePlus = companyNamePlus.lower()

    myLink = "https://www.google.com/search?q=" + "\"" + companyNamePlus + "\"" + "+\"prison+labor\""
    page = requests.get(myLink)

    soup = BeautifulSoup(page.content, 'html.parser')

    container = soup.find(id="resultStats")
    tempCount = "" + str(container)

    tempCount = tempCount.replace("<div class=\"sd\" id=\"resultStats\">About ", "")
    tempCount = tempCount.replace(" results</div>", "")
    tempCount = tempCount.replace(",", "")

    result = 0
    try:
        result = int(tempCount)
    except:
        result = 0

    return result

# loadFile:
# This fuction load the adress of the excel file that the program will be using.
# Then the program adds all of the companies to a list and finds the total
# results for each. Finally it saves what it learns to a new excel file.
def loadFile():
    fileLocation = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),))
    ################################################
    # If file exist
    ################################################
    if (fileLocation != ""):
        # run search
        companyList = []
        companyCountList = []
        ignoreText1 = "FiscalYear 2018"
        ignoreText2 = "Supplier Name (supplier number)"

        wb = load_workbook(filename=fileLocation, read_only=True)
        ws = wb['Sheet1']

        for row in ws.rows:
            for cell in row:
                cv = cell.value # I call cell.value a lot so I gave it a nickname
                if (cv is not None and cv != ignoreText1 and cv != ignoreText2):
                    cv = re.sub("\([^>]+\)", "", cv)
                    companyList.append(cv)

        for index in range(len(companyList)):
            companyCountList.append(companyTotalResults(companyList[index]))

        ################################################
        # Create an new Excel file and add a worksheet.
        ################################################
        workbook = xlsxwriter.Workbook('Company Prison Labor Search.xlsx')
        worksheet = workbook.add_worksheet()

        for i in range(len(companyList)):
            biggestCount = 0
            for k in range(len(companyList)):
                if (int(companyCountList[k]) > biggestCount):
                    biggestCount = companyCountList[k]

            tempIndex = companyCountList.index(biggestCount)

            worksheet.write(i, 0, companyList[tempIndex])
            worksheet.write(i, 1, str(companyCountList[tempIndex]))

            companyList.remove(companyList[tempIndex])
            companyCountList.remove(companyCountList[tempIndex])

        workbook.close()

        ################################################
        # Try to open new excel file for the user
        ################################################
        try:
            os.system('open Company\ Prison\ Labor\ Search.xlsx')
        except:
            fd = os.open( "Company Prison Labor Search.xlsx")

    ################################################
    # If file doesn't exist
    ################################################
    else:
        msg = messagebox.showinfo( "Error", "File did not load properly")


# Create the window
root = Tk()
v = IntVar()

# Setting app defaults
root.title("Company Search Tool")
root.geometry("200x50")
app = Frame(root)
app.grid()

# Adding all items
label = Label(app, text = "Company Search:")
selectButton = Button(app, text = "Select File", command = loadFile)

label.grid()
selectButton.grid()


root.mainloop()

################################################
# End of File
################################################
################################################
