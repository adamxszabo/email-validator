import re, pandas, glob

#function to check email format validity with regex
def validateEmail(email):
    if (re.fullmatch(regex, email)):
        return("Valid Email Format")
    else:
        return("Invalid Email Format")

regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

#put excel file names in a list
fileNames = glob.glob("*.xlsx")

#edit this to add more column names to check for emails
emailColumns = ["Email Address", "E-mail Address", "Email Addresses", "E-mail Addresses", "Email", "E-mail", "Emails", "E-mails"]

#loop through all excel files found
for fileName in fileNames:
    emails = []
    validity = []
    excelFile = pandas.read_excel(fileName)
    
    #find the column that contains email addresses
    for emailColumn in emailColumns:
        try:
            emails.append(excelFile[emailColumn].values)
            print("Column found: " + emailColumn + " in file: " + fileName)
        except:
            print("No column named: " + emailColumn + " in file: " + fileName)
    
    emailsOnly = emails[0]

    #call regex function for each email address found
    for email in emailsOnly:
        validity.append(validateEmail(email))

    excelFile["Validity"] = validity
    excelFile.to_excel(fileName, index=False)
