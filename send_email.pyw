import win32com.client as win32
import os

"""
---------------------------------------------------------------------------------
Semi-automate methode to find all people and attach thier PDF price lists to them
---------------------------------------------------------------------------------
"""
def znajdz_adresatow():
    """
    Returns
    -------
    adresaci : list_of_strings
        lits of people that need the attachments from After folder need to be sent.
    """
    directory = r'I:\\ICS\\Audit & Quality\\Projects\\PI MasterFile\\Customer Pricelist generation tool\\Spencer\\After\\'
    adresaci = []
    i=0
    for filename in os.listdir(directory):
        print (i)
        osoba = ""
        for letter in filename:
            if letter != "-" and letter !="(":
                osoba = osoba + letter
            else:
                print (osoba)
                if osoba not in adresaci:
                    adresaci.append(osoba)
                break
        i+=1
    return adresaci
   

def wyslij_maile_z_lity_adresatow(adresat_main):
    """
    Returns
    -------
    Displays every mail with attachments
    """
    temp_title = "PDF Price List"
    Tempstore="G:\\Python projects\\My\\PDF mail sender\\template.oft"
    temp_from = "PPGIndustrialCoatings@ppg.com"

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItemFromTemplate(Tempstore)
    mail.Subject = temp_title
    mail.sender = temp_from	
    directory = r'I:\\ICS\\Audit & Quality\\Projects\\PI MasterFile\\Customer Pricelist generation tool\\Spencer\\After\\'
    mail.To = adresat_main
    i = 1 
    for filename in os.listdir(directory):
        adres = ""
        path_of_file = str(directory) + str(filename) 
        for letter in filename:
            if letter != "(" and letter!= "-":
                #elif letter != "(" and letter != ")":
                adres = adres + letter
            else:
                print (adres)
                if adres == adresat_main:
                    print (str(i) + ":     " + str(adres) + "--- " + str(path_of_file))
                    mail.Attachments.Add(path_of_file)
                break
        i+=1
    mail.Display(True)


adresaci = znajdz_adresatow()
print (adresaci)
for adresat in adresaci:
    wyslij_maile_z_lity_adresatow(adresat)
    
    
"""
---------------------------------------------------------
Manual method to send lots of attachments do temp_adresat (comment loop above to work)
---------------------------------------------------------
"""

"""
temp_adresat = "Zaccone, Andrea "
temp_title = "PDF Price List"
Tempstore="G:\\Python projects\\My\\PDF mail sender\\template.oft"
temp_from = "PPGIndustrialCoatings@ppg.com"

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItemFromTemplate(Tempstore)
mail.Subject = temp_title
mail.To = temp_adresat
mail.sender = temp_from	
	


directory = r'I:\\ICS\\Audit & Quality\\Projects\\PI MasterFile\\Customer Pricelist generation tool\\Spencer\\After\\'
i = 1 
for filename in os.listdir(directory):
    adresat = ""
    path_of_file = str(directory) + str(filename) 
    for letter in filename:
        if letter != "(":
            #elif letter != "(" and letter != ")":
            adresat = adresat + letter
        else:
            print (adresat)
            if adresat == temp_adresat:
                print (str(i) + ":     " + str(adresat) + "--- " + str(path_of_file))
                mail.Attachments.Add(path_of_file)
            break
    #send_email("PDF Price List","PPGIndustrialCoatings@ppg.com",adresat,path_of_file,Tempstore)
    
    i+=1

mail.Display(True)
"""