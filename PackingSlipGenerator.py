import os
import csv
import sys
import shutil
from appJar import gui
import traceback

consignees = []
consignees_full = []

name           = 0
addressLine    = 0
city           = 0
stateProvince  = 0
postalcode     = 0
idNumber       = 0
PARSNumber     = 0

description    = 0
quantity       = 0
packagingUnit  = 0
weight         = 0
weightUnit     = 0
marksAndNumbers = 0

def processed_yet(consignee):
    new = False
    for name in consignees:
        if name == consignee:
            new = True
    return new

#Finds the index for each label eg. "Consignee Name" at index 7
def define_labels(line):
    labels = [
        "consignee_name",
        "consignee_address1",
        "consignee_city",
        "consignee_province",
        "consignee_postal_code",
        "consignee_id",
        "pars_no",
        
        "product_description",
        "quantity",
        "packaging_type",
        "net_weight",
        "weight_unit",
        "shipment_id",
        
        "Consignee Name",
        "Consignee Address1",
        "Consignee City",
        "Consignee Province",
        "Consignee Postal Code",
        "Congsinee ID No.",
        "PARS No.",
        
        "Product Description",
        "Quantity",
        "Packaging Unit",
        "Net Weight",
        "Weight Unit",
        "Shipment ID"
        ]

    global marksAndNumbers
    global name
    global addressLine
    global city
    global stateProvince
    global postalcode
    global idNumber
    global PARSNumber

    global description
    global quantity
    global packagingUnit
    global weight
    global weightUnit
    
    print(line)
    print(" ")
    for part in line:
        for label in labels:
            if part == label:
                if part == "consignee_name" or part == "Consignee Name":
                    name = line.index(part)
                if part == "consignee_address1" or part == "Consignee Address1":
                    addressLine = line.index(part)
                if part == "consignee_city" or part == "Consignee City":
                    city = line.index(part)
                if part == "consignee_province" or part == "Consignee Province":
                    stateProvince = line.index(part)
                if part == "consignee_postal_code" or part == "Consignee Postal Code":
                    postalcode = line.index(part)
                if part == "consignee_id" or part == "Congsinee ID No.":
                    idNumber = line.index(part)
                if part == "pars_no" or part == "PARS No.":
                    PARSNumber = line.index(part)
                
                if part == "product_description" or part == "Product Description":
                    description = line.index(part)
                if part == "quantity" or part == "Quantity":
                    quantity = line.index(part)
                if part == "packaging_type" or part == "Packaging Unit":
                    packagingUnit = line.index(part) # Unused?
                if part == "net_weight" or part == "Net Weight":
                    weight = line.index(part)
                if part == "weight_unit" or part == "Weight Unit":
                    weightUnit = line.index(part) #Unused. Always "LBR"
                if part == "shipment_id" or part == "Shipment ID":
                    marksAndNumbers = line.index(part)
   
def create_slips():
    try:
        arg = app.getEntry("fileEntry")
        date = app.getDatePicker("datePicker")
        
        print(arg)
        print(date)
        
        if not arg.endswith(".csv"):
            print('Error: Please use a .csv file')
            sys.exit()
        if not arg == "":
            lines = list(csv.reader(open(arg)))
            slips_folder = os.getcwd() + "\\slips\\"
            if os.path.exists(slips_folder):
                shutil.rmtree(slips_folder)
            os.mkdir(slips_folder)
            
            define_labels(lines[0])
            del lines[0] #First column is label
            
            #Find all the consignees
            for line in lines:
                if not processed_yet(line[name]):
                    consignees.append(line[name])
                    consignees_full.append(line)
            
            #For each Consignee, and for each shipment per consignee, print the .json stuff
            i = 0
            for consignee in consignees:
                index = consignees.index(consignee)
                
                # PARS handling
                PARS = app.getEntry("PARSNumberBox")
                
                # Per consignee, creater a packign slip for them
                slip = open(slips_folder + consignee.replace('/', '') + ".rtf", "w+")
                slip_text = '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033\\deflangfe1033{\\fonttbl{\\f0\\fnil\\fcharset0 Lucida Console;}}\\n{\\*\\generator Riched20 10.0.10586}{\\*\\mmathPr\\mdispDef1\\mwrapIndent1440 }\\viewkind4\\uc1 \\n\\pard\\nowidctlpar\\sa200\\sl276\\slmult1\\qr\\b\\f0\\fs40\\lang9 PACKING SLIP\\b0\\fs22\\par\\n ' + str(date) + '\\par CID: ' + consignees_full[index][idNumber] + '\\b0\\fs22\\par\\n\\pard\\nowidctlpar\\sa200\\sl276\\slmult1 DeFranco Hardware\\line 3105 Pine Ave, Niagara Falls, NY, 14301\\line 1-877-863-7447\\par\\n SHIP TO:\\line ' + str(consignee) + '\\line ' + consignees_full[index][addressLine] +' '+ consignees_full[index][addressLine+1] + '\\line ' + consignees_full[index][addressLine+2] + ', ' + consignees_full[index][addressLine+3] + '\\line Canada\\par\\n ORDER DATE\\tab\\tab PURCHASE ORDER\\line ' + str(date) + '\\tab\\tab ' + PARS + '\\par\\n ORDER Q#\\tab SHIP Q#\\tab ITEM\\line '
                
                total = 0
                total2 = 0
                for line in lines:
                    if line[name] == consignee:
                        # Per commodity, add it to the consignee's shitpment
                        slip_text += line[quantity] + '\\tab\\tab 1\\tab\\tab ' + line[description] + '\\line '
                        total += int(line[quantity])
                        total2 += 1
                # Finish the loading slip
                slip_text += '---------------------------\\line ' + str(total) + '\\tab\\tab ' + str(total2) + '\\tab\\tab TOTAL\\par\\n}' 
                slip.write(slip_text)
                slip.close()
                
                i += 1
    except:
        traceback.print_exc()

def fetch_pars():
    print("Fetching Pars")
    try:
        arg = app.getEntry("fileEntry")
        if not arg == "":
            lines = list(csv.reader(open(arg)))
            print(lines[1])
            for i in range(len(lines[0])):
                if lines[0][i] == "pars_no" or lines[0][i] == "PARS No.":
                    app.setEntry("PARSNumberBox", lines[1][i]) # Fetches the first PARS number from the file and assumes all are the same
    except:
        traceback.print_exc()
#BEGIN

app = gui()
app.addLabel("testLabel", "Packing Slip Generator")
app.addFileEntry("fileEntry")
app.setEntryChangeFunction("fileEntry", fetch_pars)

app.startLabelFrame("PARS Number Adjuster")
app.addEntry("PARSNumberBox")
app.stopLabelFrame()

app.startLabelFrame("Packing Slip Generator")
app.addDatePicker("datePicker")
app.setDatePicker("datePicker")
app.setDatePickerRange("datePicker", 2019, 2037)
app.addButton("Generate", create_slips)
app.stopLabelFrame()
app.go()
