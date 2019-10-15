import os
import csv
import re
import sys
import shutil

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
        "shipment_id"
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
                if part == "consignee_name":
                    name = line.index(part)
                if part == "consignee_address1":
                    addressLine = line.index(part)
                if part == "consignee_city":
                    city = line.index(part)
                if part == "consignee_province":
                    stateProvince = line.index(part)
                if part == "consignee_postal_code":
                    postalcode = line.index(part)
                if part == "consignee_id":
                    idNumber = line.index(part)
                if part == "pars_no":
                    PARSNumber = line.index(part)
                
                if part == "product_description":
                    description = line.index(part)
                if part == "quantity":
                    quantity = line.index(part)
                if part == "packaging_type":
                    packagingUnit = line.index(part) # Unused?
                if part == "net_weight":
                    weight = line.index(part)
                if part == "weight_unit":
                    weightUnit = line.index(part) #Unused. Always "LBR"
                if part == "shipment_id":
                    marksAndNumbers = line.index(part)
         
#BEGIN

for arg in sys.argv[1:]:
    if not arg.endswith('.csv'): # Exit if a zip file wan't specified as an argument
        print('Error: Please use a .csv file')
        sys.exit()

lines = list(csv.reader(open(arg)))
slips_folder = os.getcwd() + "\\slips\\"
if os.path.exists(slips_folder):
    shutil.rmtree(slips_folder)
os.mkdir(slips_folder)

define_labels(lines[0])
del lines[0] #First column is label

#ask for date
date_pattern = re.compile(r'\d{2}-\d{2}-\d{2}')
done = False
while done == False:
    date = input("Please enter the arrival date in the following format: YY-MM-DD\n")
    if date_pattern.match(date):
        done = True
        #print(shipment)
    else:
        print("Date entered in improper format")

#Find all the consignees
for line in lines:
    new = ''
    if not processed_yet(line[name]):
        consignees.append(line[name])
        consignees_full.append(line)

#For each Consignee, and for each shipment per consignee, print the .json stuff
i = 0
for consignee in consignees:
    index = consignees.index(consignee)
    # Per consignee, creater a packign slip for them
    slip = open(slips_folder + consignee.replace('/', '') + ".rtf", "w+")
    slip_text = '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033\\deflangfe1033{\\fonttbl{\\f0\\fnil\\fcharset0 Lucida Console;}}\\n{\\*\\generator Riched20 10.0.10586}{\\*\\mmathPr\\mdispDef1\\mwrapIndent1440 }\\viewkind4\\uc1 \\n\\pard\\nowidctlpar\\sa200\\sl276\\slmult1\\qr\\b\\f0\\fs40\\lang9 PACKING SLIP\\b0\\fs22\\par\\n 20' + str(date) + '\\par CID: ' + consignees_full[index][idNumber] + '\\b0\\fs22\\par\\n\\pard\\nowidctlpar\\sa200\\sl276\\slmult1 DeFranco Hardware\\line 3105 Pine Ave, Niagara Falls, NY, 14301\\line 1-877-863-7447\\par\\n SHIP TO:\\line ' + str(consignee) + '\\line ' + consignees_full[index][addressLine] +' '+ consignees_full[index][addressLine+1] + '\\line ' + consignees_full[index][addressLine+2] + ', ' + consignees_full[index][addressLine+3] + '\\line Canada\\par\\n ORDER DATE\\tab\\tab PURCHASE ORDER\\line 20' + str(date) + '\\tab\\tab ' + consignees_full[index][PARSNumber] + '\\par\\n ORDER Q#\\tab SHIP Q#\\tab ITEM\\line '
    
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

#print(out_text)
input("Press Enter to exit")
