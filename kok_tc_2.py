from re import T
import pandas as pd
import PyPDF2
import pprint
import glob
import os
pp = pprint.PrettyPrinter()

#kok needs to fix Equip Repaired (row 1)
col_names = ["Employee Name", "Employee Number", "Job NumberRow", "CostPhase CodeRow", "Class Line ", "earn", "Today Date", "Regularhrs", "Overtimehrs",  
   "Equip Repaired ", "Equip Used ", "Repair Code Line ", "Hours Used", "Super", "Equip Repaired Meter Reading", "Problem Log", "Safety Talk", "Description of WorkRow"]

has_rows = ["Job NumberRow", "CostPhase CodeRow", "Repair Code Line ", "Equip Repaired ", "Class Line ", "earn", 
"Regularhrs", "Overtimehrs", "Equip Used ", "Hours Used", "Super", "Equip Repaired Meter Reading", "Problem Log", "Description of WorkRow"]

#declare empty df
df = pd.DataFrame(columns=col_names)


#function to convert pdf fields to rows of data
def create_rows(fields):
    print("Generating rows...")
    rows = []
    pp.pprint(fields)
    
    count = 1
    #loop through every row of the time card
    while count < 8:
        #create array to store current row data
        data = []
        #loop through the columns
        for i in col_names:
            
            #if the column has multuple rows            
            if i in has_rows:
                #get field name based on the current row count and append data to array                
                field_name = f'{i}{count}'
                #setting the value for earn based on client request - 2 = OT 1 = reg
                if i == 'earn' or i == 'jcco' or i == 'eqco' or i == 'Repair Code Line ':
                    data.append(fields[field_name].value.split('-')[0])
                else:
                    data.append(fields[field_name].value)            
            else: 
                #else just grab the field data
                data.append(fields[i].value)
        #combine dataframes
        zipped = zip(col_names, data)
        #add new row to rows
        rows.append(dict(zipped))
        #increment count to go to next row
        count += 1
    return rows

def create_csv(dateframe):
    print("Creating import...")
    writer = pd.ExcelWriter('import.xlsx')
    df.to_excel(writer, sheet_name='time card', index=False, na_rep='')
    writer.save()
    print("Process complete")

#create a empty var to store pdfs
pdf_list = []
#open pdf directory
os.chdir("mech_cards")
#grab all pdfs and add them to the list
for file in glob.glob("*.pdf"):
    pdf_list.append(file)
card_count = len(pdf_list)
print(f'Found {card_count} time cards to process')
#loop through pdfs
for card in pdf_list:
    print("Getting data from card...")
    #open the current card and read first (only) page
    pdf_file = open(card, 'rb')
    reader = PyPDF2.PdfFileReader(pdf_file)
    page = reader.getPage(0)
    fields = reader.getFields()
    #print(fields)
    #call method to create rows from pdf data 
    card_rows = create_rows(fields)

    #add all rows to the dataframe    
    df = df.append(card_rows, True)

create_csv(df)