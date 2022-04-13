'''
Dynamically generate word documents using data from a CSV - with 1 template file.
'''
# Use with : template.docx in same dir
# pip install python-docx
# pip install docxtpl <- Better for making new files from a template
import random
import time
import csv
import pandas as pd
from docxtpl import DocxTemplate

# Source CSV - column names that must match the *** that are {{***}} inside "template.docx"
csvfn = "data.csv"

def mkw(n,fname):
    tpl = DocxTemplate("test.docx") # In same directory
    df = pd.read_csv(csvfn)
    df_to_doct = df.to_dict() # dataframe -> dict for the template rendering
    x = df.to_dict(orient='records')
    context = x
    tpl.render(context[n])
    tpl.save("{}.docx".format(fname))

    # Wait a random time - increase to (60,180) for real production run.
    wait = time.sleep(random.randint(1,2))
    
df2 = len(pd.read_csv(csvfn))

print ("There will be ", df2, "files")

for i in range(0,df2):
    df = pd.read_csv(csvfn)
    columnsData = df.loc[ : , 'Name' ]
    fname = columnsData[i]
    print("Making file: ", fname)
    mkw(i,fname)

print("All Files created")
