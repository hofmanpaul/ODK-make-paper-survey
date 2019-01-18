#This script takes all ODK formatted excel files in the current directory and turns them into paper surveys
#In no way is this supposed to lead to a perfect file. Extensive tweaking will be necessary, as the structure of ODK allows for completely different surveywriting.
#Python 2 only because of xlrd (but might fix that)

from docx import Document
from xlrd import open_workbook
import re
import os

# Initialise Files
listfiles=os.listdir(os.getcwd())
excelfiles=[]
for stuff in listfiles:
    if stuff.endswith('.xls') or stuff.endswith('.xlsx') and ('~$') not in stuff:
        excelfiles=excelfiles+[stuff]
print "This file will work on:" 
for excelfile in excelfiles:
    print excelfile
for excelfile in excelfiles:
    inbook= open_workbook(excelfile)
    survey= inbook.sheet_by_name('survey')
    choices=inbook.sheet_by_name('choices')
    settings=inbook.sheet_by_name('settings')
    outdoc=Document()
    for col in range(settings.ncols):
        if settings.cell_value(0,col)=="default_language":
            language = settings.cell_value(1,col).lower()     

    #Make choices dictionary
    for col in range(choices.ncols):
        if choices.cell_value(0,col)=="list_name":
            listname_col=col
        if choices.cell_value(0,col).lower()=="label::"+language:
            label_col=col
        if choices.cell_value(0,col)=="name":
            num_col=col

    choicesdict={}
    choicelist=[]
    listname=choices.cell_value(1,listname_col).strip()
    for row in range(1,choices.nrows):
        if choices.cell_value(row,listname_col)=="":
            continue
        if listname==choices.cell_value(row,listname_col).strip():
            label= choices.cell_value(row,label_col)
            num= choices.cell_value(row,num_col)
            strnum=str(num).rstrip('0').rstrip('.')
            choicelist=choicelist + [strnum+'. '+ label]
            if row==choices.nrows-1: #for the last row
                choicesdict[listname]=choicelist
        else: #write the choicelist to the dictionary, then continue with the next choices
            choicesdict[listname]=choicelist
            choicelist=[]
        
            listname= choices.cell_value(row,listname_col).strip()
            label= choices.cell_value(row,label_col)
            num= choices.cell_value(row,num_col)
            strnum=str(num).rstrip('0').rstrip('.')
            choicelist=choicelist + [strnum+'. '+ label]

    #Write Initial stuff
    outdoc.add_heading(settings.cell_value(1,0))
    outdoc.add_paragraph("This is an automatically generated paper survey based on an ODK excel file. Each question consists of four parts: 1. the variable name (bold). This also includes a question number (which is not used) and matches the name of the corresponding variable in the dataset. 2. A label. 3. a hint (italic). 4. The answer option(s). Before the hint there can be some conditions that dictate when this question is asked. The answer options are: ellipses (...) for text answers, white space between hyphens (-    -) for numbers, o if interviewees have to choose one and u if interviewees have to select multiple.")

    #find relevant columns
    for col in range(survey.ncols):
        if survey.cell_value(0,col)=="type":
            type_col=col
        if survey.cell_value(0,col).lower()=="label::"+language:
            label_col=col
        if survey.cell_value(0,col).lower()=="hint::"+language:
            hint_col=col
        if survey.cell_value(0,col)=="name":
            name_col=col
        if survey.cell_value(0,col)=="relevant":
           relevant_col=col 

    #Find relevant rows
    variablerows=[]
    for row in range(1, survey.nrows):
        if survey.cell_value(row,type_col)=="begin group" or survey.cell_value(row,type_col)=="end group" or survey.cell_value(row,type_col)=="" or survey.cell_value(row,type_col)=="start" or survey.cell_value(row,type_col)=="end" or survey.cell_value(row,type_col)=="deviceid" or survey.cell_value(row,type_col)=="end repeat":
            continue
        else:
            variablerows=variablerows+[row]
    
    # Make list of lists of the question properties and write
    variablelist=[]
    questionnumber=1
    for row in variablerows:
        print "Now writing row " + str(row+1)
        #assign names
        if ' ' in survey.cell_value(row,type_col).strip():
            totaltype=survey.cell_value(row,type_col)
            totaltype=totaltype.strip()
            typelist=totaltype.split(' ')
            qtype=typelist[0]
            choices=typelist[1]
        else:
            qtype=survey.cell_value(row,type_col)
            choices=''
        name=survey.cell_value(row,name_col)
        label=survey.cell_value(row,label_col)
        relevant=survey.cell_value(row,relevant_col)
        hint=survey.cell_value(row,hint_col)
        variablelist.append([questionnumber, qtype, choices, name, label, hint, relevant])
        #Write!
        #Name
        namewrite=outdoc.add_paragraph()
        namewrite.add_run(str(questionnumber)+ ' ' + name).bold=True
        
        #Label
        labelwrite=outdoc.add_paragraph()
        labelwrite.add_run(label)
    
        #Relevant
        if not relevant=="":
            #Keep going until all ODK references are replaced with human-readeable references
            while '$' in relevant:
                #find the first relevant variable
                match=re.search('\{', relevant)
                first=match.start()+1
                match=re.search('}', relevant)
                last=match.end()-1
                variable=relevant[first:last]
                #Remove the curly braces, and add the question number
                Found=False
                i=0
                while not Found:
                    if variablelist[i][3]==variable:
                        relevant=relevant.replace('${' + variable + '}', 'Q' + str(variablelist[i][0]) + ' ' + variablelist[i][3])
                        Found=True
                    i+=1
                #Fix relevant syntax
                while 'selected' in relevant:
                    print relevant
                    match=re.search('selected\(.+?\)', relevant)
                    relevant_orig=relevant[match.start():match.end()]
                    relevant_new=relevant_orig
                    relevant_new=relevant_new.replace('selected(', '')
                    relevant_new=relevant_new.replace(',', '=')
                    relevant_new=relevant_new.replace("'", '')
                    relevant_new=relevant_new.replace(")", '')
                    relevant=relevant.replace(relevant_orig, relevant_new)
                    print relevant
                    
                    
                
            relevantwrite='Only ask if ' + relevant
            relevantwrite=outdoc.add_paragraph(relevantwrite)            
    
        #Hint
        if not hint=='':
            hintwrite=outdoc.add_paragraph()
            hintwrite.add_run(hint).italic=True
    
        #Choices
        if qtype=="text" or qtype=="string":
            columnwrite=outdoc.add_paragraph('................................................................................................................................................')
    
        if qtype=="integer" or qtype=="decimal":
            columnwrite=outdoc.add_paragraph('-             -')
        if qtype=='select_one' or qtype=='select_multiple':
            if qtype=='select_one':
                bullet='o'
            elif qtype=='select_multiple':
                bullet="u"
            if choices=='ID':
                optionwrite=outdoc.add_paragraph('A select multiple of ID 1-200')
            else:
                for option in choicesdict[choices]:
                    optionwrite=outdoc.add_paragraph(bullet + ' ' + option)
    
    
    
        questionnumber+=1

    

    outdoc.save(excelfile+'.docx')
    
print 'Finished'
