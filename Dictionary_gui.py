from tkinter import *
import requests
import json
import xlwt
import xlrd
from xlutils.copy import copy


app_id = '28a38761'
app_key = '2c1ebea46528df2b1d54d4a7906201f2'
language = 'en'
def search():
    global answer
    word_id=e1.get()
    patt0=xlwt.easyxf('font: name Times New Roman,color-index red ,bold on',num_format_str='#,##0.00')
    patt1=xlwt.easyxf('font: name Arial,color-index blue ,bold on')
    rb=xlrd.open_workbook('Dictionary.xls')
    r=rb.sheet_by_index(0)
    i=r.nrows
    wb=copy(rb)
    flag=False
    for j in range(i):
        if(r.cell(j,0).value==word_id):
            flag=True
            break
    if(flag==True):
        #print("...........value is already in the Dictionary file..........")
        #print(r.cell(j,1).value)
        answer=f"meaning:   {r.cell(j,1).value}"
    else:
        url = 'https://od-api.oxforddictionaries.com:443/api/v2/entries/'  + language + '/'  + word_id.lower()
        req=requests.get(url, headers = {'app_id' : app_id, 'app_key' : app_key})
        a=req.json()
        meaning=a['results'][0]['lexicalEntries'][0]['entries'][0]['senses'][0]['definitions'][0]
        #example=a['results'][0]['lexicalEntries'][0]['entries'][0]['senses'][0]['examples'][0]['text']
        #audio=a['results'][0]['lexicalEntries'][0]['pronunciations'][0]['audioFile']
        ws=wb.get_sheet(0)
        ws.write(i,0,word_id,patt1)
        ws.write(i,1,meaning,patt0) 
        wb.save('Dictionary.xls')
        #print("............New Entry entered in the Dictionary File............")
        #print(meaning)
        answer=f"meaning:   {meaning}"
    Label(root,text=answer,bd=2,font=2,fg='#ff0000',bg='white').grid(row=3,column=1,columnspan=2,padx=20,pady=70)
    import pyttsx3 as tts
    s=tts.init()
    s.say(answer)
    s.runAndWait()
    

root=Tk()
root.grid()
root.geometry("600x480+50+50")
root.configure(background='light blue')
root.title("Dictionary GUI")
l1=Label(root,text="Enter word",bd=3,font=3).grid(row=0,column=1,sticky="E",pady=50,padx=50)
e1=Entry(root,font=('Verdana,30'))
e1.grid(row=0,column=2,pady=50,padx=50)
e1.focus_set()
temp=Button(text='submit',bd=2,font=2,activebackground="blue",activeforeground="yellow",relief='groove',bg='white',fg='black',command=search)
temp.grid(row=1,column=2)
