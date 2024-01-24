import pandas as pd
# py -m pip install pandas 
import numpy as np



import gspread
credentials={
  "type": "service_account",
  "project_id": "id",
  "private_key_id": "key id",
  "private_key": "key",
  "client_email": "bot email",
  "client_id": "id",
  "auth_uri": "auth url",
  "token_uri": "token",
  "auth_provider_x509_cert_url": "certs",
  "client_x509_cert_url": "",
  "universe_domain": "googleapis.com"
}
#to get this credentials for your account visit https://docs.gspread.org/en/latest/index.html

spreadsheetNum:int=1
excelFileName="data directory"


class Transaction:
    context:str
    type:str
    amount:int
      
    def __init__(self,type="",context="" ,amount=0) -> None:
        
        self.type=type
        self.context=context
        self.amount=amount

    def from_list(self, l:list):
        self.type=l[0]
        self.context=l[1]
        self.amount=l[3]
     
    def __str__(self) -> str:
        return self.context+" "+self.type+" "+str(self.amount)


gc = gspread.service_account_from_dict(credentials)
sh=gc.open("Personal Finances")

worksheet=sh.get_worksheet(spreadsheetNum)

#Important change this data path when changing the data
df: pd.DataFrame=pd.read_excel(excelFileName)
# py -m pip install xlrd 
df=df.iloc[14:,:-2]

#some data manipulation for Garanti Bank
df=df.rename(columns={
    "Unnamed: 0":"Date",
    "T. GARANTİ BANKASI A.Ş.\nGenel Müdürlük: Nispetiye Mah. Aytar Cad.No: 2, Beşiktaş, 34340, Levent, İstanbul\nBüyük Mükellefler Vergi Dairesi Başkanlığı Vergi No: 8790017566\nMersis Numarası: 0879 0017 5660 0379\nwww.garantibbva.com.tr":"Content",
    "Unnamed: 2":"Type",
    "Unnamed: 3":"Amount"
})
df.reset_index(drop=True, inplace=True)

#REGION DATACLEANUP
cont=np.array(df.loc[:,"Content"])
amount=np.array(df.loc[:,"Amount"])


typeList=[]
contextList=[]

for i in range(0,len(cont)):
    context: np.ndarray=np.full(shape=(2),fill_value="Bilgi Yok",dtype="U30")
    elements=cont[i].split("-")
    if(elements[0][-11:-1]=="YKP Ödemes" or elements[0][:3]=="FAS" or elements[0][:11]=="Mobil DÖVİZ"):
        df.drop(
            labels=i,
            axis=0,
            inplace=True
        )
    else:
        if(len(elements)>=3):
            context[0]=elements[0]
            context[1]=elements[2]
        else:
            context[:len(elements)]=elements[:]
        
        typeList.append(context[0])
        contextList.append(context[1])


df.drop(
    columns="Content",
    inplace=True
)

df.insert(1,"Transaction Type",typeList)
df.insert(2,"Context",contextList)
df.reset_index(drop=True, inplace=True)


print(df)
#Something to be careful i did not check the year, so if you miss up the year result will be bad
# dic={
#     "12":{"first half":[1,-2,-3], -4, types={"Konukevi":40....},
#            "second half":[-4,-5,-6,7]
#            },
#     "11":{...
#           }
# } structure is like this
months={}        
dictionaryOf={
    "KONUKEVI":0,
    "MİGROS":0,
    "Diğer":0
}

def dicFiller(months,date,transaction, half):
    months[date][half][1]+=transaction.amount
    months[date][half][0].append(transaction)
    months[date]["Total"] += transaction.amount
    keyWord=transaction.context.split(" ")[0]
    if(keyWord in dictionaryOf):
        months[date][half][2][keyWord]+=transaction.amount
    else:
        if(keyWord=="TOBB"):
            months[date][half][2]["KONUKEVI"]+=transaction.amount
        else:
            months[date][half][2]["Diğer"]+=transaction.amount


for i in range(0,len(df)):
    date=df.iloc[i,0].split("/")
    transaction:Transaction=Transaction()
    transaction.from_list(df.iloc[i,1:].to_list())

    if(str(date[1])[0]=="0"):
        date[1]=str(date[1])[1]

    if(str(date[1]) not in months):
        months[str(date[1])]={
            "First Half":[[],0, dict(dictionaryOf)],
            "Second Half":[[],0, dict(dictionaryOf)],           
            "Total":0
        }
    if(int(date[0])<=15):
        dicFiller(months=months,date=str(date[1]),transaction=transaction,half="First Half")
    else:
        dicFiller(months=months,date=str(date[1]),transaction=transaction,half="Second Half")


alphabet_array = np.array([chr(i) for i in range(65, 91)])
monthArray = np.array(range(1, 13))
halfList=["First Half","Second Half"]

for month in monthArray:
    if(str(month) not in months):
        continue

    f : int = 0
    for half in halfList:
        expenseList=list(months[str(month)][half][2].values())
        print(expenseList)
        c=alphabet_array[month*2-1+f]
        print(c)
        
        cell_list = worksheet.range(f"{c}4:{c}6")
        
        i:int=0
        for cell in cell_list:
            cell.value = expenseList[i]
            i+=1
            # Update in batch
        worksheet.update_cell(row=7, col = month*2+f, value=months[str(month)][half][1])
        print(months[str(month)][half][1])
        worksheet.update_cells(cell_list)
        f+=1
    worksheet.update_cell(row=8, col = month*2, value=months[str(month)]["Total"])
    



    
    