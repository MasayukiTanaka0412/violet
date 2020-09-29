import logging
import os
import win32com.client
import pandas as pd

logging.basicConfig(level=logging.INFO)
templateName = "Template.msg"
recipientsFile = "Recipients.xlsx"

logging.info('Violet App Start')

path = os.getcwd()
logging.info('Current Directory {}'.format(path))

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

mail = outlook.OpenSharedItem(os.path.join(path,templateName))

logging.info("件名: {}".format(mail.subject))
logging.info("本文: {}".format(mail.HTMLBody))

originalBody = mail.HTMLBody

df = pd.read_excel(os.path.join(path,recipientsFile), sheet_name='Recipients')
logging.info(df)

outputDir = os.path.join(path,'output')
if not os.path.isdir(outputDir):
    os.mkdir(outputDir)

for index, row in df.iterrows():
    replacedBody = originalBody
    recipient = ""
    for indexName in row.index:
        logging.info("indexName {}".format(indexName))
        if indexName == "TO":
            mail.Recipients.Add(row[indexName])
            recipient =row[indexName]
        else:
            replacedBody = replacedBody.replace(indexName,row[indexName])
    mail.HTMLBody = replacedBody
    mail.SaveAs(os.path.join(outputDir,"{}.msg".format(recipient.replace("@","_"))))
    mail.Recipients.Remove(1)

logging.info('Violet App End')