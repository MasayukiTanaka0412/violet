import logging
import os
import win32com.client

logging.basicConfig(level=logging.INFO)
templateName = "Template.msg"

logging.info('Violet App Start')

path = os.getcwd()
logging.info('Current Directory {}'.format(path))

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

mail = outlook.OpenSharedItem(os.path.join(path,templateName))

logging.info("件名: {}".format(mail.subject))
logging.info("本文: {}".format(mail.HTMLBody))

originalBody = mail.HTMLBody

replacedBody = originalBody.replace("{Recipient Name}","TestName")

mail.HTMLBody = replacedBody

mail.SaveAs(os.path.join(path,"output.msg"))

logging.info('Violet App End')