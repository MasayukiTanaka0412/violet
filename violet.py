import logging
import win32com.client

logging.basicConfig(level=logging.INFO)

logging.info('Violet App Start')

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
mail = outlook.OpenSharedItem(r"C:\Users\masayuki.tanaka\source\repos\Violet\Violet\年末年始のご挨拶.msg")

logging.info("件名: {}".format(mail.subject))
logging.info("本文: {}".format(mail.HTMLBody))

originalBody = mail.HTMLBody

replacedBody = originalBody.replace("{Recipient Name}","TestName")

mail.HTMLBody = replacedBody

mail.SaveAs(r"C:\Users\masayuki.tanaka\source\repos\Violet\Violet\年末年始のご挨拶2.msg")

logging.info('Violet App End')