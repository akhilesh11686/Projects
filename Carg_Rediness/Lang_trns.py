
# #%%
# from cgitb import text
# from tkinter import Text
# import six
# from google.cloud import translate_v2 as translate
# from google.oauth2 import service_account
# import pandas as pd

# def get_all_lang(target,txt):
#     credentials = service_account.Credentials.from_service_account_file(filename='sa-api-access.json',scopes=["https://www.googleapis.com/auth/cloud-platform"])
#     translate_client = translate.Client(credentials=credentials)
#     if isinstance(text,six.binary_type):
#         text = text.decode("utf-8")


#     result = translate_client.translate(text,target_language=target)
#     translated_text = result['translatedText']

#     print(u"Text: {}".format(result["input"]))
#     print(u"Translation: {}".format(result["translatedText"]))
#     print("whole text")
#     print("result")
# get_all_lang('hi','how are you')

