from email import header
import win32com.client
import pandas as pd
import os

csv = pd.read_csv('Copy of Reg 3rd Phase calling final Update - Sheet1 (1).csv',header=None)

psApp = win32com.client.Dispatch("Photoshop.Application")


psApp.Open(r"C:\Users\goura\OneDrive\\Desktop\alumni card\\ACC.psd")

doc = psApp.Application.ActiveDocument
# for key in dir(doc):
#     method = getattr(doc,key)
#     if str(type(method)) == "<type 'instance'>":
#         print (key)
#         for sub_method in dir(method):
#             if not sub_method.startswith("_") and not "clsid" in sub_method.lower():
#                 print ("\t")
#                 print(sub_method)
#     else:
#         print ("\t");
#         print(method);
# print(csv[3][1])
temp = "ANIKET PATEL"
batch = "DESIGN TEAM"
for i in range(1,len(csv)):
    layer_facts = doc.ArtLayers[temp]
    text_of_layer = layer_facts.TextItem
    text_of_layer.contents = csv[0][i]
    # temp = csv[3][i]

    layer_facts = doc.ArtLayers[batch]
    text_of_layer = layer_facts.TextItem
    text_of_layer.contents = csv[2][i]
    # batch = csv[32][i]


    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13   # PNG Format
    options.PNG8 = False  # Sets it to PNG-24 bit

    # pngfile = 'a' + 'b'

    pngfile = 'C:\\Users\\goura\\OneDrive\\Desktop\\alumni card\output\\' + csv[0][i] + '.png'
    # pngfile = 'umang.png'
    doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)







