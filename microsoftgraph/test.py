from microsoftgraph.client import Client
import pprint
from urllib.parse import urlencode, urlparse, quote_plus

client_id = "77683f90-a032-47fb-a71d-41e909137513"
client_secret = "ucXJSF#{:;jpqgmMXF91394"

petition = Client(client_id=client_id, client_secret=client_secret)
data_token = {"token_type": "Bearer",
              "scope": "User.Read Files.ReadWrite Contacts.Read Contacts.ReadWrite Files.ReadWrite.All Mail.Send Mail.Read Calendars.ReadWrite",
              "expires_in": 3600, "ext_expires_in": 0,
              "access_token": "EwBQA8l6BAAURSN/FHlDW5xN74t6GzbtsBBeBUYAAVGk0V/lIZhO0fEFcUdGGDEWd9k+MRdWfnZSSRjF8fFTVCXM/ObhQ6yjWvxY6HG+BNa7LZLC2nzH/7jiusbCwHuxgIrPCIeD9p+q4QndZEvGo37r5K0DG5099AK2/FvCNs0PqkhIR/pCDD7g5yaEksjbDhBAnb+6js62bgUyGzumXLPJqiMQb19Tqcco2d+Uprx5d9VJnIk4QiXJMqdCaMevEKzzVTARHVurlx3R23Eu7yiVRmuUanCUoIfFzuKhDfSvyMs2DydNaDqLvki2RVT2AsH4XbPyzbefZmzVsMB5TUA6VDrjNYOMKPpbj5eBPcROiemUWyzsbOcW0vig8zIDZgAACHInkL0LetadIAKghaseL3Zqs2/4GCX8nyWkSWtikryJOA1rNUJp5K2na5+XtG4zUmluWm8BA5y55NamWvCOBspR88+0E3kpvA0uz5ANmtLtHtFDcDITDnbSnXVbOQZQpnD6i1f1vcO3OCfuKXoBawx2R+6pV2lYcS1Zs/ZYeFIbBv9gRMOziL5CWc1M6CT4dG4JeH73zzk5L5jHAyNL11pZsP2RTjWrJE5s5MqVcv7R72XIZnud35IIfKMGQfs4X49jTb8RKKR6eekz/4B+xzMQyDkbEz9begs/MqA/PzTqhpJnbRAV+cSgsmaVyXHlbc6BktVTfqokULMKlTEEdEty5JDnQqQ1jvIgnsdYMVFwnGFbT340mycki5DrzOiIjoj3UFKeTy1DS7LoUB5ej1nlAKXWEcGqta23x9FlfOS1XAQj1DveRKHstfFJxPYUnGlR9/sspu1NhoJnfx2cJSp4RizDxjxWiORoH3hnkDBrszuqZ3kMdDmEGIV4DBEr2hFwAjUUM7Iy+uZLvFt1sSVlAbG8ia8KI4m/VRdamYzu3ag0BWUFveVBCkh1RjQJmZbv3FFXVzBmIFsk10CKYOAGME2MRujlsbNODx2+9r3CQEVNXf8s3dyDzDk44JczATftN4WWi7bDcVdpkMjEiVNQxFv3aHSif9VfQ/yCPXU765dcYLyn+hsyrP25hWxDDz2CQxBulZmxTd794F+3AzmSKE51OJGmYkzaUgI=",
              "refresh_token": "MCepaWXzkQ5wX14EyIkuQNPFM2BEEjUvixVOwhZPLp3lOKV7Q!58bAuenS8w7nlrq07DZut2tUqFCztljCIZqqgeINO6Tdru9HaF5pcdf7HcTShBHmKn6ItGUZlha4GrKgSvwlrwGqiMW!IUXCJJ6p7PgFegXyYnvWUKugJ9jAsKFaPtbd5o6z!gCMzm7HFPilwGbHJlgTt2e7VAHpYOQMe!EW5nC9ZfT4pVduDP17YnHPiVfKXTtcfoA6XqT0lb6M6UHZdR!8Rg3fdASGTJvIEv6a2A*jk220mKf4BYvayFlAxW*kX4LRBX2uxv*2BCDx7YCGc8O*GIHVQw129i!RFak*4iaBC0yYUYfNjYaOIGehtUA8fqtjRnlzrs4y2itxg$$"}

folders = ['6D03D5E478E60BD4!153', '6D03D5E478E60BD4!1982', '6D03D5E478E60BD4!1371', '6D03D5E478E60BD4!1347']
files = ['6D03D5E478E60BD4!2189']
worksheet_id = ['{00000000-0001-0000-0000-000000000000}', '{00000000-0001-0000-0100-000000000000}',
                '{00000000-0001-0000-0200-000000000000}', '{00000000-0001-0000-0300-000000000000}',
                '{00000000-0001-0000-0400-000000000000}']

school_file = ['01ROQ7VNPPORM32R5ODVD2GJA42ZGDYSH7']
worksheet_school = ['{00000000-0001-0000-0000-000000000000}', '{00000000-0001-0000-0100-000000000000}']

# r = petition.refresh_token("https://28c0c29b.ngrok.io/excel/oauth", data_token['refresh_token'])
# pprint.pprint(r)

petition.set_token(data_token)
# data = {
#     "givenName": "Guzman",
# "surname": "RODRIGUES",
# "emailAddresses": [
#     {
#         "address": "guzman@mmg.com",
#         "name": "EMAIL"
#     }
# ],
# "businessPhones": [
#     "+57 12345459"
# ]
# }

# data = {
#     'profession': 'Estudent',
#     'businessAddress': {
#         "city": "bogota",
#         "countryOrRegion": "cundinamarca",
#         "postalCode": "111141",
#         "state": "bogota",
#         "street": "Cra"
#     },
#     'middleName': 'gelvez',
#     'emailAddresses': [
#         {'name': 'EMAIL',
#          'address': 'yordy.gelvez@gmail.com'}
#     ],
#     'givenName': 'yordy', 'mobilePhone': '3144773179',
#     'homeAddress': {
#         "city": "bogota",
#         "countryOrRegion": "cundinamarca",
#         "postalCode": "111141",
#         "state": "bogota",
#         "street": "Cra"
#     },
#     "businessPhones": [
#         "+1 732 555 0102"
#     ]
# }


# create = petition.outlook_create_me_contact(json=data)
# print(create)

# contacts = petition.outlook_get_me_contacts(params={"$top": "1", "$orderby": "createdDateTime desc"})
# pprint.pprint(contacts)

# data = petition.get_metadata()
# pprint.pprint(data)

# folders = petition.outlook_get_contact_folders()
# pprint.pprint(folders)
#
# data_id = "AQMkADAwATM3ZmYAZS1iZTIzLWEzYTQtMDACLTAwCgAuAAADpjgmjJZTxECfi4sAXBGnLa4BAMzOBDmIdJtLsV2Ylx9GiTEAAYIhX5kAAAA="
# create = petition.outlook_create_contactin_folder(data_id, json=data)

# response = petition.outlook_get_me_contacts(params={"$top": "1", "$orderby": "createdDateTime desc"})
# pprint.pprint(response)

# data_folder = {
#   "displayName": "SQUEEZE FOLDER"
# }
# create = petition.outlook_create_contact_folder(json=data_folder)
# print(create)

# my_id = "AQMkADAwATM3ZmYAZS1iZTIzLWEzYTQtMDACLTAwCgBGFBAF1585-54F1-4E52-A1C4-6C4F463101A2AAADpjgmjJZTxECfi4sAXBGnLa4HAMzOBDmIdJtLsV2Ylx9GiTEAAAIBDgAAAMzOBDmIdJtLsV2Ylx9GiTEAAYBwOF8AAAA="
# contacts = petition.outlook_get_me_contacts(data_id=my_id)
# print(contacts)
# print(contacts["id"])

#################################################################

"""EXCEL TEST"""
# data = petition.drive_root_items()
# pprint.pprint(data)
#
# data1 = petition.drive_root_children_items()
# # pprint.pprint(data1)
# for folder in data1['value']:
#     print(folder['name'], folder['id'])

# data2 = petition.drive_specific_folder(folders[2])
# pprint.pprint(data2)
#
# data3 = petition.excel_get_worksheets('6D03D5E478E60BD4!2189')
# pprint.pprint(data3)


addheader = {
    "workbook-session-id": 'cluster=US1&session=12.c2bfeb36dffc1.A172.1.E92.http%3a%2f%2ftier1%3fid%3dhttps%253A%252F%252Fwopi%252Eonedrive%252Ecom%252Fwopi%252Ffiles%252F6D03D5E478E60BD4%2521218914.5.en-US5.en-US24.6d03d5e478e60bd4-Private1.A24.DKmtSjChoU%2bo1UZCVKZrxQ%3d%3d14.16.0.9116.502714.5.en-US5.en-US1.M1.N0.1.A'}

# data_session = {
#   "persistChanges": True
# }
# data3 = petition.drive_create_session(files[0], json=data_session)
# pprint.pprint(data3)

data31 = petition.drive_refresh_session(files[0], headers=addheader)
print(data31)
#
# worksheet = {"name": "soy sho 3"}
# data6 = petition.excel_add_worksheet(files[0], json=worksheet)
# print(data6)

# data7 = petition.excel_get_specific_worksheet(item_id=files[0], worksheet_id=worksheet_id[0])
# print(data7)

# up_worksheet = {
#     'position': 0,
#     'name': 'MAMAGUEVOS',
#     'visibility': 'Visible'
# }
# data8 = petition.excel_update_worksheet(item_id=files[0], worksheet_id=worksheet_id[4], json=up_worksheet)
# print(data8)

data9 = petition.excel_get_charts(item_id=files[0], worksheet_id=worksheet_id[1], headers=addheader)
print(data9)

# chart_data = {
#     "type": "ColumnStacked",
#     "sourceData": "A1:D1",
#     "seriesBy": "Auto"
# }
# data9 = petition.excel_add_chart(item_id=files[0], worksheet_id=worksheet_id[1], json=chart_data)
# print(data9)


# data = {
#     "address": "Hoja1!A1:D1",
#     "hasHeaders": True
# }
# data10 = petition.excel_add_table(item_id=school_file[0], json=data)
# print(data10)

#
# tables = petition.excel_get_tables(item_id=school_file[0])
# print(tables)

# data = {
#     # "index": 2,
#     # "values": [
#     #     [1, 2, 3, 4],
#     #     [4, 5, 6, 7]
#     # ]
#     "values": [['hola']]
# }
result = petition.excel_add_row(item_id=school_file[0], worksheets_id=worksheet_school[1], table_id=1, json=data)
print(result)

# close = petition.drive_close_session(item_id=files[0], headers=addheader)
# print(close)

# get_rows = petition.excel_get_rows(item_id=school_file[0], table_id=2)
# print(get_rows)


# get_cell = petition.excel_get_cell(item_id=files[0], worksheets_id=worksheet_id[1])
# print(get_cell)

# data = {
#     "name": "Prueba",
# }
# create_column = petition.excel_add_column(item_id=school_file[0], worksheets_id=worksheet_school[0], table_id=2,
#                                           json=data)
# print(create_column)

# data = {
#     "values": [['aguevonados']]
# }
# cell = petition.excel_add_cell(item_id=school_file[0], worksheets_id=worksheet_school[1], json=data)

# range = petition.excel_get_range(item_id=files[0], worksheets_id=worksheet_id[1])
# print(range)
