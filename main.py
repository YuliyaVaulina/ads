import pandas as pd
import re
from operator import itemgetter

'''Ссылки на входящие файлы.
Подразумевается, что файл app-ads.xlsx находится в папке со скриптом.
Файл tomerge.txt можно выбрать из любой директории'''

app_ads_xmlsx = 'app-ads.xlsx'
locate = input('Укажите путь к файлу tomerge.txt или ., если файл в текущей папке: ')

first_of_all_delete = 'Caramel Ads'
data_ads = []
new_app_ads = {}

try:
    with open(fr'{locate}\tomerge.txt', 'r') as f:
        publisher_tomerge = f.readline().strip()
        adds_tomerge = f.readlines()
        for line in adds_tomerge:
            data_ads += [re.sub(r"\s+", "", line).split(',') + [publisher_tomerge]]
except Exception as error:
    print(error)
    print('Некорректный путь к файлу или название файла')

try:
    app_ads = pd.read_excel(app_ads_xmlsx)
    list_app_ads = app_ads.to_dict(orient='records')
    for data in list_app_ads:
            data_ads += [re.sub(r"\s+", "", data['app-ads.txt']).split(',') + [data['Publisher']]]
except Exception as error:
    print(error)
    print('Файл app-ads.xlsx не найден.')

for data in data_ads: 
    if len(data) == 4:
        sample = ', '.join([data[0].lower()] + [data[1]] + ['RESELLER' if 'RES' in data[2] else 'DIRECT'])
        if sample not in new_app_ads:
            new_app_ads[sample]  = [data[-1]]
        else:
            new_app_ads[sample] += [data[-1]]
    if len(data) > 4:
        sample = ', '.join([data[0].lower()] + [data[1]] + ['RESELLER' if 'RES' in data[2] else 'DIRECT'])
        if sample not in new_app_ads:
            new_app_ads[sample] = [[str(data[3]) if len(str(data[3]))>=9 else '', data[-1]]]
        else:
            new_app_ads[sample] += [[str(data[3]) if len(str(data[3]))>=9 else '', data[-1]]]

insert_xlsx = []

for key, value in new_app_ads.items():
    if len(value) > 1:
        for element in value:
            if first_of_all_delete in element:
                value.remove(element)
                break    
    if len(value) > 1:
        del value[:-1]
    for element in value:
        insert_xlsx += [key.split(', ') + element.split(', ')] if type(element) == str else [key.split(', ') + [element]]


sorted_insert_xlsx = sorted(insert_xlsx, key=itemgetter(0, 1))


df = pd.DataFrame({'app-ads.txt': [', '.join(i[:-1]+ [i[-1][0]]).strip(', ') 
                                    if type(i[-1]) == list else ', '.join(i[:-1]).strip(', ') 
                                    for i in sorted_insert_xlsx],
                   'Publisher': [i[-1][1] 
                                if type(i[-1]) == list else i[-1] 
                                for i in sorted_insert_xlsx]})

df.to_excel('new_app-ads.xlsx', index= False )

print('Файл new_app-ads.xlsx готов')
