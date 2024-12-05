# -*- coding: utf-8 -*-

import requests
import pandas as pd

df = pd.read_excel('C:\\Users\\Murali\\Desktop\\IP Distro_09062022.xlsx')
df1 = pd.DataFrame()
for index, row in df.iterrows():
    i = row["PropertyName"]
    j = row['Vendor']
    url = str(i)
    response = requests.get(url)
    soup = BeautifulSoup(response.text)
    anchors = soup.find_all('a', {'href': True})
    
    for anchor in anchors:
        d = anchor.text
        data = d.split(",")
        df2 = pd.DataFrame(data, columns=["output"])
        df2["Client ID"] = j
        df2["URL"] = i
        print(df2)
        df1 = df1.append(df2, ignore_index = True)
df1.to_excel('C:\\Users\\johnk\\Desktop\\ClientInventoryLookup1.xlsx')
