# psx
PsxScrapper is a repository for scripts that read data from Pakistan Stocks Exchange mirror website; as mentioned 'http://hamariweb.com/finance/stockexchanges/psx_market_summary.aspx'. The ups and downs in the stocks are time dependent and for analysis purposes we must make sure that we have the correct data. This script is the first step towards making all that possible. 

###### Getting started with the code:

** Initializing an array of headers for the data to be extracted **
```
text1 = ['SCRIP', 'LDCP', 'Open', 'High', 'Low', 'Curr.', 'Vol.', 'Chng']
```

** Reading the metadata from psx's(Pakistan Stocks Exchange) mirror site **
```
url = 'http://hamariweb.com/finance/stockexchanges/psx_market_summary.aspx'
source_code = requests.get(url)
plain_text = source_code.text
```

** Using BeautifulSoup lib and an html.parser for skimming across the extracted data **
```
soup = BeautifulSoup(plain_text, "html.parser")
```

** Setting name of the excel document in which filtered data will be written to current data and time - Script execution data and time **
```
temp = str(datetime.datetime.now())
for a in temp:
    if a != ':':
        filename.append(a)
```

** Initializing and adding an excel document in the workbook **
```
workbook = xlsxwriter.Workbook(''.join(filename)+'.xlsx')
worksheet = workbook.add_worksheet()
```

** Setting the widths of different cells - Presentation purpose **
```
worksheet.set_column(0, 0, 25)
worksheet.set_column(6, 6, 10)
```

** Formatting and appending names of the headers in the excel document as initialized above **
```
for j in range(0,8,1):
        if m < len(text1):
            format = workbook.add_format({'bold': True, 'font_color': 'green'})
            worksheet.set_row(1, 18, format)
            worksheet.write(1, j, text1[m])
            m += 1
```

** Formatting and appending current data and time in the excel document - DateTimeStamp **
```
format = workbook.add_format({'bold': True})
worksheet.set_row(0, 18, format)
worksheet.write(0,0,"TimeStamp")
worksheet.write(0,1, str(datetime.datetime.now()))
```

###### Implementation for extracting data against [SCRIP, LDCP, Open, High, Low, Current, Volume] which is encapsulated by a span element per the respective header

** Filtering, extracting, formatting and appending marketData_advance class data encapsulated by span element **
```
m = 0
for link in soup.findAll('span', {'class': 'marketData_advance'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
s = len(cmp1)
for i in range(2,s,1):
    for j in range(0,7,1):
        if m < s:
            worksheet.write(i,j,''.join(cmp1[m]))
            m += 1
cmp1.clear()
m = 0
```

** Filtering, extracting, formatting and appending marketData_decline class data encapsulated by span element **
```
for link in soup.findAll('span',{'class': 'marketData_decline'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
t = len(cmp1)
index = int(s/7);
for i in range(index,index+t,1):
    for j in range(0,7,1):
        if m < t:
            worksheet.write(i,j,''.join(cmp1[m]))
            m += 1

cmp1.clear()
m = 0
```

** Filtering, extracting, formatting and appending marketData_noChange class data encapsulated by span element **
```
for link in soup.findAll('span',{'class': 'marketData_noChange'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
u = len(cmp1)
index1 = int(index+t/7)
for i in range(index1,index1+u,1):
    for j in range(0,7,1):
        if m < u:
            worksheet.write(i,j,''.join(cmp1[m]))
            m += 1

cmp1.clear()
m = 0
```

###### Implementation for extracting data against [Change] which is encapsulated by div element'''

** Filtering, extracting, formatting and appending marketData_advance class data encapsulated by div or section element **
```
for link in soup.findAll('div',{'class': 'marketData_advance'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
v = len(cmp1)
for i in range(1,s,1):
    if m < v:
        worksheet.write(i,7,*cmp1[m])
        m += 1

cmp1.clear()
m = 0
```

** Filtering, extracting, formatting and appending marketData_decline class data encapsulated by div or section element **
```
for link in soup.findAll('div',{'class': 'marketData_decline'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
w = len(cmp1)
for i in range(index,index+w,1):
    if m < w:
        worksheet.write(i,7,*cmp1[m])
        m += 1

cmp1.clear()
m = 0
```

** Filtering, extracting, formatting and appending marketData_noChange class data encapsulated by div or section element **
```
for link in soup.findAll('div',{'class': 'marketData_noChange'}):
    metadata = link.text
    lines = metadata.split()
    cmp1.append(lines)
x = len(cmp1)
for i in range(index1,index1+x,1):
    if m < x:
        worksheet.write(i,7,*cmp1[m])
        m += 1
m = 0
```
** Closing the workbook allocated for the excel document **
```
workbook.close()
```




