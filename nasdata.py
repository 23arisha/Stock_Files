from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
import datetime

#Borders
border_style = 'thin'  # Choose the border style you want ('thin', 'medium', 'thick', etc.)

border = Border(left=Side(style=border_style),right=Side(style=border_style),top=Side(style=border_style),bottom=Side(style=border_style))
###########################################

# Set up the Selenium service and driver
path = "C:/Users/Selltricks/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe"
s = Service(path)
options = webdriver.ChromeOptions()
options.add_argument("--incognito")

driver = webdriver.Chrome(service=s, options=options)

url = "https:/stockhouse.com/markets/stocks/nasdaq"
driver.get(url)

time.sleep(15)

# CLICK MORE BUTTON
wait = WebDriverWait(driver, 10)
element = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "mhcs-st-more")))

more_elements = driver.find_elements(By.CLASS_NAME, "mhcs-st-more")

for more_element in more_elements:
    driver.execute_script("arguments[0].click();", more_element)
    time.sleep(2)  

##############################################################

# FIND TABLE
tables = driver.find_elements(By.CSS_SELECTOR, "table.mhcs-st")
table=tables[3]

#['Symbol', 'Company', 'Last', '$ Chg', '% Chg'] HEADERS
tr_tag=table.find_element(By.CSS_SELECTOR, "tr.mhcs-st-row.mhcs-header")
td_sp=tr_tag.find_elements(By.CSS_SELECTOR, 'span.mhcs-header-title')
header=[i.text for i in td_sp]

index=1
if index<len(header):
    del header[index]

header.append('Date')
# Create a workbook and a worksheet
workbook = openpyxl.Workbook()
sheet1=workbook.active
sheet1.title="52weekhigh"

# Populate the worksheet with headers
sheet1.append(header)
header_row = sheet1[1]
for cell in header_row:
    cell.border = border
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font=Font(bold=True)

#DATE
ele = driver.find_element(By.CSS_SELECTOR, "div.mhcs-headers > div:nth-child(2)")
text = ele.text.strip().split(",")  # Get the text content and remove leading/trailing spaces

date_str=text[0].strip()
date_object = datetime.datetime.strptime(date_str, '%b %d')
date = date_object.strftime('%d-%b')
date_s=date_object.strftime('%b%d')

# NASDAQ PERCENT GAINER

print("NDAQ PERCENT GAINER")
data_tr=table.find_elements(By.CSS_SELECTOR, "tr.mhcs-st-row.mhcs-pointer")
for j in data_tr:
    td_tag=j.find_elements(By.CSS_SELECTOR, "td.mhcs-st-col")
    data=[x.text for x in td_tag]
    index=2
    if index < len(data):
        del data[index]

    data.append(date)

    sheet1.append(data[1:])
    for row in sheet1.iter_rows(min_row=2, max_row=sheet1.max_row, min_col=1, max_col=sheet1.max_column):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

##########################################

# NASDAQ PERCENT DECINER
# Create a new worksheet for "NDAQ PERCENT DECLINER"
ws_decline = workbook.create_sheet(title="52weeklow")
ws_decline.append(header)
ws_decline.append(header)
header_row = ws_decline[1]
for cell in header_row:
    cell.border = border
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font=Font(bold=True)

print("NDAQ PERCENT DECLINER")
table_4=tables[4]
data_tr_4=table_4.find_elements(By.CSS_SELECTOR, "tr.mhcs-st-row.mhcs-pointer")
for d in data_tr_4:
    td_ta=d.find_elements(By.CSS_SELECTOR, "td.mhcs-st-col")
    data_4=[t.text for t in td_ta]
    index_2=2
    if index < len(data_4):
        del data_4[index_2]
    
    data_4.append(date)
   
    ws_decline.append(data_4[1:])
    for row in ws_decline.iter_rows(min_row=2, max_row=ws_decline.max_row, min_col=1, max_col=ws_decline.max_column):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

# SAVE THE WORKBOOK
# filename = input("Enter filename: ")
output_file = r"C:/Users/Selltricks/Downloads/52week"+ date_s + ".xlsx"
workbook.save(output_file)