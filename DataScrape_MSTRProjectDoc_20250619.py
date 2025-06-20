import os
from bs4 import BeautifulSoup
import pandas as pd

'''
Filename: DataScrape_MSTRProjectDoc_20250619.py
Description: This python program reads a folder full of HTML files that were created by MicroStrategy Project Documentation Export
Author: Vik Makrani
Date: 20250619
'''

html_folder = r"C:\Users\User\Desktop\MSTR Project Doc\Test_Python\Supply Chain Analytics-TEST Environment (20250618112430)\Supply Chain Analytics-TEST Environment (20250618112430)"
#attributeSamplePage = r"C:\Users\User\Desktop\MSTR Project Doc\Test_Python\DA48790140DAE778FD2C70BACD59E3F8_1.html"
#metricSamplePage = r"C:\Users\User\Desktop\MSTR Project Doc\Test_Python\E10BACDA45718B594D658A8267B399DE_1.html"
#factSamplePage = r"C:\Users\User\Desktop\MSTR Project Doc\Test_Python\0269187C41043B4D6FB0D291CA96BA9E_1.html"
excel_path = r"C:\Users\User\Desktop\MSTR Project Doc\Test_Python"
project_name = "Supply Chain Analytics"
attribute_folder = r"\Schema Objects\Attributes"
metric_folder = r"\Public Objects\Metrics"
fact_folder = r"\Schema Objects\Facts"

attributes = []
metrics = []
facts = []

def getAttributeDetails(soup):
    loc_tds = soup.find_all("td")

    # Print to see values are correct
    #print(loc_tds[2].text.strip()) #name
    #print(loc_tds[6].text.strip()) #location
    #print(loc_tds[20].text.strip()) # ID
    #print(loc_tds[57].text.strip()) #column
    #print(loc_tds[59].text.strip()) #table
    #print(loc_tds[73].text.strip()) #data type

    attributes.append({
        "MicroStrategy Project": project_name,
        "Attribute Name": loc_tds[2].text.strip(),
        "Attribute Location": loc_tds[6].text.strip(),
        "Attribute Column": loc_tds[57].text.strip(),
        "Attribute Table": loc_tds[59].text.strip(),
        "Attribute Data Type": loc_tds[73].text.strip(),
        "Attribute ID": loc_tds[20].text.strip(),
    })

    return attributes

def getMetricDetails(soup):
    loc_tds = soup.find_all("td")

    # Print to see values are correct
    #print(loc_tds[2].text.strip()) #name
    #print(loc_tds[6].text.strip()) #location
    #print(loc_tds[20].text.strip()) #ID
    #print(loc_tds[45].text.strip()) #metric type
    #print(loc_tds[47].text.strip()) #metric formula

    metrics.append({
        "MicroStrategy Project": project_name,
        "Metric Name": loc_tds[2].text.strip(),
        "Metric Location": loc_tds[6].text.strip(),
        "Metric Type": loc_tds[45].text.strip(),
        "Metric Formula": loc_tds[47].text.strip(),
        "Metric ID": loc_tds[20].text.strip()
    })
    return metrics

def getFactDetails(soup):
    loc_tds = soup.find_all("td")

    # Print to see values are correct
    #print(loc_tds[2].text.strip()) #name
    #print(loc_tds[6].text.strip()) #location
    #print(loc_tds[20].text.strip()) #ID
    #print(loc_tds[39].text.strip()) #fact expression
    #print(loc_tds[41].text.strip()) #fact table

    facts.append({
        "MicroStrategy Project": project_name,
        "Fact Name": loc_tds[2].text.strip(),
        "Fact Location": loc_tds[6].text.strip(),
        "Fact Column": loc_tds[39].text.strip(),
        "Fact Table": loc_tds[41].text.strip(),
        "Fact ID": loc_tds[20].text.strip()
    })

    return facts

# this is to read from all files in folder variable html_folder
for filename in os.listdir(html_folder):
    if not filename.lower().endswith(".html"):
        continue

    file_path = os.path.join(html_folder, filename)
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        soup = BeautifulSoup(file, "html.parser")
    print("Analyzing file: ",file_path)
    object_block = soup.find_all("table", class_="MAINBODY", border="1")
    for i in object_block:
        loc_tds = i.find_all("td")
        objLocation = loc_tds[6].text.strip()
        if attribute_folder in objLocation:
            print("This object is an attribute")
            attributes = getAttributeDetails(i)
        elif metric_folder in objLocation:
            print("This object is an metric")
            metrics = getMetricDetails(i)
        elif fact_folder in objLocation:
            print("This object is a fact")
            facts= getFactDetails(i)


# write attributes, metrics, and facts to an Excel file
df_attributes = pd.DataFrame(attributes)
df_metrics = pd.DataFrame(metrics)
df_facts = pd.DataFrame(facts)

print(f"\nSummary: {len(attributes)} attributes, {len(metrics)} metrics, {len(facts)} facts")

output_path = os.path.join(excel_path, project_name+" Data Dictionary.xlsx")
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_attributes.to_excel(writer, sheet_name="Attributes", index=False)
    df_metrics.to_excel(writer, sheet_name="Metrics", index=False)
    df_facts.to_excel(writer, sheet_name="Facts", index=False)

print(f"\nExcel file saved to: {output_path}")