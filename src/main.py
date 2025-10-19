import requests
import os
import math
import json
from openpyxl import load_workbook
from dotenv import load_dotenv

load_dotenv()

BASE_URL = os.getenv("BASE_URL")
API_KEY = os.getenv("API_KEY")
BASE_CURRENCY = os.getenv("BASE")
CURRENCIES = os.getenv("CURRENCIES")
WORK_DIRECTORY_PATH = os.getenv("WORK_DIRECTORY_PATH")
CURRENT_DIRECTORY = os.getcwd()
GRAMS_PER_OUNCE = 31.1035
CARAT = "22K"



def changeDirectory(directory):
    os.chdir(directory)

def hit_API():
    API_URL = f"{BASE_URL}?api_key={API_KEY}&base={BASE_CURRENCY}&currencies={CURRENCIES}"
    response = requests.get(API_URL)
    response.raise_for_status()
    data = response.json()
    print(str(data).replace("'" ,'"' ).replace("TRUE" , '"TRUE"'))

    if not os.path.exists("data.json"):
        file = open("data.json" , 'x')
        file.close()

    file = open("data.json" , "w")
    file.write(str(data).replace("'" ,'"' ))


def getFile(mode):
    file = open("data.json" , mode)
    return file

def extractGoldRate(currency):
    file = getFile("r")
    responseMap = json.loads(file.read())
    CURRENCY_PER_GRAM = round(responseMap["rates"][currency] / GRAMS_PER_OUNCE , 2)
    return CURRENCY_PER_GRAM

def displayGoldRateInfo():
    INR_PER_GRAM = extractGoldRate("INR")
    print(f"1 Gram Gold Rate in Indian Rupees : {INR_PER_GRAM}")

def updateGoldRateInSheet():
    changeDirectory(WORK_DIRECTORY_PATH)
    workbook_name = "Gold Loan Details.xlsx"
    sheet_name = "Interest Calculation"
    carat_22_cell = "I6"
    wb = load_workbook(workbook_name)
    changeDirectory(CURRENT_DIRECTORY)
    INTEREST_CALCULATION_SHEET = wb[sheet_name]
    INTEREST_CALCULATION_SHEET[carat_22_cell] = extractGoldRate("INR")
    changeDirectory(WORK_DIRECTORY_PATH)
    wb.save(workbook_name)
    changeDirectory(CURRENT_DIRECTORY)
    
    
displayGoldRateInfo()
updateGoldRateInSheet()




