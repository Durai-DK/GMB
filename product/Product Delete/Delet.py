from selenium import webdriver
from openpyxl import load_workbook
from GMB.Google.Google_login import Google
import time

wb = load_workbook(r"D:\Durai\GMB\product\Data\GMB Product URL.xlsx")
ws = wb.active