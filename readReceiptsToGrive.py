
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from datetime import datetime


#import sys

#import json
import re

import logging
import logging.config


#from openpyxl import Workbook

import pdfplumber

import os
import io

INPUT_DATE_FORMAT = '%d-%m-%Y'
OUTPUT_DATE_FORMAT = '%d/%m/%Y'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']

# The ID and range of a sample spreadsheet.
FOLDER_ID = '1E7inZqbV9klmwy2EEQoE8WcFqlu4T1XN';

SPREADSHEET_ID = '1u-X7uW8QHikSnDW5tt9SbqvtpPzmrnuYr2U84foEe14'
RANGE_NAME = 'Receipts!A2:I'

COLS = ["filename", 
      "receiptId", 
      "date", 
      "time",
      "name",
      "price",
      "tot",
      "discount",
      "vat"];

reStart = r"^\s+IVA\s+EURO\s+$";

reLinePrice = r"^([^\s].{1,22})\s+\*([acd])\s+([0-9]+),([0-9]{2})\s+$";

reDiscount = r"^\s+SCONTO\s+(FIDATY\s+)?[0-9]{2}%\s+([0-9]+),([0-9]{2})-S\s+$";
reTot = r"^TOTALE\s+EURO\s+([0-9]+),([0-9]{2})\s+\*\s+$";
reDate = r"^\s+([0-9]{2}-[0-9]{2}-[0-9]{4})\s+([0-9]{2}:[0-9]{2})\s+$";
reDocN = r"^\s+DOCUMENTO\s+N\.\s+([0-9\-]+)\s+$";

class Item:
  def __init__(self, filename, receiptId, recDate, recTime, name, price, vat):
    self.filename = filename
    self.receiptId = receiptId
    self.date = recDate
    self.time = recTime
    self.name = name
    self.price = price
    self.tot = price
    self.discount = 0
    self.vat = vat
  
  def __str__(self):
    return "receiptId:%s, %s, date:%s, time:%s, name:%s, price:%d, tot:%d, discount:%d, vat:%s" %( 
      self.filename,
      self.receiptId, 
      self.date, 
      self.time,
      self.name,
      self.price,
      self.tot,
      self.discount,
      self.vat)
  
  def toTuple(self):
    return (
      self.filename,
      self.receiptId, 
      self.date, 
      self.time,
      self.name,
      self.price,
      self.tot,
      self.discount,
      self.vat)

def getCredentials():
  creds = None
  # The file token.pickle stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.pickle'):
      with open('token.pickle', 'rb') as token:
          creds = pickle.load(token)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
      if creds and creds.expired and creds.refresh_token:
          creds.refresh(Request())
      else:
          flow = InstalledAppFlow.from_client_secrets_file(
              'credentials.json', SCOPES)
          creds = flow.run_console()
        
      # Save the credentials for the next run
      with open('token.pickle', 'wb') as token:
          pickle.dump(creds, token)
  return creds


def getPdfList(drive_service):
  results = drive_service.files().list(q="mimeType='application/pdf' and parents in '%s'" % FOLDER_ID, fields="nextPageToken, files(id, name)").execute()
  items = results.get('files', [])

  if not items:
      print('No files found.')
  else:
      print('Files:')
      for item in items:
          print(u'{0} ({1})'.format(item['name'], item['id']))
  return items

def downloadFile(drive_service, filename, fileId):
  request = drive_service.files().get_media(fileId=fileId)
  #fh = io.FileIO(filename, mode='wb')
  fh = io.BytesIO()
  downloader = MediaIoBaseDownload(fh, request)
  
  done = False
  while done is False:
    status, done = downloader.next_chunk()
    if status:
      logger.debug("Download %s %d%%." % (filename, int(status.progress() * 100)))
  return fh


def readPdf(filename, fh):
  logger.info("Reading file %s", filename);
  with pdfplumber.open(fh) as pdf:
    pages = pdf.pages
    text = ''
    for page in pages:
      text += "\n" + page.extract_text()
    return text

def readPdfList(drive_service):
  pdfList = getPdfList(drive_service);
  globalList = []
  for item in pdfList:
    filename = item['name']
    fileId = item['id']
    fh = downloadFile(drive_service, filename, fileId)
    text = readPdf(filename, fh)
    items = readItems(filename, text)
    globalList += items
    #print([str(elem) for elem in items])
  print([str(elem) for elem in globalList])
  return globalList;


def convertToNum(n1, n2): 
  return int(n1)*100 + int(n2);

def cleanUpName(n):
  return n.strip().replace('\xa0', " ")

def readItems(filename, text):
  s = 0;
  
  items = [];
  tvalue = 0;
  
  logger.debug(text)

  resDocN = re.search(reDocN, text, re.MULTILINE);
  if (resDocN):
    docN = resDocN.group(1)
    logger.info("Document n. %s", docN)
  else:
    logger.error("Cannot find doc N")

  resDate = re.search(reDate, text, re.MULTILINE);
  if (resDate):
    docDate = resDate.group(1)
    docTime = resDate.group(2)
  else:
    logger.error("Cannot find doc date/time")

  lines = text.split("\n")
  for line in lines:
    if (s == 0):
      if (re.search(reStart, line)):
        logger.info("Found start line, begin to read");
        s = 1;
    elif (s == 1):
      rtot = re.search(reTot, line);
      if (rtot):
        tvalue = convertToNum(rtot.group(1), rtot.group(2));   
        logger.info("Found tot line, stop processing; tot = %d ", tvalue)
        break;
      else:
        d = re.search(reDiscount, line);
        if (d):
          #logger.info("Found discount")
          dvalue = convertToNum(d.group(2), d.group(3));
          #logger.info("Found discount %d" + dvalue)
          items[len(items) -1].discount = dvalue;
          items[len(items) -1].tot -= dvalue;
        else:
          res = re.search(reLinePrice, line);
          if (res):
            logger.debug("line: >%s<", line)
            
            value = convertToNum(res.group(3), res.group(4));
            logger.debug("value = %d", value)
            itemObj = Item(filename, docN, docDate, docTime, cleanUpName(res.group(1)), value, res.group(2))
            items.append(itemObj);
          else:
            logger.info("skipped:>%s<", line)
  tcalc = sum(c.tot for c in items)
  logger.info("#Items %d - tvalue = %d - tot calc = %d", len(items), tvalue, tcalc)       
  return items;

def convertDate(dateStr):
  dateObj = datetime.strptime(dateStr, INPUT_DATE_FORMAT)
  return datetime.strftime(dateObj, OUTPUT_DATE_FORMAT)

def escapePlus(str):
  res = str
  if (res[0] == '+'):
    res = "'" + str
  return res

def writeGSheet(sheet_service, data) :
  values = []
  for d in data:
    values.append([d.filename, 
      d.receiptId, 
      convertDate(d.date), 
      d.time,
      escapePlus(d.name),
      d.price,
      d.tot,
      d.discount,
      d.vat])
  value_range_body = {
    'values': values
  }

  # Call the Sheets API
  sheet = sheet_service.spreadsheets()
  request = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                              range=RANGE_NAME,
                              valueInputOption='USER_ENTERED',
                              body=value_range_body)
  request.execute()




def main():
  creds = getCredentials()

  sheet_service = build('sheets', 'v4', credentials=creds)
  drive_service = build('drive', 'v3', credentials=creds)
  pdfList = readPdfList(drive_service)
  writeGSheet(sheet_service, pdfList)



if __name__ == '__main__':
  logging.config.fileConfig(fname='log.conf', disable_existing_loggers=False)

  # Get the logger specified in the file
  logger = logging.getLogger('sampleLogger')
  main()
'''
  args = sys.argv[1:];  
  itemList = readPdfList(args[0]);
  writeExcel(args[1], itemList)

  #handleArguments()
'''