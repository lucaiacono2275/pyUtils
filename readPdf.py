import pdfplumber;
import sys;

import logging
import logging.config

def readPdf(filename):
  logger.info("Reading file %s", filename)
  #fh = open(filename, "r", encoding = "ISO-8859-1")
  
  with pdfplumber.open(filename) as pdf:
    pages = pdf.pages
    text = ''
    for page in pages:
      text += "\n" + page.extract_text()
    return text

def main():
  filename = sys.argv[1]
  logger.info("file=%s", filename)
  text = readPdf(filename)
  print(text)

if __name__ == '__main__':
  logging.config.fileConfig(fname='log.conf', disable_existing_loggers=False)

  # Get the logger specified in the file
  logger = logging.getLogger('sampleLogger')
  main()  