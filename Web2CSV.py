#-------------------------------------------------------------------------------
# Name:        Web2CSV.py
# Purpose:
#
# Author:      Stefan Dittforth
#
# Created:     20.10.2011
# Version history:
#
#   1.0 - 21.10.2011    Initial version, transfer of Barclays code to Python
#   1.1 - 19.11.2011    Fix: add capability to read two line transactions
#   1.2 - 28.12.2011    Fix: print empty line in Barclays transactions with no
#                       "Verwendungszweck"
#   1.3 - 11.02.2012    Fix: remove '&amp;' from HSBC transaction details text
#   1.4 - 25.12.2012    Fix: added detection of HSBC UAE eSaver Accounts
#   2.0 - 28.08.2012    Fixes after HSBC webiste chnage, Refactored for multiple
#                       CSV file merge, code migrated to Python 3.6
#-------------------------------------------------------------------------------
#!/usr/bin/env python

from win32com.client import Dispatch
import re
from datetime import datetime
from bs4 import BeautifulSoup
import csv

class BarclaysAccount:
    """Class for interpreting Barclays account statements."""

    def __init__(self, ie):
        self.type = 'Barclays Account'
        self.ie = ie
        self.fileName = ''

    def recognise(self):
        """Method to detect whether the current website contains a Barclays \
           account table"""

        r = False

        # Criterias for identifying the Barclays account statement website:
        # 1.) find string "Barclays Online Banking" within the <head> block
        # 2.) the 5th <table> in the html document has a summary attribute
        #     with value "Summary view of transactions"
        r = False
        t1 = self.ie.document.head.getElementsByTagName('title')(0).innerText
        t2 = self.ie.document.body.getElementsByTagName('table')(4).summary
        if (t1.find('Barclays Online Banking') != -1) and \
           (t2.find('Summary view of transactions') != -1):
            r = True
        else:
            r = False

        return r

    def getIdentifier(self):
        """Method to extract a useful identifier for the data extracted from \
           the website. For Barclays accounts this is the account number and \
           sort code."""

        # Default value in case account number and sort code cannot be found.
        i = 'Baclays Account'
        # The default identifier is the Barclays account number and sort code.
        # These are stored in the 3rd <table> in the HTML document.
        tables = self.ie.document.body.getElementsByTagName('table')
        if tables.length > 2:
            t = tables(2).innerText
            m = re.search('\d{2}-\d{2}-\d{2} \d{8}', t)
            if m != None:
                # Found account number and sort code.
                i = m.group()
            else:
                # There is no string that matches the format of account number
                # sort code.
                print('Cannot find the account number and sort code for \
                       the Barclays account within the <table> tag.')
        else:
            # table with account number and sort code not found
            print('Cannot find the <table> tag with account number and sort \
                   code for the Barclays account.')

        return i

    def writeCSV(self):
        """ Method that reads the <table> with the account transactions and \
            writes them into a file in CSV format."""

        soup = BeautifulSoup(self.ie.document.body.innerHTML)
        # Extract the table that contains the transactions
        table = soup.find('table', {'summary':'Summary view of transactions'})
        rows = table.findAll('tr')

        # Apparently Beautiful Soup objects take up quite a bit of memory. It's
        # propably a good idea to delete the variable now that we no longer
        # need it.
        del soup, table

        # Open the CSV file
        f = open(self.fileName, 'wb')
        CSVFile = csv.writer(f, delimiter=';')
        # Write CSV header
        CSVFile.writerow(['Wertstellung', 'Buchungsdatum', 'Empfaengername', 'Verwendungszweck', 'Betrag'])

        txnWertstellung = ''
        txnBuchungsdatum = ''
        txnBuchungsart = ''
        txnSenderEmpfaenger = ''
        txnVerwendungszweck = ''
        txnBetrag = ''
        txnBalance = ''
        l = 1
        txnNo = 1
        for row in rows:

            cols = row.findAll('td')

            if len(cols) > 0: # this will make sure we skip the table header <th> tags
                c = []
                for col in cols:
                    # strip out Pound character & html string
                    t = col.string.replace(u'\xa3', '').strip()
                    t = t.replace('&nbsp;', '').strip()
                    # tidy up a bit and remove excess (more than one) spaces between words
                    t = re.sub('(\s{2,})', lambda matchobj: ' ', t)
                    c.append(t)

                if c[0] + c[1] + c[2] + c[3] + c[4] == '':
                    # we reached the last row of a transaction
                    l = 1 # reset the line clounter

                    # Print transaction details on console
                    print('Transaction No. ' + str(txnNo))
                    print
                    print('Wertstellung        : ' + txnWertstellung)
                    print('Buchungsdatum       : ' + txnBuchungsdatum)
                    print('Buchungsart         : ' + txnBuchungsart)
                    print('Sender / Empfaenger : ' + txnSenderEmpfaenger)
                    print('Verwendungszweck    : ' + txnVerwendungszweck)
                    print('Betrag              : ' + txnBetrag)
                    print('Balance             : ' + txnBalance)
                    print('------------------------------------------------')

                    # Write transaction to CSV file
                    CSVFile.writerow([txnWertstellung, txnBuchungsdatum, txnSenderEmpfaenger, txnVerwendungszweck, txnBetrag])

                    txnWertstellung = ''
                    txnBuchungsdatum = ''
                    txnBuchungsart = ''
                    txnSenderEmpfaenger = ''
                    txnVerwendungszweck = ''
                    txnBetrag = ''
                    txnBalance = ''

                    txnNo = txnNo + 1

                else:

                    # pull out the details from the various lines within a transaction
                    if l == 1:
                        txnWertstellung = c[0]
                        txnBuchungsdatum = c[0]
                        txnBuchungsart = c[1]
                        txnBetrag = self.__commaPoint(self.__getBetrag(c[2], c[3]))
                        txnBalance = self.__commaPoint(c[4])
                    elif l == 2:
                        txnSenderEmpfaenger = c[1]
                    elif l == 3:
                        txnVerwendungszweck = c[1]
                    elif l == 4:
                        txnVerwendungszweck = txnVerwendungszweck + ' ' + c[1]

                    l = l + 1 # next row of a transaction

        # Close the CSV file
        f.close()

    def __getBetrag(self, x, y):
        if x == '': b = y
        else: b = x
        return b

    def __commaPoint(self, s):
        s = s.replace(',', '')
        s = s.replace('.', ',')
        return s

class HSBCAccount:
    """Class for interpreting HSBC account statements."""

    def __init__(self, ie):
        self.type = 'HSBC'
        self.ie = ie
        self.fileName = ''

    def recognise(self):
        """Method to detect whether the current website contains a HSBC \
           account table"""

        # Criterias for identifying the HSBC account statement website:
        # 1.) find string " HSBC UAE - Internet Banking - Account History " in the title
        # 2.) The account <table> in the html document has a class attribute
        #     with value "hsbcTableStyle07"
        r = False
        t1 = self.ie.document.head.getElementsByTagName('title')(0).innerText
        tables = self.ie.document.body.getElementsByClassName('hsbcTableStyle07')
        if t1.find('HSBC UAE - Internet Banking - Account History') != -1 and \
           tables(0) != None:
            r = True
        else:
            r = False

        return r

    def getIdentifier(self):
        """Method to extract a useful identifier for the data extracted from \
           the website. For HSBC accounts this is the account number """

        # Default value in case account number cannot be found.
        i = 'HSBC Account'
        # The default identifier is the HSBC account or credit card number.
        # These are stored in a <span> tag with id = "LongSelection1Output".
        t = ''
        t = self.ie.document.getElementByID('LongSelection1Output').innerText
        if len(t) > 0:
            # Test if string is an account number or a credit card number.
            m = re.search('\d{3}-\d{6}-\d{3}|(\d{4}-){3}\d{4}', t)
            if m != None:
                # found number
                i = m.group()
            else:
                # There is no string that matches.
                print('Cannot find the account or credit card number for the HSBC account.')
            # The following is special for HSBC accounts. We pull out the type
            # of the account (current account or credit card). This is used in writeCSV()
            # method to read the account transaction table correctly.
            if t.find('HSBC PREMIER CARD') != -1: self.type = 'HSBC Premier Card'
            if t.find('CURRENT ACCOUNT') != -1: self.type = 'HSBC Current Account'
            if t.find('eSAVER ACCOUNT') != -1: self.type = 'HSBC Current Account'
        else:
            # HTML element with account number not found
            print('Cannot find the <span> tag with the account number or credit card number for the HSBC account.')

        return i

    def writeCSV(self):
        """ Method that reads the <table> with the account transactions and \
            writes them into a file in CSV format."""

        soup = BeautifulSoup(self.ie.document.body.innerHTML)
        # Extract the table that contains the transactions
        table = soup.find('table', {'class':'hsbcTableStyle07'})
        rows = table.findAll('tr', {'class':re.compile('hsbcTableRow03 hsbcTableRow05|hsbcTableRow04 hsbcTableRow05')})

        # Apparently Beautiful Soup objects take up quite a bit of memory. It's
        # propably a good idea to delete the variable now that we no longer
        # need it.
        del soup, table

        # read all html lines and build a list that contains all transactions
        txn = [] # transaction list
        for row in rows:

            cols = row.findAll('td')

            if len(cols) > 0: # this will make sure we skip the table header <th> tags
                c = [] # stores the values from the columns in the current row
                for col in cols:
                    if (col.string) < 1:
                        # For current accounts the second column contains the
                        # transaction details. These are multiple line separated
                        # by <br/> tags. We strip them out and concatenate to one line.
                        if self.type == 'HSBC Current Account':
                            t = col.renderContents().replace('<br />', ' ').strip()
                            t = t.replace('&amp;','').strip()
                        # For credit card accounts the last column is either empty
                        # or contains 'Cr' enclosed in <b> tag for CC balance payments.
                        if self.type == 'HSBC Premier Card':
                            t = col.text.strip()
                    else:
                        t = col.string.replace('&nbsp;', '').strip()
                        t = t.replace('&amp;', '').strip()
                    c.append(t)

                # Check if we are the transaction stretches over two lines
                if len(c) == 6: # double check we have 6 columns (each transaction line has 6 columns)
                    if c[0] + c[1] + c[2] + c[4] + c[5] == '':
                        # The second line only contains foreign currency information in column 4. We
                        # attach that information to the details text of column 3 of the previous line.
                        p = txn.pop() # get the information from the previous line
                        p[2] = p[2] + ' ' + c[3]
                        txn.append(p) # save amended transaction details
                    else:
                        txn.append(c)

        # Open the CSV file
        f = open(self.fileName, 'wb')
        CSVFile = csv.writer(f, delimiter=';')
        # Write CSV header
        CSVFile.writerow(['Wertstellung', 'Buchungsdatum', 'Verwendungszweck', 'Betrag'])
        # Write transaction details.
        txnNo = 1
        for p in txn:

            # current accounts
            if self.type == 'HSBC Current Account':
                # Reformat the date representation from '<name of month> <day>, <year>'
                # to '<day>/<month>/<year>'
                txnWertstellung = datetime.strptime(p[0], '%B %d, %Y').strftime('%d/%m/%Y')
                txnBuchungsdatum = txnWertstellung
                txnVerwendungszweck = p[1]
                txnBetrag = self.__commaPoint(self.__getBetrag(p[2], p[3]))
                txnBalance = self.__commaPoint(p[4])

            # credit cards
            if self.type == 'HSBC Premier Card':
                txnBuchungsdatum = datetime.strptime(p[0], '%B %d, %Y').strftime('%d/%m/%Y')
                txnWertstellung = datetime.strptime(p[1], '%B %d, %Y').strftime('%d/%m/%Y')
                txnVerwendungszweck = p[2]
                txnBetrag = self.__commaPoint(self.__getCCBetrag(p[4], p[5]))
                txnBalance = '' # CC statements don't show a balance

            # Print transaction details on console
            print('Transaction No. ' + str(txnNo))
            print
            print('Wertstellung        : ' + txnWertstellung)
            print('Buchungsdatum       : ' + txnBuchungsdatum)
            print('Verwendungszweck    : ' + txnVerwendungszweck)
            print('Betrag              : ' + txnBetrag)
            print('Balance             : ' + txnBalance)
            print('------------------------------------------------')

            # Write transaction to CSV file
            CSVFile.writerow([txnWertstellung, txnBuchungsdatum, txnVerwendungszweck, txnBetrag])

            txnWertstellung = ''
            txnBuchungsdatum = ''
            txnBuchungsart = ''
            txnSenderEmpfaenger = ''
            txnVerwendungszweck = ''
            txnBetrag = ''
            txnBalance = ''

            txnNo = txnNo + 1

        # Close the CSV file
        f.close()

    def __getBetrag(self, x, y):
        if x == '': b = y
        # Need to add the "-" sign to the debit values.
        else: b = '-' + x
        return b

    def __getCCBetrag(self, x, y):
        # If there is no 'Cr' shown in the last column of the account statement
        # then it's a debit transaction from the credit card account and we
        # need to add the '-' sign.
        if y == '':
            b = '-' + x
        else:
            b = x
        return b

    def __commaPoint(self, s):
        s = s.replace(',', '')
        s = s.replace('.', ',')
        return s

if __name__ == '__main__':

    # Start the web browser.
    ie = Dispatch('InternetExplorer.Application')
    ie.Visible = True

    # Build a list of interpreter objects
    interpreterObjectClasses = [BarclaysAccount, HSBCAccount]
    interpreters = []
    for interpreterObjectClass in interpreterObjectClasses:
        interpreterObject = interpreterObjectClass(ie)
        interpreters.append(interpreterObject)

    # Main loop
    i = ''
    while i.lower() != 'q':
        # Wait for user input before trying to read and covert the browser content.
        i = input('Please navigate to the website with the information to be converted to CSV and press \'Enter\'. To finish the program press \'q\'.')
        if i.lower() != 'q':
            for interpreter in interpreters:
                if interpreter.recognise() == True:
                    identifier = interpreter.getIdentifier()
                    CSVFileName = interpreter.type + ' ' + identifier + ' ' + \
                                  datetime.now().strftime('%Y-%m-%d %H-%M-%S') + \
                                  '.csv'
                    interpreter.fileName = CSVFileName
                    interpreter.writeCSV()
                    print
                    print('Transactions written to file '' + CSVFileName + ''')

    # Close the browser window.
    ie.quit()
