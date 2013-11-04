import win32com.client
import time
class ExcelAction(object):
    """attach to excel and perform necessary actions
    """
    appList = [] # here goes, for all instances, the excel application object
    positions = {} # dict with book/sheet/(col, row) (tuple)
    rows = {} # dict with book/sheet/row  (str)
    columns = {} # dict with book/sheet/col (str)
    def __init__(self):
        self.app = self.connect()
        if not self.app:
            raise Exception("ExcelAction, cannot connect to Excel: %s"% self.app)
        self.prevBook = self.prevSheet = self.prevPosition = None
        
    def checkForChanges(self):
        """return 1 if book, sheet or position has changed since previous call
        """
        changed = 0
        if not self.app:
            self.prevBook = self.prevSheet = self.prevPosition = None
            self.connect()
            if not self.app:
                raise Exception("ExcelAction, checkForChanges, cannot connect to Excel: %s"% self.app)
                
        self.book = self.app.ActiveWorkbook
        self.bookName = str(self.book.Name)
        self.sheet = self.app.ActiveSheet
        self.sheetName = str(self.sheet.Name)
        if self.prevBook != self.bookName:
            self.prevBook = self.bookName
            changed += 4
        if self.prevSheet != self.sheetName:
            self.prevSheet = self.sheetName
            changed += 2
        if changed:
            self.positions.setdefault(self.bookName,{})
            self.positions[self.bookName].setdefault(self.sheetName, [])
            self.columns.setdefault(self.bookName,{})
            self.columns[self.bookName].setdefault(self.sheetName, [])
            self.rows.setdefault(self.bookName,{})
            self.rows[self.bookName].setdefault(self.sheetName, [])
            self.currentPositions = self.positions[self.bookName][self.sheetName]
            self.currentColumns = self.columns[self.bookName][self.sheetName]
            self.currentRows = self.rows[self.bookName][self.sheetName]
        cr = self.savePosition()
        if cr != self.prevPosition:
            self.prevPosition = cr
            changed += 1
        return changed
    
    #connect to programs:
    def connect(self):
        """connect to excel and leave app"""
        if self.appList:
            app = self.appList[0]
            if app and app.ActiveCell:
                return app
            else:
                del app
            print 'connection to excel not valid anymore, disconnect...'
            self.disconnect()
            time.sleep(1)
            print 'and reconnect...'
        app = win32com.client.Dispatch('Excel.Application')
        print 'python is now connected to excel'
        self.appList.append(app)
        return app

    def disconnect(self):
        """disconnect from excel
        
        by removing the connecting 
        """
        # previous connection was not valid anymore, so remove old list and go for a new connection
        while self.appList:
            pop = self.appList.pop()
            del pop

    def savePosition(self):
        """save current position in positions dict
        
        returns the current (col, row) tuple
        """
        cr = self.getCurrentPosition()
        c, r = cr
        self.pushToListIfDifferent(self.currentPositions, cr)
        self.pushToListIfDifferent(self.currentRows, r)
        self.pushToListIfDifferent(self.currentColumns, c)
        
        return cr
    
    def getBooksList(self):
        """get list of strings, the names of the open workbooks
        """
        books = []
        for i in range(self.app.Workbooks.Count):
            b = self.app.Workbooks(i+1)
            print 'b: name: %s'% b.Name
            books.append(str(b.Name))
        return books
    
    def getSheetsList(self, book=None):
        """get list of strings, the names of the open workbooks
        """
        if book is None:
            book = self.book
        sheets = []
        for i in range(book.Sheets.Count):
            s = self.app.Worksheets(i+1)
            sheets.append(str(s.Name))
        return sheets
    
    def selectSheet(self, sheet):
        """select the sheet by name of number
        """
        self.app.Sheets(sheet).Activate()
    
    def getCurrentPosition(self):
        """return row and col of activecell
        
        as a side effect remember the (changed position)
        """
        if not self.app:
            return None, None
        ac = self.app.ActiveCell
        comingFrom = ac.Address
        cr = [str(value).lower() for value in comingFrom.split("$") if value]
        if len(cr) == 2:
            return tuple(cr)
        else:
            return None, None

    def pushToListIfDifferent(self, List, value):
        """add to list (in place) of value differs from last value in list
        
        for positions, value is (c,r) tuple
        """
        if not value:
            return
        if List and List[-1] == value:
            return
        List.append(value)
    
    def getPreviousRow(self):
        """return the previous row number
        """
        cr = self.getCurrentPosition()
        c, r = cr
        while 1:
            newr = self.popFromList(self.currentRows)
            if newr is None:
                return
            if r != newr:
                return newr

    def getPreviousColumn(self):
        """return the previous col letter
        """
        cr = self.getCurrentPosition()
        c, r = cr
        while 1:
            newc = self.popFromList(self.currentColumns)
            if newc is None:
                return
            if c != newc:
                return newc
    
    def popFromList(self, List):
        """pop from list a different value than currentValue
        
        and return None if List is exhausted
        """
        if List:
            value = List.pop()
            return value


# functions that do an action from the action.py module, in case of excel:
# one parameter should be given    
def gotoline(rowNum):
    """overrule for gotoline meta-action
    
    goto line in the current column
    """
    rowNum = str(rowNum)
    xls = ExcelAction()
    if not xls.app:
        return
    cPrev, rPrev = xls.getCurrentPosition()
    if rowNum == rPrev:
        print 'row already selected'
        return
    range = cPrev + rowNum
    #print 'current range: %s, %s'% (rPrev, cPrev)
    sheet = xls.app.ActiveSheet
    #print 'app: %s, sheet: %s (%s), range: %s'% (app, sheet, sheet.Name, range)
    sheet.Range(range).Select()
        
def selectline(dummy=None):
    """select the current line
    """
    xls = ExcelAction()
    if not xls.app:
        return
    xls.checkForChanges()
    xls.app.ActiveCell.EntireRow.Select()
        
def lineback(dummy=None):
    """goes back to the previous row
    """
    xls = ExcelAction()
    if not xls.app:
        return
    xls.checkForChanges()
    xls.app.ActiveCell.EntireRow.Select()
    prevRow = xls.getPreviousRow()
    print 'prevRow: %s'% prevRow
    if prevRow:                
        xls.gotoRow(prevRow)    