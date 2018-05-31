from openpyxl import load_workbook

class Column:
    name = ''
    table = ''
    
class Table:
    name = ''
    infoArea = ''
    tabletType = ''
    isNullable = ''
    description = ''
    isPrimaryKey = ''
    numFracDigs = ''
    size = ''
    colPos = ''
    techDataType = ''
    isCapturedIn = '' #can be blank
    
class InfoArea:
    childInfoArea = ''
    name = ''
    child = ''
    table = ''
    def __init__(self,name='',child=''):
        self.childInfoArea = childInfoArea

class SapDataParser:
    #TODO calculate this
    uppperRange = 0
    currentParentInfoArea = ''
    currentChildInfoArea = ''
    currentTable = ''
    def __init__(self,fName):
        self.wb = load_workbook(filename = fName)
        self.ws = self.wb.active
        self.output = load_workbook(filename = fName)
        self.outputSheet = self.output.active
        self.buildFieldMap()
        self.createNewRow(1)
    def setDataLength(self, upperRange = 49646):
        #TODO: Calculate upperRange
        self.upperRange = upperRange + 1
    def buildFieldMap(self):
        #src fields
        self.fieldSrc={}
        self.fieldSrc['Parent'] = 'A'
        self.fieldSrc['Child'] = 'B'
        self.fieldSrc['DD_TABLENAME'] = 'C'
        self.fieldSrc['LONG_DESC'] = 'D'
        self.fieldSrc['DD_TABLETYPE'] = 'E'
        self.fieldSrc['DD_FIELDNAME'] = 'F'
        self.fieldSrc['SHORT_DESC'] = 'G'
        self.fieldSrc['POSIT'] = 'H'
        self.fieldSrc['MANDATORY'] = 'I'
        self.fieldSrc['DD_DATATYPE_ERP'] = 'J'
        self.fieldSrc['DATA_LENGTH'] = 'K'
        self.fieldSrc['DATA_DECIMALS'] = 'L'
        self.fieldSrc['KEY_FLAG'] = 'M'
        
        tb = load_workbook(filename = 'template.xlsx')
        ts = tb.active
        self.fieldTemp={}
        for col in ts.iter_cols():
            self.fieldTemp[col[0].column] = col[0].value]
            #print(self.fieldTemp[col[0].column])
        '''
        self.fieldTemp={}
        self.fieldTemp['Name'] = ['A'
        self.fieldTemp['Status'] = 'B'
        self.fieldTemp['Type'] = 'C'
        self.fieldTemp['Domain'] = 'D'
        self.fieldTemp['Community'] = 'E'
        self.fieldTemp['Domain Type'] = 'F'
        self.fieldTemp[''] =
        self.fieldTemp['Table Type'] = 'G'
        self.fieldTemp['Is Nullable'] = 'H'
        self.fieldTemp['Description'] = 'I'
        self.fieldTemp['Is Primary Key'] = 'J'
        self.fieldTemp['Number of Fractional Digits'] = 'K'
        self.fieldTemp['Size'] = 'L'
        self.fieldTemp['Column Position'] = 'M'
        self.fieldTemp['Technical Data Type'] = 'N'
        self.fieldTemp['is a child of [Info Area] > Info Area'] = 'O'
        self.fieldTemp['is a child of [Info Area] > Type'] = 'P'
        self.fieldTemp[''] = 'Q'
        self.fieldTemp[''] = 'R'
        self.fieldTemp[''] = 'S'
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = 'T'
        self.fieldTemp[''] = 'U'
        self.fieldTemp[''] = 'V'
        self.fieldTemp[''] = 'W'
        self.fieldTemp[''] = 'X'
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        self.fieldTemp[''] = ''
        
        self.template['Table Type'] = 'G'
        self.template['Is Nullable'] = 'H'
        self.template['Description'] = 
        
        self.mapping={}
        self.mapping['schema'] = ['R','G']
        self.mapping['size'] = ['P','H']
        self.mapping['nullable'] = ['M','I']
        self.mapping['col_pos_id'] = ['F','J']
        self.mapping['frac_digs'] = ['O','K']
        self.mapping['default_value'] = ['I','L']
        self.mapping['desc'] = ['E','M']
        self.mapping['pk'] = ['N','N']
        '''
        '''
        self.template = {
                'Name':'A',
                'Status':'B',
                'Type':'C',
                'Domain':'D',
                'Community':'E',
                'Domain Type':'F'
                }
        '''
    def start(self, startingRow = 2):
        rowNum = startingRow
        while rowNum < numRows:
            self.processRow(rowNum)
            rowNum+=1
    def processRow(self,rowNum):
        #logic
        buildTable = False
        data = self.ws[fieldSrc['Parent']+str(rowNum)]
        if data != self.currentParentInfoArea and data != '':
            parentInfoArea = InfoArea()
            parentInfoArea.name = data # assign name
            self.currentParentInfoArea = infoArea.name
            #TODO: create new row for parent info area  
            if self.ws(fieldSrc['Child']+str(rowNum)] != self.currentChildInfoArea:
            childInfoArea = InfoArea()
            childInfoArea.name = name
            #TODO: create new row for child info area
        else:
            pa
        if self.ws[fieldSrc['Table']+str(rowNum)] !=self.currentTable:
           table = Table()
           #TODO: create new row for table
           
    def createNewRow(self,rowNum):
    
                
    def createNewInfoArea(self,rowNum):
    
    def createNewColumn(self,rowNum)
        #self.outputSheet[
        
    def convertToCommonTerm(self,v):
        v = str(v)
        returnValue = v
        for c in ('yes', 'true', 't', 'y'):
            if v.lower() == c:
                returnValue = 'True'
        for c in ('no', 'false', 'f', 'n'):
            if v.lower() == c:
                returnValue = 'False'
        for c in ("none",'none'): # there is a difference between ' and " !
            if v.lower() == c:
                returnValue = 'False'
        for c in ("",''):
            if v.lower() == c:
                returnValue = 'False'
        return returnValue