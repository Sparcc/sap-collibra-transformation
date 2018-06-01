from openpyxl import load_workbook

class SapDataParser:
    #TODO calculate this
    uppperRange = 0
    currentParentInfoArea = ''
    currentChildInfoArea = ''
    currentTable = ''
    outputRowNum = 0
    
    domain = 'SAFYR SAP Test'
    community = 'Technical Metadata Community'
    domainType = 'Physical Data Dictionary'
    def __init__(self,fName):
        self.wb = load_workbook(filename = fName)
        self.ws = self.wb.active
        self.output = load_workbook(filename = fName)
        self.sOutput = self.output.active
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
            self.fieldTemp[col[0].column] = col[0].value
            #print(self.fieldTemp[col[0].column])
            
    def start(self, startingRow = 2):
        rowNum = startingRow
        while rowNum < numRows:
            self.processRow(rowNum)
            rowNum+=1
            
    def processRow(self,rowNum):
        '''
        #check for change in info area
        data = self.ws[fieldSrc['Parent']+str(rowNum)].value
        if data != self.currentParentInfoArea and data != '':
            parentInfoArea = InfoArea()
            parentInfoArea.name = data # assign name
            self.currentParentInfoArea = infoArea.name
            #TODO: create new row for parent info area  
            if self.ws(fieldSrc['Child']+str(rowNum)] != self.currentChildInfoArea:
            childInfoArea = InfoArea()
            childInfoArea.name = name
            #TODO: create new row for child info area
        '''
        #check for change in table
        data = self.ws[fieldSrc['DD_TABLENAME']+str(rowNum)].value
        if data !=self.currentTable:
            self.currentTable = data
            self.createNewTable(rowNum)
           
        #build columns
        self.createNewColumn(rowNum)

           
    def createNewTable(self,rowNum, hasInfoArea = False):
        self.sOutput[fieldTemp['Name']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)]
        self.sOutput[fieldTemp['Status']+str(self.outputRowNum)] = 'Candidate'
        self.sOutput[fieldTemp['Type']+str(self.outputRowNum)] = 'Table'
        self.sOutput[fieldTemp['Domain']+str(self.outputRowNum)] = self.domain
        self.sOutput[fieldTemp['Community']+str(self.outputRowNum)] = self.domain
        self.sOutput[fieldTemp['Domain Type']+str(self.outputRowNum)] = self.domainType
        
        self.sOutput[fieldTemp['TableType']+str(self.outputRowNum)] = 'Table'
        self.sOutput[fieldTemp['Description']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_FIELDNAME']+str(rowNum)]
        
        #relation (sometimes has no info area)
        if hasInfoArea:
            self.sOutput[fieldTemp['is captured in [Info Area] > Info Area']+str(self.outputRowNum)] = self.ws[self.fieldSrc['Parent']+str(rowNum)]
            self.sOutput[fieldTemp['is captured in [Info Area] > Type']+str(self.outputRowNum)] = 'Info Area'
            self.sOutput[fieldTemp['is captured in [Info Area] > Community']+str(self.outputRowNum)] = self.community
            self.sOutput[fieldTemp['is captured in [Info Area] > Domain Type']+str(self.outputRowNum)] = self.domainType
            self.sOutput[fieldTemp['is captured in [Info Area] > Domain']+str(self.outputRowNum)] = self.domain
        
        self.outputRowNum +=1
     
    def createNewColumn(self,rowNum):
        self.sOutput[fieldTemp['Name']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_FIELDNAME']+str(rowNum)]
        self.sOutput[fieldTemp['Status']+str(self.outputRowNum)] = 'Candidate'
        self.sOutput[fieldTemp['Type']+str(self.outputRowNum)] = 'Column'
        self.sOutput[fieldTemp['Domain']+str(self.outputRowNum)] = self.domain
        self.sOutput[fieldTemp['Community']+str(self.outputRowNum)] = self.domain
        self.sOutput[fieldTemp['Domain Type']+str(self.outputRowNum)] = self.domainType
        
        #attributes
        self.sOutput[fieldTemp['Is Nullable']+str(self.outputRowNum)] = self.ws[self.fieldSrc['MANDATORY']+str(rowNum)]
        self.sOutput[fieldTemp['Description']+str(self.outputRowNum)] = self.ws[self.fieldSrc['LONG_DESC']+str(rowNum)]
        self.sOutput[fieldTemp['Is Primary Key']+str(self.outputRowNum)] = self.ws[self.fieldSrc['KEY_FLAG']+str(rowNum)]
        self.sOutput[fieldTemp['Number of Fractional Digits']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DATA_DECIMALS']+str(rowNum)]
        self.sOutput[fieldTemp['Size']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DATA_LENGTH']+str(rowNum)]
        self.sOutput[fieldTemp['Column Position']+str(self.outputRowNum)] = self.ws[self.fieldSrc['POSIT']+str(rowNum)]
        self.sOutput[fieldTemp['Technical Data Type']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_DATATYPE_ERP']+str(rowNum)]
        
        #relation
        self.sOutput[fieldTemp['is part of [Table] > Table']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)]
        self.sOutput[fieldTemp['is part of [Table] > Type']+str(self.outputRowNum)] = 'Table'
        self.sOutput[fieldTemp['is part of [Table] > Community']+str(self.outputRowNum)] = self.community
        self.sOutput[fieldTemp['is part of [Table] > Domain Type']+str(self.outputRowNum)] = self.domainType
        self.sOutput[fieldTemp['is part of [Table] > Domain']+str(self.outputRowNum)] = self.domain
            
        self.outputRowNum +=1

    def createNewInfoArea(self,rowNum, child = False):
        if child:
            print('do stuff')
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