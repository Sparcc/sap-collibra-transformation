from openpyxl import load_workbook
import sys, os

class SapDataParser:
    #TODO calculate this
    upperRange = 0
    currentInfoArea = ''
    currentTable = ''
    outputRowNumInfoArea = 2
    outputRowNumTable = 2
    outputRowNumColumn = 2
    sourceFileName = ''
    outputFileName = 'output.xlsx'
    hasInfoArea = False
    hasTable = False 
    domain = 'SAFYR SAP Test'
    community = 'Technical Metadata Community'
    domainType = 'Physical Data Dictionary'
    def __init__(self,input, output1, output2, output3):
        print('Loading Excel Files')
        self.sourceFileName = input
        self.wb = load_workbook(filename = input)
        self.ws = self.wb.active
        self.resetOutputFile(output)
        self.outputFileName = output
        
        self.outputIO = load_workbook(output1)
        self.sOutputIO = self.outputIO.active
        
        self.outputT = load_workbook(output2)
        self.sOutputT = self.outputT.active
        
        self.outputC = load_workbook(output3)
        self.sOutputC = self.outputC.active
        
        self.buildFieldMap()
        self.buildHeaders()
    def resetOutputFile(self, fileName1 = 'output1.xlsx', fileName2 = 'output2.xlsx', fileName3):
        fileName[0] = 
        pathName = '.\\emptyOutput\\' + fileName
        destination = '.\\'
        for i in range(0,3)
            
            os.system('del {d}\{fn}'.format(d=destination,fn=fileName))
            os.system('copy {pn} {d}"'.format(pn=pathName,d=destination))
    def buildHeaders(self):    
        for k,v in self.fieldTemp.items():
            self.sOutput[v+'1'] = k
    def setDataLength(self, upperRange = 49646):
        #TODO: Calculate upperRange
        self.upperRange = upperRange + 1
    def buildFieldMap(self):
        print('Building Map')
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
        
        tb1 = load_workbook(filename = 'info-area-template.xlsx')
        ts1 = tb1.active
        self.fieldTemp={}
        for col in ts1.iter_cols():
            self.fieldTemp1[col[0].value] = col[0].column
            
        tb2 = load_workbook(filename = 'table-template.xlsx')
        ts2 = tb2.active
        self.fieldTemp={}
        for col in ts2.iter_cols():
            self.fieldTemp2[col[0].value] = col[0].column
            
        tb3 = load_workbook(filename = 'column-template.xlsx')
        ts3 = tb3.active
        self.fieldTemp={}
        for col in ts3.iter_cols():
            self.fieldTemp3[col[0].value] = col[0].column
            
        
            
    def start(self, startingRow = 2, limit = 0):
        print('Starting transformation')
        rowNum = startingRow
        temp = limit
        if temp >0:
            self.upperRange = limit
        while rowNum < self.upperRange:
            self.processRow(rowNum)
            rowNum+=1
        print('Complete! Saving File...')
        self.output.save(self.outputFileName)
        print('Done')
    def processRow(self,rowNum):
        print('Current info area={info}'.format(info=self.currentInfoArea))
        #check for change in info area
        data = self.ws[self.fieldSrc['Child']+str(rowNum)].value
        data = str(data)
        if data == '' or data is None: #initial state or no info area
            self.hasInfoArea = False
        elif data != self.currentInfoArea: #not equal to current info area, change detected
            self.currentInfoArea = data
            self.createNewInfoArea(rowNum)
            self.createNewInfoArea(rowNum,isChild=True)
            self.hasInfoArea = True
            
        #check for change in table
        data = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)].value
        data = str(data)
        print('TABLE = '+data)
        if data == '' or data is None: #initial state or no table
            print('No more of this table exists')
            self.hasTable = False
        elif data !=self.currentTable: #mismatch detected
            self.hasTable = True
            self.currentTable = data
            self.createNewTable(rowNum)
        
        #build columns
        if data != '' and data is not None and self.hasTable and self.ws[self.fieldSrc['DD_FIELDNAME']+str(rowNum)].value is not None: #columns must be put in a table
            print('Creating column')
            self.createNewColumn(rowNum)
        '''
            print("Creating column but not checking if has table yet")
            if self.hasTable and self.ws[self.fieldSrc['DD_FIELDNAME']+str(rowNum)].value is not None:
                print('Creating column')
                self.createNewColumn(rowNum)
        '''

           
    def createNewTable(self,rowNum):
        self.sOutput2[self.fieldTempTable['Name']+str(self.outputRowNumTable)] = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)].value
        self.sOutput2[self.fieldTempTable['Status']+str(self.outputRowNumTable)] = 'Candidate'
        self.sOutput2[self.fieldTempTable['Type']+str(self.outputRowNumTable)] = 'Table'
        self.sOutput2[self.fieldTempTable['Domain']+str(self.outputRowNumTable)] = self.domain
        self.sOutput2[self.fieldTempTable['Community']+str(self.outputRowNumTable)] = self.community
        self.sOutput2[self.fieldTempTable['Domain Type']+str(self.outputRowNumTable)] = self.domainType
        
        self.sOutput[self.fieldTemp['Table Type']+str(self.outputRowNumTable)] = self.ws[self.fieldSrc['DD_TABLETYPE']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Description']+str(self.outputRowNumTable)] = self.ws[self.fieldSrc['LONG_DESC']+str(rowNum)].value
        
        #relation (sometimes has no info area)
        if self.hasInfoArea:
            self.sOutput2[self.fieldTempTable['is captured in [Info Area] > Info Area']+str(self.outputRowNumTable)] = self.ws[self.fieldSrc['Child']+str(rowNum)].value
            self.sOutput2[self.fieldTempTable['is captured in [Info Area] > Type']+str(self.outputRowNumTable)] = 'Info Area'
            self.sOutput2[self.fieldTempTable['is captured in [Info Area] > Community']+str(self.outputRowNumTable)] = self.community
            self.sOutput2[self.fieldTempTable['is captured in [Info Area] > Domain Type']+str(self.outputRowNumTable)] = self.domainType
            self.sOutput2[self.fieldTempTable['is captured in [Info Area] > Domain']+str(self.outputRowNumTable)] = self.domain
        
        self.outputRowNumTable +=1
     
    def createNewColumn(self,rowNum):
        self.sOutput[self.fieldTemp['Name']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_FIELDNAME']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Status']+str(self.outputRowNum)] = 'Candidate'
        self.sOutput[self.fieldTemp['Type']+str(self.outputRowNum)] = 'Column'
        self.sOutput[self.fieldTemp['Domain']+str(self.outputRowNum)] = self.domain
        self.sOutput[self.fieldTemp['Community']+str(self.outputRowNum)] = self.community
        self.sOutput[self.fieldTemp['Domain Type']+str(self.outputRowNum)] = self.domainType
        
        #attributes
        self.sOutput[self.fieldTemp['Is Nullable']+str(self.outputRowNum)] = self.convertToCommonTerm(self.ws[self.fieldSrc['MANDATORY']+str(rowNum)].value)
        self.sOutput[self.fieldTemp['Description']+str(self.outputRowNum)] = self.ws[self.fieldSrc['SHORT_DESC']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Is Primary Key']+str(self.outputRowNum)] = self.convertToCommonTerm(self.ws[self.fieldSrc['KEY_FLAG']+str(rowNum)].value)
        self.sOutput[self.fieldTemp['Number of Fractional Digits']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DATA_DECIMALS']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Size']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DATA_LENGTH']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Column Position']+str(self.outputRowNum)] = self.ws[self.fieldSrc['POSIT']+str(rowNum)].value
        self.sOutput[self.fieldTemp['Technical Data Type']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_DATATYPE_ERP']+str(rowNum)].value
        
        #relation
        self.sOutput[self.fieldTemp['is part of [Table] > Table']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)].value
        self.sOutput[self.fieldTemp['is part of [Table] > Type']+str(self.outputRowNum)] = 'Table'
        self.sOutput[self.fieldTemp['is part of [Table] > Community']+str(self.outputRowNum)] = self.community
        self.sOutput[self.fieldTemp['is part of [Table] > Domain Type']+str(self.outputRowNum)] = self.domainType
        self.sOutput[self.fieldTemp['is part of [Table] > Domain']+str(self.outputRowNum)] = self.domain
            
        self.outputRowNum +=1

    def createNewInfoArea(self,rowNum, isChild = False):
        self.sOutput[self.fieldTemp['Status']+str(self.outputRowNum)] = 'Candidate'
        self.sOutput[self.fieldTemp['Type']+str(self.outputRowNum)] = 'Info Area'
        self.sOutput[self.fieldTemp['Domain']+str(self.outputRowNum)] = self.domain
        self.sOutput[self.fieldTemp['Community']+str(self.outputRowNum)] = self.community
        self.sOutput[self.fieldTemp['Domain Type']+str(self.outputRowNum)] = self.domainType

        if isChild:
            self.sOutput[self.fieldTemp['Name']+str(self.outputRowNum)] = self.ws[self.fieldSrc['Child']+str(rowNum)].value
        
            #relation to parent
            self.sOutput[self.fieldTemp['is a child of [Info Area] > Info Area']+str(self.outputRowNum)] = self.ws[self.fieldSrc['Parent']+str(rowNum)].value
            self.sOutput[self.fieldTemp['is a child of [Info Area] > Type']+str(self.outputRowNum)] = 'Info Area'
            self.sOutput[self.fieldTemp['is a child of [Info Area] > Community']+str(self.outputRowNum)] = self.community
            self.sOutput[self.fieldTemp['is a child of [Info Area] > Domain Type']+str(self.outputRowNum)] = self.domainType
            self.sOutput[self.fieldTemp['is a child of [Info Area] > Domain']+str(self.outputRowNum)] = self.domain
            
            #relation to table #not needed since table relates to info area
            '''
            if self.hasTable:
                print('This info area [{a}] has a table [{b}]'.format(a=self.ws[self.fieldSrc['Child']+str(rowNum)].value,b=self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)].value))
                self.sOutput[self.fieldTemp['captures [Table] > Table']+str(self.outputRowNum)] = self.ws[self.fieldSrc['DD_TABLENAME']+str(rowNum)].value
                self.sOutput[self.fieldTemp['captures [Table] > Type']+str(self.outputRowNum)] = 'Table'
                self.sOutput[self.fieldTemp['captures [Table] > Community']+str(self.outputRowNum)] = self.community
                self.sOutput[self.fieldTemp['captures [Table] > Domain Type']+str(self.outputRowNum)] = self.domainType
                self.sOutput[self.fieldTemp['captures [Table] > Domain']+str(self.outputRowNum)] = self.domain
            '''
        else:
            self.sOutput[self.fieldTemp['Name']+str(self.outputRowNum)] = self.ws[self.fieldSrc['Parent']+str(rowNum)].value
        self.outputRowNum +=1
        
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