"""
    NOTE on google sheets format:
        
    The google sheets document must have the following:
    Columns A and B from cell 2 down are all the sample details
    Column A is the name of the detail and Column B is the value of the detail
    Columns B and onwards can be anything but there must be three speicifc columns
    Replicate, Test Name, Result
    Each row in these columns is counted as a test result and are grouped together
    by test name, replicate
    All rows must have a unique TestName-Replicate combination or an error is shown

"""
from SampleData import SampleData
from SRGJob import SRGJob
import time

class GoogleSheetsJobParser:
    """ Opens a google sheets document and parses the contents into a job class """
    
    def __init__(self, view):
        self.view = view
        pass
    

    
    def parse_document(self, service, document_id):
        """ The main function that opens the document, parses the data into a jobs
        object and returns the resulting job with all the calculated values
        
        Args:
            service (google sheets service): the googlse sheets api service
            document_id (str): the google sheets document id to fetch and process
            
        Returns:
            job (SRGJob): the job object containing all the job information
                and calculated results from the data
        
        """
        # Call the Sheets API to get a reference to the sheet    
        sheet_ref = service.spreadsheets().get(spreadsheetId=document_id)
        
        #get the sheet details such as individual sheet names, each test sample
        #will be on a separate sheet
        sheet_details = sheet_ref.execute()
   
        #The job object will hold the list of samples and their data
        job = SRGJob()
        
        #go through each sheet in the spreadsheets and process into a SampleData object
        for sheet in sheet_details.get('sheets'):
                    
            title = sheet.get('properties').get('title')
            self.view.display_message("Processing: {}".format(title))
            
            #special Details tab is used to extract the details required for the report
            if title == "Details":
                self.parse_details(sheet, job, service, document_id)
            else:
                self.parse_sample(sheet, job, service, document_id)
                
            #slow down the requests as to not brech the x requets in 100 seconds
            time.sleep(10)
            
    
        #return None if no samples were added to this job 
        if len(job.samples) > 0:
            return job
        else:
            return None

    def parse_details(self, sheet, job, service, document_id):
        """ Parses the details tab which has information about the report
        
        Args:
            sheet (google sheet): The google sheet to process
            job (SRGJob): the pointer to the job object to hold all the results
            service (google sheets service): the google sheets api service
            document_id (str): the google sheets document id to fetch and process
        
        """
        
        #get the first row with all the column headings as well as the first two
        #columns which contain the sample details
        result = service.spreadsheets().values().get(spreadsheetId=document_id,
                                range='Details!A2:B101').execute()
        
        #fields columns, first column is name of field and second 
        #column is the value for the field
        values = result.get('values', [])
        for row in values:
            if len(row) == 2 and row[0] != '':
                job.fields[row[0]] = row[1]
            
        
    
    def parse_sample(self, sheet, job, service, document_id):
        """ Parses each tab in the sheet as a separate sample.
        
        Args:
            sheet (google sheet): The google sheet to process
            job (SRGJob): the pointer to the job object to hold all the results
            service (google sheets service): the google sheets api service
            document_id (str): the google sheets document id to fetch and process

        
        """
        
        #get the title of the sheet to be used for the ranges reference in batchGet()
        title = sheet.get('properties').get('title')
            
        #rowCount is taken here so we know how many rows to extract from the sheet
        row_count = sheet.get('properties').get('gridProperties').get('rowCount')
        
        #get the first row with all the column headings as well as the first two
        #columns which contain the sample details
        result = service.spreadsheets().values().batchGet(spreadsheetId=document_id,
                                ranges=['{0}!A2:B101'.format(title),
                                        '{0}!A1:Z1'.format(title)]).execute()
        
        #The first element in this array is the sample details columns
        #the second element is the column names ie. first row of sheet
        valueRanges = result.get('valueRanges', [])
        
        #create a sample data object to store all the extracted data
        sample_data = SampleData()

        #Sample details columns, first column is name of detail and second 
        #column is the value for the detail
        values = valueRanges[0].get('values', [])
        for row in values:
            if len(row) == 2 and row[0] != '':
                sample_data.add_detail(row[0], row[1])
            
        #names of all the columns in the spreadsheet. Need to get the index
        #of the Test Name and Result columns, these are the data columns that
        #need to be extracted from the sheet
        values = valueRanges[1].get('values', [])
        try:
            tn_col_index = values[0].index("Test Name")
            res_col_index = values[0].index("Result")
        except ValueError:
            #don't add this sample because the required columns did not exist
            return
        
        #convert the index of these columns to the column letter, columns
        #are letters in spreadsheets not numbers
        alpha_codes = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"            
        tn_code = alpha_codes[tn_col_index]
        res_code = alpha_codes[res_col_index]
        
        #make another request to the sheets api to get the test result data
        #from the Test Name and Result columns found above
        data_result = service.spreadsheets().values().batchGet(spreadsheetId=document_id,
                                ranges=['{0}!{1}2:{1}{2}'.format(title, tn_code, row_count),
                                        '{0}!{1}2:{1}{2}'.format(title, res_code, row_count)]).execute()
        
        #The first element in data_values will be the Test Name column array
        #the second element will be the Result column array
        data_values = data_result.get('valueRanges', [])            
        tn_data = data_values[0].get('values', [])
        res_data = data_values[1].get('values', [])
        
        #go through each row in the extracted data and get the value for
        #Test Name and Result
        for i in range(len(tn_data)):
            #Add the Result for this Test Name to the sample_data test result array
            if len(tn_data[i]) > 0 and len(res_data[i]) > 0:
                sample_data.add_result(tn_data[i][0], res_data[i][0])    

        #if this sample had some useable data then add it to the job object
        if len(sample_data.details) > 0 and len(sample_data.test_results) > 0:
           job.add_sample(sample_data) 
    
    