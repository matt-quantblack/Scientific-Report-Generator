from __future__ import print_function
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google.auth import exceptions
from apiclient import errors
import time
import datetime
from GoogleSheetsJobParser import GoogleSheetsJobParser
from ResultsTableBuilder import ResultsTableBuilder
from MicrosoftDocxParser import MicrosoftDocxParser
import os

#Scope to give full access to the google drive account
SCOPES = ['https://www.googleapis.com/auth/drive']
#Service account credentials - Use on secure server only
DEFAULT_CREDENTIALS_FILE = 'credentials.json'
#Seconds between each poll for document changes
POLL_TIME = 20

class SRGController:
    """ Controller for the Scientific Report Generator """
    
    def __init__(self, view):
        """ Init function for the controller
        
        Args:
            view (class): A view class that contains the appropriate functions for a 
            SRGView        

        Attributes:
            view (class): A view class that contains the appropriate functions for a 
                SRGView
            service (googleapiclient.discovery.build): google drive service will all 
                the functions required for accessing the drive
            sheets_service (googleapiclient.discovery.build): google sheets service will all 
                the functions required for accessing google sheets
            docs_service (googleapiclient.discovery.build): google docs service will all 
                the functions required for accessing google docs
            permission_id (str): the permission id for the service account used
                for adding and removing permissions to the files
            session_id (str): A unique id associated with this background process
            team_drive_id (string): The id of the google team drives used
        """
        self.view = view  
        self.service = None
        self.sheets_service = None
        self.docs_service = None
        self.permission_id = None
        self.session_id = None
        self.team_drive_id = None

    def full_path(self, filename):
        """ Gets the full path of the passed filename
        
        Args:
            filename (str): the filename to get the path of
            
        Returns:
            str: The full path of the filename
        """
        return os.path.join(os.path.dirname(os.path.realpath(__file__)), filename)
    
    
    
    def start(self):
        """ The main loop searches for documents starting with 'PROCESS' every 20 seconds.
        This is the unique keyword used to activate a document for processing
        If a new sheets document has been found the sheet is parsed and 
        a report produced. The report is shared back with the original user
        """
        
        
        #Create the service credentials for the google drive api
        self.create_service(self.full_path(DEFAULT_CREDENTIALS_FILE))
   
        #use a time stamp to keep track of the thread session
        self.session_id = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        
        #create a session lock file required to keep the threads loop
        #running. Delete this file to stop the thread
        f=open(self.full_path("session-" + self.session_id + ".lock"), "w+")
        f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"))
        f.close()
        
        #run the main program loop
        self.main_loop()

    def create_service(self, cred_file):
        """ Creates the google api services with the passed credentials file
        this service is then used throughout the lifetime of this object
        
        Args:
            cred_file (str): the location of the credentials file
        
        Returns:
            bool: True if the services was created sucessfully
        """
        
        try:
            creds = service_account.Credentials.from_service_account_file(cred_file, scopes=SCOPES)
        except FileNotFoundError:
            self.display_error("Can't load credentials {0} is missing.".format(cred_file))
            return False
        except ValueError:
            self.display_error("Can't load credentials {0} is an invalid format.".format(cred_file))
            return False
            

        #Create the google api service objects
        self.service = build('drive', 'v3', credentials=creds)        
        self.sheets_service = build('sheets', 'v4', credentials=creds)
        self.docs_service = build('docs', 'v1', credentials=creds)
        
        #get this service account permissions id - used for adding and removing permissions
        try:
            response = self.service.about().get(fields="user").execute()
        except exceptions.RefreshError:
            self.display_error("Service account email not valid in {0}".format(cred_file))
            return False
        
        self.permission_id = response.get('user').get('permissionId')
        
        #get the team drive id if one is shared with this account
        try:
            response = self.service.teamdrives().list(pageToken=None).execute()
        except errors.HttpError:
            self.display_error("Could not find team drives. Did you share the folder with the service account email instead of the personal email address? Service account email needs access.")
            return False
        
        team_drives = response.get('teamDrives', [])
        if len(team_drives) > 0:
            self.team_drive_id = team_drives[0].get('id')
        
        return True
    

    def main_loop(self):
        """ Main loop that searches for documents starting with the keyword 
        'PROCESS'        
        """
        
        self.display_message("SRG session " + self.session_id + " Started.")
        print("SRG session " + self.session_id + " Started.")

        #This is the main loop so keep looping on this thread until the program
        #has been marked as not active by deleting the session.lock file
        while os.path.isfile(self.full_path("session-" + self.session_id + ".lock")):        

            #different calls are rquired depending if team drives are being used
            #find files with the keyword 'PROCESS ' that are of type google sheets            
            if self.team_drive_id is None:                    
                    results = self.service.files().list(
                            q="name contains 'PROCESS ' and mimeType='application/vnd.google-apps.spreadsheet'",
                            spaces='drive',
                            pageSize=100,
                            fields="files(id, name, mimeType)").execute()
            else:
                results = self.service.files().list(
                    q="name contains 'PROCESS ' and mimeType='application/vnd.google-apps.spreadsheet'",
                    spaces='drive',
                    corpora='teamDrive',
                    supportsTeamDrives=True,
                    includeTeamDriveItems = True,
                    teamDriveId=self.team_drive_id,
                    pageSize=100,
                    fields="files(id, name, mimeType)").execute()
      
            #check the details of each file found
            for file in results.get('files'):
                
                #break the loop if the session is shutdown
                if not os.path.isfile(self.full_path("session-" + self.session_id + ".lock")):
                    break                                
                                        
                self.display_message("PROCESS command found for file: {0}".format(file.get('name')))
                    
                #process the google sheets document into a job
                try:                            
                    self.process_job(file)                        
                #Just catch all errors here and log them to ensure the main
                #loop continues to run without crashing
                except Exception as e:
                    self.display_error(e.__str__())
                    self.display_error("Could not process job " + file.get('name'))   
                    
            

            #display a status and sleep thread until the next poll
            now_string = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.display_status("Last poll: {0}".format(now_string))
            time.sleep(POLL_TIME)
            
        #main loop has exited so display the session stopped message
        self.display_message("SRG session " + self.session_id + " Stopped.")
        print("SRG session " + self.session_id + " Stopped.")
    
    def process_job(self, file):
        """ Completes all required tasks to process the google sheet into a 
        finished test report.
        Parses the sheet into product results compiled into a job file
        Calcualtes all the results, parses the results to the Microsoft Doc template
        to fill in the details.
        
        Args:
            file (googlesheets.file): the file that was found to process
        """
        
        
        file_id = file.get('id')
        sheet_name = file.get('name')
        
        self.display_message("Spreedsheet {0} is being processed.".format(sheet_name))             
        
        #Create a sheet parser to generate a Job with a results collection
        sheets_parser = GoogleSheetsJobParser()

        #Run sheets parser and get the results collection in a job object
        job = sheets_parser.parse_document(self.sheets_service, file_id)
        
        #remove the unique key PROCESS from the filename now the file has been processed
        self.service.files().update(fileId=file_id, body={'name': sheet_name.replace('PROCESS ', '')}, supportsTeamDrives=True).execute()
        
        
        if job is not None:
            self.display_message("Spreedsheet {0} has parsed {1} products successfully.".format(sheet_name, len(job.samples)))  
            
            #need to check all the required details are in the job
            missing_data = ""
            if 'ReportTemplate' not in job.fields:
                missing_data += " ReportTemplate"
            if  'UploadFilename' not in job.fields:
                missing_data += " UploadFilename"
            if  'ShareWith' not in job.fields:
                missing_data += " ShareWith"
            if len(missing_data) > 0:
                self.display_error("Data sheet is missing " + missing_data + " in the Details tab")
                return
                        
            name = job.fields['ReportTemplate']
            new_name = job.fields['UploadFilename']
            
            #create the MicrosofDocxParser to parse the template daocument
            doc_parser = MicrosoftDocxParser()
            
            #download the template file found in the fields dictionary
            success = False
            try:
                success = doc_parser.download_report_template(self.service, name, self.full_path(name), self.team_drive_id)
            except IOError: 
                self.display_error("Save template error. Can't save to " + name)
                return

            if success:
                self.display_message("Downloaded template " + name)
                
                #get the table commands so we can build the required tables from the data
                table_commands = doc_parser.extract_table_commands(self.full_path(name))            
                
                #build the tables
                table_builder = ResultsTableBuilder()
                try:
                    tables = table_builder.create_tables(table_commands, job)
                except ValueError:
                    self.display_error("Could not build the results tables")
                    return
                
                #generate the word document now that all the data is ready to insert
                try:
                    doc_parser.generate_report(self.full_path(name), job.fields, tables)
                except KeyError as ex:
                    self.display_error(ex)
                
                self.display_message("Genereated report.")
                
                #upload the document back to google drive
                from googleapiclient.http import MediaFileUpload
        
                file_metadata = {'name': new_name}
                media = MediaFileUpload(self.full_path(name),
                                        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                try:
                    file = self.service.files().create(body=file_metadata,
                                                    media_body=media,
                                                    supportsTeamDrives=True,
                                                    fields='id').execute()
                except FileNotFoundError:
                    self.display_error("Could not upload " + new_name)
                    return
                    
                self.display_message("Uploaded report.")
                
                #add the permissions for the user or domain given in ShareWith field
                #When google api allows transfer of ownership that would be a better method
                new_file_id = file.get('id')  
                
                #sometimes permission can't be granted until afew seconds
                #after the document has been created so try this
                #a few times every 2 seconds to allow time for it to work
                attempts = 0
                perm_granted = False
                while perm_granted == False and attempts < 4:
                    try:                                      
                        if '@' in job.fields['ShareWith']:
                            self.service.permissions().create(fileId=new_file_id, body={'role': 'writer', 'type': 'user', 'emailAddress': job.fields['ShareWith']}).execute()
                        else:
                            self.service.permissions().create(fileId=new_file_id, body={'role': 'writer', 'type': 'domain', 'domain': job.fields['ShareWith'], 'allowFileDiscovery': True}).execute()
                            
                        perm_granted = True
                        
                    except errors.HttpError:
                        time.sleep(2)
                        attempts += 1
                        perm_granted = False
                        
                if perm_granted:
                    self.display_message("File {0} is now shared with {1}".format(new_name, job.fields['ShareWith']))
                        

    def display_message(self, message):
        """Function that calls the self.view display_message function if it has one
        displaying the message typically on a new line
        
        Args:
            message (str): the message to display
            
        Returns:
            bool: True if the message was succesfully sent to the view
        """
        if callable(getattr(self.view, "display_message", None)) and message is not None:            
            return self.view.display_message(message) 
        
        return False
            
    def display_status(self, message):
        """Function that calls the self.view display_status function if it has one
        which will display an updating message in the same position
        
        Args:
            message (str): the message to display as a status update
            
        Returns:
            bool: True if the message was succesfully sent to the view
        """
        if callable(getattr(self.view, "display_status", None)) and message is not None:
            return self.view.display_status(message) 
        
        return False
            
    def display_error(self, message):
        """Function that calls the self.view display_error function if it has one
        
        Args:
            message (str): the message to display as the error
            
        Returns:
            bool: True if the message was succesfully sent to the view
        """
        if callable(getattr(self.view, "display_error", None)) and message is not None:
            return self.view.display_error(message) 
        
        return False
            