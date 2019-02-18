import datetime
import os

class SRGConsoleView:
    """ Basic View for displaying messages on the console """
    
    def __init__(self):
        pass
    
    
    def display_message(self, message):   
        try:
            f=open(os.path.join(os.path.dirname(__file__), "activity.log"), "a+")
            f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f") + ": " +  message + "\n")
            f.close()
        except PermissionError:
            self.display_error("Can't access activity.log")
            return False
        return True
        
        
    def display_status(self, status):
        try:
            f=open(os.path.join(os.path.dirname(__file__), "status.txt"), "w+")
            f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f") + ": " + status + "\n")
            f.close()
        except PermissionError:
            self.display_error("Can't access status.txt")
            return False
        return True

    def display_error(self, message):
        try:
            f=open(os.path.join(os.path.dirname(__file__), "errors.log"), "a+")
            f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f") + ": " +  message + "\n")
            f.close()
        except PermissionError:
            print("Can't access error.log")
            return False
        return True


        
        