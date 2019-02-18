import sys, os, glob


from SRGController import SRGController
from SRGConsoleView import SRGConsoleView

def main():
    """ Main Entry point to program
    Should be run in a new thread. 
    For linux use 'SRG.py start &'
    """
    
    #full path for any session.loc files
    #session.lock files are what keep the background threads running
    #removing a session.lock file with terminate the background thread loop
    #session.lock files are in the format session-{sessionId}.lock
    session_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "session-*")
        
    if 'start' in sys.argv:
        #Can only have one session running at a time
        #remove any other session.lock files present to close
        #the old threads            
        for filename in glob.glob(session_path):
            print("Shutting down open SRG background process ...")
            #delete the session.lock file which will stop the loop   
            os.remove(filename) 

        #create the view and controller
        view = SRGConsoleView()
        controller = SRGController(view)
        
        #Start the main loop       
        controller.start()
        
    elif 'stop' in sys.argv:
        
        #delete the session.lock file which will stop the loop  
        sessions = glob.glob(session_path)
        
        #delete all session files 
        if len(sessions) > 0:
            print("Shutting down SRG background process ...")
            for filename in sessions:
                os.remove(filename) 
        else:
            print("No SRG sessions running.")
        
    
    else:
        print("GUI not currently supported.\nRun 'SRG.py start' to start the background process or 'SRG.py stop' to stop the process.")
    

if __name__ == '__main__':
    main()