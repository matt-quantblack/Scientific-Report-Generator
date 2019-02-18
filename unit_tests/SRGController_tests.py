import unittest
from SRGController import SRGController
from SRGConsoleView import SRGConsoleView

class ControllerTestCase(unittest.TestCase):
    
    def setUp(self):
        """ Run before each use case """
        self.c = SRGController(SRGConsoleView())

    def test_create_service_missing_cred_file(self):
        self.assertFalse(self.c.create_service("missing.json"))
        
    def test_create_service_invalid_cred_file(self):
        self.assertFalse(self.c.create_service("invalid.json"))
        
    def test_create_service_invalid_cred_file_correct_json(self):
        self.assertFalse(self.c.create_service("incorrect_creds.json"))
    
    def test_create_service_invalid_service_email(self):
        self.assertFalse(self.c.create_service("unauth_creds.json"))
        
    def test_create_service_success(self):
        self.assertTrue(self.c.create_service("test_env_creds.json"))
    
    def test_clear_job_files_failure(self):       
        #service not created and permission id not defined
        self.assertFalse(self.c.clear_job_files())        
        
    def test_clear_job_files_success(self):        
        self.c.create_service("test_env_creds.json")
        self.assertTrue(self.c.clear_job_files())



def suite():
    suite = unittest.TestSuite()  
    suite.addTest(ControllerTestCase('test_create_service_invalid_cred_file'))
    suite.addTest(ControllerTestCase('test_create_service_missing_cred_file'))
    suite.addTest(ControllerTestCase('test_create_service_invalid_cred_file_correct_json'))
    suite.addTest(ControllerTestCase('test_create_service_invalid_service_email'))
    suite.addTest(ControllerTestCase('test_create_service_success'))
    suite.addTest(ControllerTestCase('test_clear_job_files_failure'))
    suite.addTest(ControllerTestCase('test_clear_job_files_success'))
    return suite

if __name__ == '__main__':
    #unittest.main()
    unittest.TextTestRunner(verbosity=2).run(suite())