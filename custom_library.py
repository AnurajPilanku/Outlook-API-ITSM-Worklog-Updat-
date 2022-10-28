from robot.api import logger

class custom_library(object):

    ROBOT_LIBRARY_SCOPE = 'TEST CASE'

    def raiseException(self, param1, param2, param3=""):
        '''  sample raiseException function '''
        try:
            print("Do something")
            return "output"
            # Robot takes this output as success with rc = Zero & action_status = True
        except Exception as ex:
            raise Exception("Exeption occured ex - " + str(ex))
            # Robot takes this output as failed with rc = Non Zero & action_status = False
            # This Exception will be Caught in Robot Framework TearDown

    def returnException(self,param1):
        '''   sample returnException function '''
        try:
            print("Do something")
            if "something Failure":
                return (False, "Error_Message")
                # Robot takes this output as failed with rc = Non Zero & action_status = False
            else:
                return (True, "output")
            # Robot takes this output as Success with rc =  Zero & action_status = True
        except Exception as ex:
            return (False,"Exeption occured ex - " + str(ex))
            # Update HiveCenter will sense as action_status = False due to tuple[0] = False
