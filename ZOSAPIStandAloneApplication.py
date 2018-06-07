from win32com.client.gencache import EnsureDispatch, EnsureModule

class PythonStandaloneApplication(object):
    def __init__(self):
        EnsureModule('{EA433010-2BAC-43C4-857C-7AEAC4A8CCE0}', 0, 1, 0)
        EnsureModule('{F66684D7-AAFE-4A62-9156-FF7A7853F764}', 0, 1, 0)
        
        self.TheConnection = EnsureDispatch("ZOSAPI.ZOSAPI_Connection")

        print(self.TheConnection)

        self.TheApplication = self.TheConnection.CreateNewApplication()

        print(self.TheApplication)

        self.TheSystem = self.TheApplication.PrimarySystem

        print(self.TheSystem)

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None
        self.TheConnection = None