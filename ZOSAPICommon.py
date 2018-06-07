import os

from win32com.client import CastTo, constants
import matplotlib.pyplot as plt
import numpy as np
import seaborn

class AnalysisNotRunException(Exception):
    def __init__(self):
        print("Analysis has not been perfomed.")

class PathNotFoundException(Exception):
    def __init__(self):
        print("A file path has not been found.")          

class ZOSAPI_MCE(object):
    '''
      Method wrappers for MCE functionality.
    '''
    def __init__(self, TheSystem, TheApplication):
        self.TheSystem = TheSystem
        self.TheApplication = TheApplication
        self.TheMCE = self.TheSystem.MCE

    def set(self, index):
        self.TheMCE.SetCurrentConfiguration(index)

class ZOSAPI_MTF(object):
    '''
      Method wrappers for ZOSAPI MTF functionality.
    '''
    def __init__(self, TheSystem, TheApplication):
        self.TheSystem = TheSystem
        self.TheApplication = TheApplication

        self.TheAnalyses = self.TheSystem.Analyses
        
        self.winh = self.TheAnalyses.New_FftMtf()
        self.winh_Settings = self.winh.GetSettings()
        self.winh_SettingsCast = CastTo(self.winh_Settings,'IAS_FftMtf')

        self.winh_ResultsCast = None

    def set(self, MaximumFrequency=0, ShowDiffractionLimit=True):
        self.winh_SettingsCast.MaximumFrequency = MaximumFrequency
        self.winh_SettingsCast.ShowDiffractionLimit = ShowDiffractionLimit

    def run(self):
        self.winh.ApplyAndWaitForCompletion()
        self.winh_Results = self.winh.GetResults()
        self.winh_ResultsCast = CastTo(self.winh_Results, 'IAR_') 

    def plot(self, title):
        if self.winh_ResultsCast is None:
            raise AnalysisNotRunException

        colors = ('b','g','r','c', 'm', 'y', 'k')
        for s in range(0, self.winh_ResultsCast.NumberOfDataSeries, 1):
            data = self.winh_ResultsCast.GetDataSeries(s)
            basetitle = data.Description.split()[1]
            try:
                basetitle = str(round(float(basetitle), 2))
            except:
                pass
            x = np.array(data.XData.Data)
            y = np.array(data.YData.Data)
            plt.plot(x[:], y[:,0], color=colors[s], 
                     label=basetitle + ' (' + data.SeriesLabels[0][0] + ')')
            plt.plot(x[:],y[:,1],linestyle='--', color=colors[s], 
                      label=basetitle + ' (' + data.SeriesLabels[1][0] + ')')
    
        plt.title(title)
        plt.xlabel('Spatial Frequency (cycles/mm)')
        plt.ylabel('MTF')
        plt.grid(True)
        plt.margins(x=0, y=0)
        plt.legend(loc='upper left')

class ZOSAPI_System(object):
    '''
      Method wrappers for ZOSAPI system functionality.
    '''
    def __init__(self, TheSystem, TheApplication):
        self.TheSystem = TheSystem
        self.TheApplication = TheApplication
    
    def addField(self, x, y, weight):
        self.TheSystem.SystemData.Fields.AddField(x, y, weight)

    def deleteAllFields(self):
        TheFields = self.TheSystem.SystemData.Fields.DeleteAllFields()

    def getFields(self):
        TheFields = self.TheSystem.SystemData.Fields
        fields = []
        for f_idx in range(1, TheFields.NumberOfFields + 1, 1):
            x = TheFields.GetField(f_idx).X
            y = TheFields.GetField(f_idx).Y
            this_entry = {'X': x, 'Y': y}
            fields.append(this_entry)
        return fields

    def getMaxFieldIndex(self):
        fields = self.getFields()
        return np.argmax([((f['X']**2) + (f['Y']**2))**0.5 for f in fields])

    def loadFile(self, fpath):
        if not os.path.exists(fpath):
            raise PathNotFoundException
        self.TheSystem.LoadFile(fpath, False)
