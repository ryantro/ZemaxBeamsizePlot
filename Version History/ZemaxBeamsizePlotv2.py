# -*- coding: utf-8 -*-
"""
Created on Fri Oct 15 10:47:54 2021

@author: ryan.robinson

Fuction:
    Plot color plot and knife edge plot for a given ZEMAX file

Known issues: 
    fails if a detector is only 1 pixel

TODO:
    1. Edit number of analysis rays
        - How should it hangle multiple sources?
    2. Open most recent file
    3. Add GUI
"""

from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
from win32com.client import gencache
import matplotlib.pyplot as plt
import numpy as np




# Notes
#
# The python project and script was tested with the following tools:
#       Python 3.4.3 for Windows (32-bit) (https://www.python.org/downloads/) - Python interpreter
#       Python for Windows Extensions (32-bit, Python 3.4) (http://sourceforge.net/projects/pywin32/) - for COM support
#       Microsoft Visual Studio Express 2013 for Windows Desktop (https://www.visualstudio.com/en-us/products/visual-studio-express-vs.aspx) - easy-to-use IDE
#       Python Tools for Visual Studio (https://pytools.codeplex.com/) - integration into Visual Studio
#
# Note that Visual Studio and Python Tools make development easier, however this python script should should run without either installed.

class PythonStandaloneApplication(object):
    class LicenseException(Exception):
        pass

    class ConnectionException(Exception):
        pass

    class InitializationException(Exception):
        pass

    class SystemNotPresentException(Exception):
        pass

    def __init__(self):
        # make sure the Python wrappers are available for the COM client and
        # interfaces
        gencache.EnsureModule('{EA433010-2BAC-43C4-857C-7AEAC4A8CCE0}', 0, 1, 0)
        gencache.EnsureModule('{F66684D7-AAFE-4A62-9156-FF7A7853F764}', 0, 1, 0)
        # Note - the above can also be accomplished using 'makepy.py' in the
        # following directory:
        #      {PythonEnv}\Lib\site-packages\wind32com\client\
        # Also note that the generate wrappers do not get refreshed when the
        # COM library changes.
        # To refresh the wrappers, you can manually delete everything in the
        # cache directory:
        #	   {PythonEnv}\Lib\site-packages\win32com\gen_py\*.*
        
        self.TheConnection = EnsureDispatch("ZOSAPI.ZOSAPI_Connection")
        if self.TheConnection is None:
            raise PythonStandaloneApplication.ConnectionException("Unable to intialize COM connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise PythonStandaloneApplication.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None

        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)
    def CloseFile(self, save):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException("Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        if self.TheApplication.LicenseStatus is constants.LicenseStatusType_PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_ProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_StandardEdition:
            return "Standard"
        else:
            return "Invalid"









def main():
    zosapi = PythonStandaloneApplication()
    
    # Insert Code Here
    TheSystem = zosapi.TheSystem
    #TheApplication = zosapi.TheApplication
    #Logic = ['False', 'True']
    
    # Lets add a path to the file I'm using to test
    # testFile = r'C:\Users\lab\Documents\Zemax Module Collimator Optics\Module_Collimator_Rev5.01.zmx'
    testFile = r'C:\Users\lab\Documents\Zemax Module Collimator Optics\Module_Collimator_Rev8.00_PlanoConvexFF.zmx'
    
    # LOAD FILE
    TheSystem.LoadFile(testFile, False)
    print(TheSystem.SystemFile)

    # RUN RAYTRACE
    print('Running ray trace...')
    NSCRayTrace = TheSystem.Tools.OpenNSCRayTrace()
    NSCRayTrace.SplitNSCRays = True
    NSCRayTrace.ScatterNSCRays = False
    NSCRayTrace.UsePolarization = True
    NSCRayTrace.IgnoreErrors = True
    NSCRayTrace.SaveRays = False
    NSCRayTrace.ClearDetectors(0)
    baseTool = CastTo(NSCRayTrace, 'ISystemTool')
    baseTool.RunAndWaitForCompletion()
    baseTool.Close()
    print('Ray trace finished.')

    objNum = TheSystem.NCE.NumberOfObjects
    print('Total Number of Objects in System: {}'.format(objNum))
    
    for i in range(0,objNum):
        TypeName = TheSystem.NCE.GetObjectAt(i+1).TypeName
        if(TypeName == "Detector Rectangle"):
            A,o_x,o_y = parseDetector(TheSystem,i+1)
            
            det = i+1
            BP = BeamProfile(A,o_x,o_y)
            # Plotting
            print('Plotting colorplot for detector object {}'.format(det))
            BP.colorPlot(title="Detector Object {} Color Plot".format(det))
            print('Plotting knife edge for detector object {}'.format(det))
            BP.knifeEdge(title="Detector Object {} Knife Edge Plot".format(det))
            print('')
      
    del zosapi
    zosapi = None
  
    
  
    
    
    
    
  
def parseDetector(TheSystem,det):
    """
    Def:
        - Extract detector array and settings
    
    Input:
        - NCE: NCE object
        - det: Detector number
    
    Output:
        - None
        """
    # Get detector total size in pixels
    dsize = TheSystem.NCE.GetDetectorSize(det)
    
    # Get detector dimensions in pixels
    ddims = TheSystem.NCE.GetDetectorDimensions(det)
    x_pixels = ddims[1]
    y_pixels = ddims[2]
    
    # Get detector dimensions in mm
    a = TheSystem.NCE.GetObjectAt(det).GetObjectCell(constants.ObjectColumn_Par1)
    o_x = float(str(a))*2
    b = TheSystem.NCE.GetObjectAt(det).GetObjectCell(constants.ObjectColumn_Par2)
    o_y = float(str(b))*2
    
    # Read detector values
    
    d = TheSystem.Analyses.New_Analysis(constants.AnalysisIDM_DetectorViewer)
    d_set = d.GetSettings()
    setting = CastTo(d_set, 'IAS_DetectorViewer')
    setting.Detector.SetDetectorNumber(det)
    #setting.ShowAs = constants.DetectorViewerShowAsTypes_FalseColor
    d.ApplyAndWaitForCompletion()
    d_results = d.GetResults()
    results = CastTo(d_results, 'IAR_')
    d_values = results.GetDataGrid(0).Values
    A = np.array(d_values)
    
    # Slow due to number of API commands
    """
    A = np.zeros([x_pixels,y_pixels])
    for i in range(0,dsize):
        a, A[int(i/x_pixels),i%x_pixels] = NCE.GetDetectorData(det,i+1,0)

    """
    
    return A, o_x, o_y




class BeamProfile:
    def __init__(self,array,xsize,ysize):
        """
        Inputs: 
            - NP Array: 2D Intensity Array
            - Float:    X Detector Size (um)
            - Float:    Y Detector Size (um)
        """
        self.A = array
        self.sx = xsize
        self.sy = ysize
        return
    
    def colorPlot(self,title='Color Plot'):
        o_x = self.sx/2
        o_y = self.sy/2
        
        C = self.A
        sumC = np.sum(C) # Normal eyes
        print("Power: {}".format(sumC))
        
        if(sumC == 0):
            sumC = 1.0 # Avoid divide by zero error for a detector with 0 hits
            print("Error: \""+title+"\" has no hits.")
            return
        
        # FIND DETECTOR SIZE IN PIXELS
        py = len(C[:,0])
        px = len(C[0,:])
        
        
        # GENERATE X AXIS
        x_um = np.linspace(1,px,px) * (self.sx/px)
        
        # GENERATE Y AXIS
        y_um = np.linspace(1,py,py) * (self.sy/py)
        
        # X and Y Axis
        X = np.sum(C, axis = 0) / sumC 
        Y = np.sum(C, axis = 1) / sumC
        
        # Find center by taking dot product of 1D normalized power with respective axis
        Cx = np.dot(X,x_um) - o_x
        Cy = np.dot(Y,y_um) - o_y
        
        # All I do is plot plot plot plot
        plt.figure(title)
        plt.title(title)
        plt.imshow(np.flipud(self.A), cmap='jet', extent=[-o_x, o_x, -o_y, o_y])
        plt.xlabel('x (um)')
        plt.ylabel('y (um)')
        plt.text(-o_x*0.8,-o_y*0.8,"X Center: {:.4f}\nY Center: {:.4f}".format(Cx,Cy),color='white')
        
        
        plt.colorbar()
        return
    
    
    
    def knifeEdge(self,title='Beam Width Plot'):
        C = self.A
        # ALONG AXIS 1 FIRST
        sumC = np.sum(C)
        if(sumC == 0):
            sumC = 1.0 # Avoid divide by zero error for a detector with 0 hits
            print("Error: \""+title+"\" has no hits.")
            return
        C = C / sumC
        X = np.sum(C, axis = 0)
        Y = np.sum(C, axis = 1)
        self.X = X
        self.Y = Y
        
        # DETECTOR SIZE IN MICRON
        h = self.sy
        w = self.sx
        
        # FIND DETECTOR SIZE IN PIXELS
        ph = len(C[:,0])
        pw = len(C[0,:])
        
        # FIND PIXEL SIZE
        h_scale = h/ph
        w_scale = w/pw
        
        # GENERATE X AXIS
        x_axis_pixels = np.linspace(1,pw,pw)
        
        # GENERATE Y AXIS
        y_axis_pixels = np.linspace(1,ph,ph)
        
        centered = True
        if(centered):
            xc = self.findCenter(X)
            yc = self.findCenter(Y)
            x_axis_pixels = x_axis_pixels - x_axis_pixels[xc]
            y_axis_pixels = y_axis_pixels - y_axis_pixels[yc]
            
        # GENERATE X AND Y AXIS IN UM    
        x_axis_mm = x_axis_pixels * w_scale
        y_axis_mm = y_axis_pixels * h_scale
        
        # PRINT DETECTOR SIZE IN MICRON AND IN PIXELS
        print("Detector: {}mm X {}mm; {}pixels X {}pixels".format(h, w, ph, pw))
        
        
        # FIND THE WIDTH OF THE BEAM IN X
        R, a, b = self.findWidth(X)
        plt.figure(title)
        plt.plot(x_axis_mm,R)
        plt.title(title)
        plt.xlabel('Knife Position (mm)')
        plt.ylabel('Fractional Power (W/W)')
        deltax = (b - a) * w_scale
        
        # FIND THE WIDTH OF THE BEAM IN Y
        R, a, b = self.findWidth(Y)
        plt.figure(title)
        plt.plot(y_axis_mm,R)
        plt.ylim([0,1])
        deltay = (b - a) * h_scale
        plt.text(min([min(x_axis_mm),min(y_axis_mm)]),0.6,"Space Between 5% and 95%:\n   X : {:.5f} mm\n   Y : {:.5f} mm\n".format(deltax,deltay))
        plt.legend(["Beam Width in X","Beam Width in Y"])
        return
    
    def findWidth(self,X):
        """
        Find 5%-95% Beam Width
        """
        l = len(X)
        R = np.zeros(l)
        pmin = 0.05
        pmax = 0.95
        a = 0
        b = l
        for x in range(1,l,1):
            R[x] = R[x-1] + X[x]
            #print(X[x])
        
        for j in range(0,l,1):
            if(R[j] >= pmin):
                print("5% index = {}".format(j))
                a = j
                break
        
        for k in range(0,l,1):
            if(R[k] >= pmax):
                print("95% index = {}".format(k))
                b = k
                break    
        return R, a, b

    def findCenter(self,X):
        """
        Find 5%-95% Beam Width
        """
        l = len(X)
        R = np.zeros(l)
        pmin = 0.5
        a = 0
        for x in range(1,l,1):
            R[x] = R[x-1] + X[x]
        
        for j in range(0,l,1):
            if(R[j] >= pmin):
                print("50% index = {}".format(j))
                a = j
                break
        
  
        return a




if __name__ == '__main__':
    main()