# -*- coding: utf-8 -*-
"""
Created on Wed Oct 20 09:37:33 2021

@author: ryan.robinson

tkinter layout:
    https://stackoverflow.com/questions/17466561/best-way-to-structure-a-tkinter-application
    
Purpose:
    The purpose of this code is to be used as a tutorial for learning how to
    create a GUI with tkinter.
    
    You don't need a class structure to do it, but class structure helps with
    organization.
"""

import tkinter as tk

from tkinter import filedialog as fd
import numpy as np
import sys
from ZemaxBeamsizePlot import *
import threading
import time

# TODO Change to folder
# sys.path.insert(0,r'N:\SOFTWARE\Python\Beam Profiler Data Analyzer')
# import BeamProfilerDataAnalyzerV4

sys.path.insert(0,r'C:\Users\ryan.robinson\Documents\Git Sandbox\Beam-Image-Analyzer\bia')
import DataParser as DP

FARFIELD = True

class Application:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)

        minw = 25
        # GRID CONFIGURE
        self.frame.rowconfigure([0, 1, 2], minsize=30, weight=1)
        self.frame.columnconfigure([0, 1, 2, 3], minsize=minw, weight=1)

        
        # BOX CONFIGURE
        self.master.title('ZEMAX Detector Analyzer')
        w = 100
        
        ############################## ROW 0 #################################
        r = 0
        
        # FILENAME LABEL
        self.entryLabel = tk.Label(self.frame,text="Filename:")
        self.entryLabel.grid(row=r, column=0, sticky = "w")

        
        # FILENAME ENTRY BOX
        defaultFile = r'C:/Users/ryan.robinson/Documents/Git Sandbox/Zemax-Beam-Plot/ZemaxBeamsizePlot/tests/Singe_Emitter_Sensitivity_Analysis.zmx'
        
        self.entry = tk.Entry(self.frame, text = 'Entry', width = minw*5)
        self.entry.insert(0, defaultFile)
        self.entry.grid(row=r, column=1, columnspan = 2, sticky = 'w')

        # BROWS FOR FILE
        self.browsButton = tk.Button(self.frame, text = 'Browse', command = self.brows, width = minw)
        self.browsButton.grid(row = r, column = 3)
        
        
        ############################## ROW 1 #################################
        r = 1
        
        # ROW LABEL
        entryLabel = tk.Label(self.frame,text="Only Plot Specific Detectors?")
        entryLabel.grid(row=r, column=0, sticky = "w")
        
        # CHECK BOX VARIABLE
        self.dtrsLim = tk.BooleanVar()
        
        # NEAR FIELD CHECK BOX
        box1 = tk.Checkbutton(self.frame, text = '', variable = self.dtrsLim, onvalue = True, offvalue = False)
        box1.grid(row=r, column = 1, sticky = "w")
        
        # MAGNIFICATION
        label = tk.Label(self.frame, text = "Dector Objects Numbers (seperate with comma):")
        label.grid(row=r,column = 2, sticky = "e")
        self.dtrs = tk.Entry(self.frame, text = 'Detectors', width = minw)
        self.dtrs.insert(0,'10, 6, 7,21')
        self.dtrs.grid(row = r, column = 3, sticky = "w")
        
        
        ########################### RAYTRACE SETTINGINGS #####################
        
        # TRACE CONFIG FRAME
        self.traceFrame = tk.Frame(self.frame)
        self.traceFrame.rowconfigure([0, 1], minsize=30, weight=1)
        self.traceFrame.grid(row = 2, column = 0, columnspan = 2, sticky = "EW", padx = 20, pady = 10)
        traceFrameLabel = tk.Label(self.traceFrame, text = "Raytrace Settings:")
        traceFrameLabel.grid(row = 0, sticky = "W")
        
        # TRACE CONFIG BOX
        self.traceConf = tk.Frame(self.traceFrame, borderwidth = 2, relief="groove")
        self.traceConf.rowconfigure([0, 1], minsize=30, weight=1)
        self.traceConf.columnconfigure([0, 1], minsize=minw, weight=1)
        self.traceConf.grid(row = 1, sticky = "EW")
        
        # VARS
        self.splitVar = tk.BooleanVar()
        self.scatterVar = tk.BooleanVar()
        self.polVar = tk.BooleanVar()
        self.ignoreErrVar = tk.BooleanVar()
        self.ignoreErrVar.set(True)
        
        # CHECK BOXES
        splitBox = tk.Checkbutton(self.traceConf, text = "Split Rays", variable = self.splitVar, onvalue = True, offvalue = False)
        splitBox.grid(row = 0, column = 0, sticky = "W")
        
        splitBox = tk.Checkbutton(self.traceConf, text = "Scatter Rays", variable = self.scatterVar, onvalue = True, offvalue = False)
        splitBox.grid(row = 0, column = 1, sticky = "W")
        
        splitBox = tk.Checkbutton(self.traceConf, text = "Use Polarization", variable = self.polVar, onvalue = True, offvalue = False)
        splitBox.grid(row = 1, column = 0, sticky = "W")
        
        splitBox = tk.Checkbutton(self.traceConf, text = "Ignore Errors", variable = self.ignoreErrVar, onvalue = True, offvalue = False)
        splitBox.grid(row = 1, column = 1, sticky = "W")  

        
        ############################## ROW 6 #################################
        r = 4
        
        # RUN BUTTON
        self.getButton = tk.Button(self.frame, text = 'Run!', command = self.run, width = minw*4)
        self.getButton.grid(row=r, column=0, columnspan = 4, sticky = "EW", padx = 10, pady = 10)

        # FINALLY
        self.frame.pack()
        return
    
    def brows(self):
        filename = fd.askopenfile()
        filename = filename.name
        self.entry.delete(0,'end')
        self.entry.insert(0,filename)  
        return
    
    
    def run(self):
        """
        RUN MEASUREMENT
        """
        self.getButton.configure(state = "disabled")
        self.runThread = threading.Thread(target = self.run3)
        self.runThread.start()
    
    def run2(self):
        """
        RUN MEASUREMENT USING ZEMAX CLASS
        """
        beamWidths = []
        zemaxFile = self.entry.get()
        Z = ZemaxMeasurement(zemaxFile)
        
        zCenter = 1.42412288067400000E-01
        zDelta = 0.005
        points = 21
        zs = np.linspace(zCenter-zDelta, zCenter+zDelta, num = points)
        
        print(zs)
        
        for z in zs:
            print("Z: {} mm".format(z))
            Z.zPos(z)
            Z.rayTrace()
            Z.parseDetector()
            beamWidths.append(Z.Y_w)
        
        # Z.plotDetector()
        Z.close()
        self.getButton.configure(state = "normal")
        
        
        print(beamWidths)
        
        plt.plot(zs,beamWidths)
        plt.xlabel("FAC Position (mm)")
        plt.ylabel("Full Angle Beam Divergence (mrad)")
        plt.title("FAC Z Sensitivity")
        plt.grid()
        plt.show()
        
        return

    def run3(self):
        """
        RUN MEASUREMENT USING ZEMAX CLASS
        """
        beamWidths = []
        zemaxFile = self.entry.get()
        Z = ZemaxMeasurement(zemaxFile)
        
        zCenter = 0
        zDelta = 0.005
        points = 21
        zs = np.linspace(zCenter-zDelta, zCenter+zDelta, num = points)
        
        print(zs)
        
        for z in zs:
            print("Y: {} mm".format(z))
            Z.yPos(z)
            Z.rayTrace()
            Z.parseDetector()
            beamWidths.append(Z.Y)
        
        # Z.plotDetector()
        Z.close()
        self.getButton.configure(state = "normal")
        
        
        print(beamWidths)
        zs = zs * 1000
        plt.plot(zs,beamWidths)
        plt.xlabel("FAC Y Position (um)")
        plt.ylabel("FA Beam Pointing (mrad)")
        plt.title("FAC Y Sensitivity")
        plt.grid()
        plt.show()
        
        return

class ZemaxMeasurement:
    def __init__(self, zemaxFile):
        """
        SETUP THE ZOS API
        """
        # ZOS API STUFF
        print("Connecting to ZOS API...")
        self.zosapi = PythonStandaloneApplication()
        self.TheSystem = self.zosapi.TheSystem

        # LOAD FILE
        print("Loading ZEMAX file...")
        self.TheSystem.LoadFile(zemaxFile, False)
        print(self.TheSystem.SystemFile)    
    
        # RAYTRACE SETTINGS
        print('Running ray trace...')
        self.NSCRayTrace = self.TheSystem.Tools.OpenNSCRayTrace()
        self.NSCRayTrace.SplitNSCRays = False #self.splitVar.get()
        self.NSCRayTrace.ScatterNSCRays = False #self.scatterVar.get()
        self.NSCRayTrace.UsePolarization = False #self.polVar.get()
        self.NSCRayTrace.IgnoreErrors = True #self.ignoreErrVar.get()
        self.NSCRayTrace.SaveRays = False
        self.NSCRayTrace.ClearDetectors(0)
        
        # FIND TOTAL NUMBER OF OBJECTS IN THE SYSTEM
        self.objNum = self.TheSystem.NCE.NumberOfObjects
        print('Total Number of Objects in System: {}'.format(self.objNum))
       
        self.BP = None
        return
    def yPos(self, y, obj = 6):
        """
        CHANGE THE Y POSITION OF OBJECT
        """
        TheNCE = self.TheSystem.NCE
        FAC = TheNCE.GetObjectAt(obj)
        FAC.YPosition = y
        return
    
    def zPos(self, z, obj = 6):
        """
        CHANGE THE Z POSITION OF OBJECT
        """
        TheNCE = self.TheSystem.NCE
        FAC = TheNCE.GetObjectAt(obj)
        FAC.ZPosition = z
        return
       
    def rayTrace(self):
        """
        RUN A RAYTRACE
        """
        # RAYTRACE SETTINGS
        print('Running ray trace...')
  
        self.NSCRayTrace.ClearDetectors(0)
        
        # RUN RAYTRACE
        self.baseTool = CastTo(self.NSCRayTrace, 'ISystemTool')
        self.baseTool.RunAndWaitForCompletion()
        # self.baseTool.Close()
        time.sleep(1)
        print('Ray trace finished.')
        print('-------------------------------------------------------------')
        return

    def parseDetector(self, detector = 11, lens = 1000000):
        """
        GET INFO FROM DETECTOR
        """
        
        TypeName = self.TheSystem.NCE.GetObjectAt(detector).TypeName
        if(TypeName == "Detector Rectangle"):
        
            # GET DETECTOR RECTANGLE SIZE                
            A,o_x,o_y = parseDetector(self.TheSystem, detector)
            
            if(np.sum(A) > 0):
                
                # CREATE BEAM PROFILER OBJECT
                self.BP = DP.BeamProfileFarField(A, o_x*1000, o_y*1000, lens = 1000000)
                
                # BEAM PARAMETERS
                self.X = self.BP.X.center_point/lens
                self.Y = self.BP.Y.center_point/lens
                self.X_w = self.BP.X.width/lens
                self.Y_w = self.BP.Y.width/lens
                print("X Center: {}, Y Center: {}".format(self.X,self.Y))
                print("X Width: {}, Y Width: {}".format(self.X_w, self.Y_w))
                
            else:
                print("ERROR: No light on detector")
        else:
            print("ERROR: not a detector rectangle object")
    
        return
    
    def plotDetector(self):
        """
        PLOT DETECTORS
        """
        detector = 11
        
        if(self.BP != None):
            print('Plotting colorplot for detector object {}'.format(detector))
            self.BP.colorPlot(title="Detector Object {} Color Plot".format(detector))
            plt.show()
            print('')
        else:
            print('Beam profile does not exist.')
        return

    def close(self):
        """
        DELETE ZOS API
        """
        # DELETE ZOS API OBJECTS
        self.baseTool.Close()
        del self.zosapi
        self.zosapi = None
        return

def main():
   
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

    return


    
if __name__=="__main__":
    main()