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

sys.path.insert(0,r'N:\SOFTWARE\Python\Beam Profiler Data Analyzer')
import BeamProfilerDataAnalyzerV4

class Demo1:
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
        defaultFile = r'C:\Users\lab\Documents\Zemax Module Collimator Optics\Module_Collimator_Rev8.00 HeNe Alignment Guide.zmx'
        
        self.entry = tk.Entry(self.frame, text = 'Entry', width = minw*5)
        self.entry.insert(0, defaultFile)
        self.entry.grid(row=r, column=1, columnspan = 5, sticky = 'w')

        # BROWS FOR FILE
        self.browsButton = tk.Button(self.frame, text = 'Browse', command = self.brows, width = minw)
        self.browsButton.grid(row = r, column = 6)
        
        
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
        
        

        
        ############################## ROW 6 #################################
        r = 2
        
        # RUN BUTTON
        self.getButton = tk.Button(self.frame, text = 'Run!', command = self.run, width = minw*4)
        self.getButton.grid(row=r, column=0, columnspan = 7)

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
        print('-------------------------------------------------------------')
        # GET DETECTOR ENTRY BOX DETAILS
        rawEntry = self.dtrs.get()
        tmp = rawEntry.split(',')
        tmp = np.array(tmp, dtype = int)
        
        # SHOW ZEMAX ENTRY BOX DETAILS 
        print("ZEMAX File:")
        print(self.entry.get())
        
        # SHOW DETECTOR CHECKBOX DETAILS
        print("Limit Detectors:")
        print(self.dtrsLim.get())
        if(self.dtrsLim.get()):
            print('Detector Objects Being Plotted:')
            print(tmp)
        
        # SAVE DETECTOR ENTRY BOX DETAILS
        self.dtrsList = tmp
        print('-------------------------------------------------------------')
        ######################################################################

        # ZOS API STUFF
        zosapi = PythonStandaloneApplication()
        TheSystem = zosapi.TheSystem
        
        # ZEMAX FILE TO USE
        zemaxFile = self.entry.get()
        
        # LOAD FILE
        TheSystem.LoadFile(zemaxFile, False)
        print(TheSystem.SystemFile)
    
        # RAYTRACE SETTINGS
        print('Running ray trace...')
        NSCRayTrace = TheSystem.Tools.OpenNSCRayTrace()
        NSCRayTrace.SplitNSCRays = True
        NSCRayTrace.ScatterNSCRays = False
        NSCRayTrace.UsePolarization = True
        NSCRayTrace.IgnoreErrors = True
        NSCRayTrace.SaveRays = False
        NSCRayTrace.ClearDetectors(0)
        
        # FIND TOTAL NUMBER OF OBJECTS IN THE SYSTEM
        objNum = TheSystem.NCE.NumberOfObjects
        print('Total Number of Objects in System: {}'.format(objNum))
        
        
        # RUN RAYTRACE
        baseTool = CastTo(NSCRayTrace, 'ISystemTool')
        baseTool.RunAndWaitForCompletion()
        baseTool.Close()
        print('Ray trace finished.')
        print('-------------------------------------------------------------')
        
        # LOOK THROUGH ALL OBJECTS FOR DETECTORS RECTANGLES
        for i in range(0,objNum):
            TypeName = TheSystem.NCE.GetObjectAt(i+1).TypeName
            if(TypeName == "Detector Rectangle" and (not self.dtrsLim.get() or i+1 in self.dtrsList)):
                
                # GET DETECTOR RECTANGLE SIZ
                A,o_x,o_y = parseDetector(TheSystem,i+1)
                
                if(np.sum(A) > 0):
                    # DETECTOR RETANGLE OBJECT NUMBER
                    det = i+1
                    
                    # CREATE BEAM PROFILER OBJECT
                    BP = BeamProfilerDataAnalyzerV4.BeamProfile(A,o_x*1000,o_y*1000)
                    
                    # CALL BEAM PROFILER PLOT METHODS
                    print('Plotting colorplot for detector object {}'.format(det))
                    BP.colorPlot(title="Detector Object {} Color Plot".format(det))
                    print('Plotting knife edge for detector object {}'.format(det))
                    #BP.knifeEdgePlot(title="Detector Object {} Knife Edge Plot".format(det))
                    print('')
        
        # SHOW PLOTS
        plt.show()
        
        # DELETE ZOS API OBJECTS
        del zosapi
        zosapi = None



        ######################################################################                       
        
        
        
        print('-------------------------------------------------------------')
        
        print('Done!')

        print('-------------------------------------------------------------')
        
        return


def main():
   
    root = tk.Tk()
    app = Demo1(root)
    root.mainloop()

    return


    
if __name__=="__main__":
    main()