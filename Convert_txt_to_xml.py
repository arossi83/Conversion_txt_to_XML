import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
from tkinter import Tk
from tkinter.filedialog import askdirectory
import tkinter.font as font
import xlsxwriter
import sys
import os
import shutil
import zipfile
import os.path
from os import path
import csv
import pathlib
from decimal import Decimal
from pathlib import Path
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
import xml.etree.cElementTree as ET
import xml.dom.minidom
import re
from datetime import datetime, timedelta
import openpyxl
from pathlib import Path
from capacitor import capacitorFunc
from FET import FETFunc
from MOS import MOSFunc
from strip import stripFunc
from strip_rot import strip_rotFunc
from Poly import PolyFunc
from Poly_rot import Poly_rotFunc
from pstop import pstopFunc
from pstop_rot import pstop_rotFunc
from DielectricBreakdown import DielectricBreakdownFunc
from GCD import GCDFunc
from LinewidthStrip import LinewidthStripFunc
from LinewidthPolyMeander import LinewidthPolyMeanderFunc
from Linewidthpstop import LinewidthpstopFunc
from BulkCross import BulkCrossFunc
from BulkCross_rot import BulkCross_rotFunc
from DiodeCV import DiodeCVFunc
from DiodeIV import DiodeIVFunc
from MetalClover import MetalCloverFunc
from pBridge import pBridgeFunc
from pCross import pCrossFunc
from MetalMeanderChain import MetalMeanderChainFunc
from GCD05 import GCD05Func
from stripCBKR import stripCBKRFunc
from polyCBKR import polyCBKRFunc
from nChain import nChainFunc
from pChain import pChainFunc
from polyChain import polyChainFunc

rootdir = os.getcwd()

dir_list = next(os.walk('.'))[1]
for folder in dir_list:
    folderName = str(folder)
    if (folderName != '__pycache__'):
        oldPath = Path(folderName)
        newPathString = 'Converted_' + folderName
        
        for file in os.listdir(oldPath):
            fileEx = os.fsdecode(file)

            if fileEx.endswith("_R.xlsx"):
                xlsxName_R = str(fileEx)


            if fileEx.endswith(".xlsx") and "_R" not in fileEx:
                xlsxName = str(fileEx)

        if path.exists(oldPath):
            newPath = Path(newPathString)
            pathlib.Path(newPath).mkdir(parents=True, exist_ok=True) 

            for file in os.listdir(oldPath):
                fileCurr = os.fsdecode(file)
                
                if fileCurr.endswith(".txt"):

                    keywordsrot = ['_Rot']

                    ## LEFT

                    #Capacitor
                    keywordsCapacitor = ['flute1_L_Capacitor']
                    for keyword in keywordsCapacitor:
                        if keyword in fileCurr:
                            fileOutCapacitor = capacitorFunc(oldPath, newPath, xlsxName, fileCurr)

                    #FET
                    keywordsFET = ['flute1_L_FET']
                    for keyword in keywordsFET:
                        if keyword in fileCurr:
                            fileOutFET = FETFunc(oldPath, newPath, xlsxName, fileCurr)
                    #MOS
                    keywordsMOS = ['flute1_L_MOS']
                    for keyword in keywordsMOS:
                        if keyword in fileCurr:
                            fileOutMOS = MOSFunc(oldPath, newPath, xlsxName, fileCurr)

                    #strip
                    keywordsnPlus = ['flute1_L_n+']
                    for keyword in keywordsnPlus:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutstrip_rot = strip_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutstrip = stripFunc(oldPath, newPath, xlsxName, fileCurr)


                    #Poly
                    keywordsPoly = ['flute1_L_Poly']
                    for keyword in keywordsPoly:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutPoly_rot = Poly_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutPoly = PolyFunc(oldPath, newPath, xlsxName, fileCurr)

                    #pstop
                    keywordspstop = ['flute1_L_pstop']
                    for keyword in keywordspstop:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutpstop_rot = pstop_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutpstop = pstopFunc(oldPath, newPath, xlsxName, fileCurr)


                    #Dielectric breakdown
                    keywordsDielectric = ['flute2_L_Dielectric']
                    for keyword in keywordsDielectric:
                        if keyword in fileCurr:
                            fileOutDielectricBreakdown = DielectricBreakdownFunc(oldPath, newPath, xlsxName, fileCurr)

                    #GCD
                    keywordsGCD = ['flute2_L_GCD']
                    for keyword in keywordsGCD:
                        if keyword in fileCurr:
                            fileOutGCD = GCDFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth strip
                    keywordsnPlusLinewidth = ['flute2_L_n+_linewidth']
                    for keyword in keywordsnPlusLinewidth:
                        if keyword in fileCurr:
                            fileOutLinewidthStrip = LinewidthStripFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth Poly Meander
                    keywordsPolyMeander = ['flute2_L_PolyMeander']
                    for keyword in keywordsPolyMeander:
                        if keyword in fileCurr:
                            fileOutLinewidthPolyMeander = LinewidthPolyMeanderFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth p-stop
                    keywordspstopLinewidth = ['flute2_L_pstopLinewidth']
                    for keyword in keywordspstopLinewidth:
                        if keyword in fileCurr:
                            fileOutLinewidthpstop = LinewidthpstopFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Bulk cross
                    keywordsBulkCross = ['flute3_L_BulckCross']
                    for keyword in keywordsBulkCross:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutBulkCross_rot = BulkCross_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutBulkCross = BulkCrossFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Diode CV
                    keywordsDiodeCV = ['flute3_L_DiodeCV']
                    for keyword in keywordsDiodeCV:
                        if keyword in fileCurr:
                            fileOutDiodeCV = DiodeCVFunc(oldPath, newPath, xlsxName, fileCurr)
                   
                    #Diode IV
                    keywordsDiodeIV = ['flute3_L_DiodeIV']
                    for keyword in keywordsDiodeIV:
                        if keyword in fileCurr:
                            fileOutDiodeIV = DiodeIVFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Metal clover
                    keywordsMetalClover = ['flute3_L_MetalCover']
                    for keyword in keywordsMetalClover:
                        if keyword in fileCurr:
                            fileOutMetalClover = MetalCloverFunc(oldPath, newPath, xlsxName, fileCurr)

                    #p-Bridge
                    keywordspPlusBridge = ['flute3_L_p+Bridge']
                    for keyword in keywordspPlusBridge:
                        if keyword in fileCurr:
                            fileOutpBridge = pBridgeFunc(oldPath, newPath, xlsxName, fileCurr)
                   
                    #p-Cross
                    keywordspPlusCross = ['flute3_L_p+Cross']
                    for keyword in keywordspPlusCross:
                        if keyword in fileCurr:
                            fileOutpCross = pCrossFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Metal Meander Chain
                    keywordsMetalMeanderChain = ['L_flute3_Metal_Meander_Chain']
                    for keyword in keywordsMetalMeanderChain:
                        if keyword in fileCurr:
                            fileOutMetalMeanderChain = MetalMeanderChainFunc(oldPath, newPath, xlsxName, fileCurr)
                    
                    #GCD05
                    keywordsGCD05 = ['flute4_L_GCD']
                    for keyword in keywordsGCD05:
                        if keyword in fileCurr:
                            fileOutGCD05 = GCD05Func(oldPath, newPath, xlsxName, fileCurr)

                    #strip CBKR
                    keywordsnPlusCBKR = ['flute4_L_n+CBKR']
                    for keyword in keywordsnPlusCBKR:
                        if keyword in fileCurr:
                            fileOutstripCBKR = stripCBKRFunc(oldPath, newPath, xlsxName, fileCurr)

                    #poly CBKR
                    keywordspolyCBKR = ['flute4_L_polyCBKR']
                    for keyword in keywordspolyCBKR:
                        if keyword in fileCurr:
                            fileOutpolyCBKR = polyCBKRFunc(oldPath, newPath, xlsxName, fileCurr)

                    #n-chain
                    keywordsnPlusChain = ['L_flute4_n+_Chain']
                    for keyword in keywordsnPlusChain:
                        if keyword in fileCurr:
                            fileOutnChain = nChainFunc(oldPath, newPath, xlsxName, fileCurr)

                    #p-chain
                    keywordspPlusChain = ['L_flute4_p+_Chain']
                    for keyword in keywordspPlusChain:
                        if keyword in fileCurr:
                            fileOutpChain = pChainFunc(oldPath, newPath, xlsxName, fileCurr)

                    #poly Chain
                    keywordsPolyChain = ['L_flute4_Poly_Chain']
                    for keyword in keywordsPolyChain:
                        if keyword in fileCurr:
                            fileOutpolyChain = polyChainFunc(oldPath, newPath, xlsxName, fileCurr)



                    ## RIGHT

                    #Capacitor
                    keywordsCapacitor_R = ['flute1_R_Capacitor']
                    for keyword in keywordsCapacitor_R:
                        if keyword in fileCurr:
                            fileOutCapacitor_R = capacitorFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #FET
                    keywordsFET_R = ['flute1_R_FET']
                    for keyword in keywordsFET_R:
                        if keyword in fileCurr:
                            fileOutFET_R = FETFunc(oldPath, newPath, xlsxName_R, fileCurr)
                    #MOS
                    keywordsMOS_R = ['flute1_R_MOS']
                    for keyword in keywordsMOS_R:
                        if keyword in fileCurr:
                            fileOutMOS_R = MOSFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #strip
                    keywordsnPlus_R = ['flute1_R_n+']
                    for keyword in keywordsnPlus_R:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutstrip_R_rot = strip_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutstrip_R = stripFunc(oldPath, newPath, xlsxName_R, fileCurr)


                    #Poly
                    keywordsPoly_R = ['flute1_R_Poly']
                    for keyword in keywordsPoly_R:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutPoly_R_rot = Poly_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutPoly_R = PolyFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #pstop
                    keywordspstop_R = ['flute1_R_pstop']
                    for keyword in keywordspstop_R:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutpstop_R_rot = pstop_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutpstop_R = pstopFunc(oldPath, newPath, xlsxName_R, fileCurr)


                    #Dielectric breakdown
                    keywordsDielectric_R = ['flute2_R_Dielectric']
                    for keyword in keywordsDielectric_R:
                        if keyword in fileCurr:
                            fileOutDielectricBreakdown_R = DielectricBreakdownFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #GCD
                    keywordsGCD_R = ['flute2_R_GCD']
                    for keyword in keywordsGCD_R:
                        if keyword in fileCurr:
                            fileOutGCD_R = GCDFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Linewidth strip
                    keywordsnPlusLinewidth_R = ['flute2_R_n+_linewidth']
                    for keyword in keywordsnPlusLinewidth_R:
                        if keyword in fileCurr:
                            fileOutLinewidthStrip_R = LinewidthStripFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Linewidth Poly Meander
                    keywordsPolyMeander_R = ['flute2_R_PolyMeander']
                    for keyword in keywordsPolyMeander_R:
                        if keyword in fileCurr:
                            fileOutLinewidthPolyMeander_R = LinewidthPolyMeanderFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Linewidth p-stop
                    keywordspstopLinewidth_R = ['flute2_R_pstopLinewidth']
                    for keyword in keywordspstopLinewidth_R:
                        if keyword in fileCurr:
                            fileOutLinewidthpstop_R = LinewidthpstopFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Bulk cross
                    keywordsBulkCross_R = ['flute3_R_BulckCross']
                    for keyword in keywordsBulkCross_R:
                        if keyword in fileCurr:
                            for keyword2 in keywordsrot:
                                if keyword2 in fileCurr:
                                    fileOutBulkCross_R_rot = BulkCross_rotFunc(oldPath, newPath, fileCurr)
                                else:
                                    fileOutBulkCross_R = BulkCrossFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Diode CV
                    keywordsDiodeCV_R = ['flute3_R_DiodeCV']
                    for keyword in keywordsDiodeCV_R:
                        if keyword in fileCurr:
                            fileOutDiodeCV_R = DiodeCVFunc(oldPath, newPath, xlsxName_R, fileCurr)
                   
                    #Diode IV
                    keywordsDiodeIV_R = ['flute3_R_DiodeIV']
                    for keyword in keywordsDiodeIV_R:
                        if keyword in fileCurr:
                            fileOutDiodeIV_R = DiodeIVFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Metal clover
                    keywordsMetalClover_R = ['flute3_R_MetalCover']
                    for keyword in keywordsMetalClover_R:
                        if keyword in fileCurr:
                            fileOutMetalClover_R = MetalCloverFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #p-Bridge
                    keywordspPlusBridge_R = ['flute3_R_p+Bridge']
                    for keyword in keywordspPlusBridge_R:
                        if keyword in fileCurr:
                            fileOutpBridge_R = pBridgeFunc(oldPath, newPath, xlsxName_R, fileCurr)
                   
                    #p-Cross
                    keywordspPlusCross_R = ['flute3_R_p+Cross']
                    for keyword in keywordspPlusCross_R:
                        if keyword in fileCurr:
                            fileOutpCross_R = pCrossFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #Metal Meander Chain
                    keywordsMetalMeanderChain_R = ['R_flute3_Metal_Meander_Chain']
                    for keyword in keywordsMetalMeanderChain_R:
                        if keyword in fileCurr:
                            fileOutMetalMeanderChain_R = MetalMeanderChainFunc(oldPath, newPath, xlsxName_R, fileCurr)
                    
                    #GCD05
                    keywordsGCD05_R = ['flute4_R_GCD']
                    for keyword in keywordsGCD05_R:
                        if keyword in fileCurr:
                            fileOutGCD05_R = GCD05Func(oldPath, newPath, xlsxName_R, fileCurr)

                    #strip CBKR
                    keywordsnPlusCBKR_R = ['flute4_R_n+CBKR']
                    for keyword in keywordsnPlusCBKR_R:
                        if keyword in fileCurr:
                            fileOutstripCBKR_R = stripCBKRFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #poly CBKR
                    keywordspolyCBKR_R = ['flute4_R_polyCBKR']
                    for keyword in keywordspolyCBKR_R:
                        if keyword in fileCurr:
                            fileOutpolyCBKR_R = polyCBKRFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #n-chain
                    keywordsnPlusChain_R = ['R_flute4_n+_Chain']
                    for keyword in keywordsnPlusChain_R:
                        if keyword in fileCurr:
                            fileOutnChain_R = nChainFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #p-chain
                    keywordspPlusChain_R = ['R_flute4_p+_Chain']
                    for keyword in keywordspPlusChain_R:
                        if keyword in fileCurr:
                            fileOutpChain_R = pChainFunc(oldPath, newPath, xlsxName_R, fileCurr)

                    #poly Chain
                    keywordsPolyChain_R = ['R_flute4_Poly_Chain']
                    for keyword in keywordsPolyChain_R:
                        if keyword in fileCurr:
                            fileOutpolyChain_R = polyChainFunc(oldPath, newPath, xlsxName_R, fileCurr)




        zipFileName = newPathString + '.zip'

        zf = zipfile.ZipFile(zipFileName, "w")
        for dirname, subdirs, files in os.walk(newPath):
            zf.write(dirname)
            for filename in files:
                zf.write(os.path.join(dirname, filename))
        zf.close()


# Create the FinalFiles directory
os.makedirs("FinalFiles", exist_ok=True)

# Get all directories in the current folder
directories = [dir_name for dir_name in os.listdir() if os.path.isdir(dir_name)]

# Iterate over the directories and copy the ones starting with "Converted"
for directory in directories:
    if directory.startswith("Converted"):
        shutil.copytree(directory, os.path.join("FinalFiles", directory))

# Compress the FinalFiles directory into a zip file
shutil.make_archive("FinalFiles", "zip", ".", "FinalFiles")

