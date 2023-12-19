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
from CheckDuplicates import CheckDuplicates

if len(sys.argv) != 2:
    sys.exit(0)

datadir=sys.argv[1]


rootdir = os.getcwd()

dir_list = next(os.walk(datadir))[1]
for folder in dir_list:
    outPar=[]
    outFile=[]
    folderName = str(folder)
    print("Directory --------> %s" % folderName)
    if (folderName != '__pycache__'):
        oldPath = Path(("%s/%s" % (datadir,folderName)))
        newPathString = ("%s/Converted_%s" % (datadir,folderName))
        side=re.findall("HM_[A-Z]",folderName)[0]
        ff=re.findall("[0-9]+_[0-9]+",folderName)[0]
        batch=ff.split("_")
        afnR=("%s_%s_%s_R.xlsx" % (batch[0],batch[1],side[3]))
        afn=("%s_%s_%s.xlsx" % (batch[0],batch[1],side[3]))
        for file in os.listdir(oldPath):
            fileEx = os.fsdecode(file)
            if fileEx == ("%s_%s_%s_R.xlsx" % (batch[0],batch[1],side[3])):
                print("Analysis File Right %s found: data will be used" % afnR)
                xlsxName_R = str(fileEx)
            if fileEx == ("%s_%s_%s.xlsx" % (batch[0],batch[1],side[3])):
                print("Analysis File Left %s found: data will be used" % afn)
                xlsxName = str(fileEx)

        if path.exists(oldPath):
            newPath = Path(newPathString)
            pathlib.Path(newPath).mkdir(parents=True, exist_ok=True) 

            for file in os.listdir(oldPath):
                rot=False
                left=False
                fileCurr = os.fsdecode(file)
                
                if fileCurr.endswith(".txt"):

                    rotkw = '_Rot'
                    keywordsDict={
                        'flute1': ['Capacitor','FET','MOS','n+','Poly','pstop'],
                        'flute2': ['Dielectric','GCD','n+_linewidth','PolyMeander','pstopLinewidth'],
                        'flute3': ['BulckCross','DiodeCV','DiodeIV','MetalCover','p+Bridge','p+Cross','Metal_Meander_Chain','p+_Cross'],
                        'flute4': ['GCD','n+CBKR','polyCBKR','n+_Chain','p+_Chain','Poly_Chain']}
                    fluteList=list(keywordsDict.keys())

                    for flute in fluteList:
                        if flute in fileCurr:
                            for kw in keywordsDict[flute]:
                                if kw in fileCurr:
                                    if rotkw in fileCurr:
                                        rot=True
                                    if '_L_' in fileCurr:
                                        left=True

                                    if rot:
                                        if left:
                                            outPar.append(("%s - %s - Left - Rot" % (flute,kw)))
                                            outFile.append(fileCurr)
                                            #print("%s - %s - Left - Rot: %s" % (flute,kw,fileCurr))
                                        else:
                                            outPar.append(("%s - %s - Right - Rot" % (flute,kw)))
                                            outFile.append(fileCurr)
                                            #print("%s - %s - Right - Rot: %s" % (flute,kw,fileCurr))
                                    else:
                                        if left:
                                            outPar.append(("%s - %s - Left - NoRot" % (flute,kw)))
                                            outFile.append(fileCurr)                                            
                                            #print("%s - %s - Left - NoRot: %s" % (flute,kw,fileCurr))
                                        else:
                                            outPar.append(("%s - %s - Right - NoRot" % (flute,kw)))
                                            outFile.append(fileCurr)                                            
                                            #print("%s - %s - Right - NoRot: %s" % (flute,kw,fileCurr))
            
            filteredPar, filteredFile=CheckDuplicates(outPar,outFile)
            #print(outPar)
            #print(filteredPar)
            #print(outFile)
            #print(filteredFile)
            ###Process Each Files
            nL=0
            nR=0
            nL_r=0
            nR_r=0
            for idx,par in enumerate(filteredPar):
                #Flute1 Capacitor
                if 'Capacitor' in par and 'flute1' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutCapacitor = capacitorFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        if "incomplete" not in fileOutCapacitor.name:
                            nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutCapacitor_R = capacitorFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute1 FET
                if 'FET' in par and 'flute1' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutFET = FETFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutFET_R = FETFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute1 MOS
                if 'MOS' in par and 'flute1' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMOS = MOSFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMOS_R = MOSFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute1 n+ 
                if 'n+' in par and 'flute1' in par:
                    if 'NoRot' in par:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutstrip = stripFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                            nL+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutstrip_R = stripFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                            nR+=1
                    else:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutstrip_rot = strip_rotFunc(oldPath, newPath, filteredFile[idx])
                            nL_r+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutstrip_R_rot = strip_rotFunc(oldPath, newPath, filteredFile[idx])
                            nR_r+=1
                #Flute1 Poly 
                if 'Poly' in par and 'flute1' in par:
                    if 'NoRot' in par:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutPoly = PolyFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                            nL+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutPoly_R = PolyFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                            nR+=1
                    else:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutPoly_rot = Poly_rotFunc(oldPath, newPath, filteredFile[idx])
                            nL_r+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutPoly_R_rot = Poly_rotFunc(oldPath, newPath, filteredFile[idx])
                            nR_r+=1
                #Flute1 pstop 
                if 'pstop' in par and 'flute1' in par:
                    if 'NoRot' in par:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutpstop = pstopFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                            nL+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutpstop_R = pstopFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                            nR+=1
                    else:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutpstop_rot = pstop_rotFunc(oldPath, newPath, filteredFile[idx])
                            nL_r+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutpstop_R_rot = pstop_rotFunc(oldPath, newPath, filteredFile[idx])
                            nR_r+=1
                #Flute2 DielectricBreakdown
                if 'Dielectric' in par and 'flute2' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDielectricBreakdown = DielectricBreakdownFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDielectricBreakdown_R = DielectricBreakdownFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute2 GCD
                if 'GCD' in par and 'flute2' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutGCD = GCDFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutGCD_R = GCDFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute2 LinewidthStrip
                if 'n+_linewidth' in par and 'flute2' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthStrip = LinewidthStripFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthStrip_R = LinewidthStripFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute2 LinewidthPolyMeander
                if 'PolyMeander' in par and 'flute2' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthPolyMeander = LinewidthPolyMeanderFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthPolyMeander_R = LinewidthPolyMeanderFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute2 Linewidthpstop
                if 'pstopLinewidth' in par and 'flute2' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthpstop = LinewidthpstopFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutLinewidthpstop_R = LinewidthpstopFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 BulckCross
                if 'BulckCross' in par and 'flute3' in par:
                    if 'NoRot' in par:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutBulckCross = BulkCrossFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                            nL+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutBulckCross_R = BulkCrossFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                            nR+=1
                    else:
                        if 'Left' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutBulckCross_rot = BulkCross_rotFunc(oldPath, newPath, filteredFile[idx])
                            nL_r+=1
                        if 'Right' in par:
#                            print("%s --> %s" % (par,filteredFile[idx]))
                            fileOutBulckCross_R_rot = BulkCross_rotFunc(oldPath, newPath, filteredFile[idx])
                            nR_r+=1
                #Flute3 DiodeCV
                if 'DiodeCV' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDiodeCV = DiodeCVFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDiodeCV_R = DiodeCVFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 DiodeIV
                if 'DiodeIV' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDiodeIV = DiodeIVFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutDiodeIV_R = DiodeIVFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 MetalCover
                if 'MetalCover' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMetalCover = MetalCloverFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMetalCover_R = MetalCloverFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 p+Bridge
                if 'p+Bridge' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpBridge = pBridgeFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpBridge_R = pBridgeFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 p+Cross
                if 'p+Cross' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpCross = pCrossFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpCross_R = pCrossFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute3 MetalMeanderChain
                if 'Metal_Meander_Chain' in par and 'flute3' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMetalMeanderChain = MetalMeanderChainFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutMetalMeanderChain_R = MetalMeanderChainFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute4 GCD05
                if 'GCD' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutGCD05 = GCD05Func(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutGCD05_R = GCD05Func(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute4 n+CBKR
                if 'n+CBKR' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutstripCBKR = stripCBKRFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutstripCBKR_R = stripCBKRFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute4 polyCBKR
                if 'polyCBKR' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpolyCBKR = polyCBKRFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpolyCBKR_R = polyCBKRFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute4 n+_Chain
                if 'n+_Chain' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutnChain = nChainFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutnChain_R = nChainFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
                #Flute4 p+_Chain
                if 'p+_Chain' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpChain = pChainFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpChain_R = pChainFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
		#Flute4 Poly_Chain
                if 'Poly_Chain' in par and 'flute4' in par:
                    if 'Left' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpolyChain = polyChainFunc(oldPath, newPath, xlsxName, filteredFile[idx])
                        nL+=1
                    if 'Right' in par:
#                        print("%s --> %s" % (par,filteredFile[idx]))
                        fileOutpolyChain_R = polyChainFunc(oldPath, newPath, xlsxName_R, filteredFile[idx])
                        nR+=1
            print("%d Left Measurement processed, plus %d rotated" % (nL,nL_r))
            print("%d Right Measurement processed, plus %d rotated" % (nR,nR_r))

        zipFileName = newPathString + '.zip'

        zf = zipfile.ZipFile(zipFileName, "w")
        for dirname, subdirs, files in os.walk(newPath):
            zf.write(dirname)
            for filename in files:
                zf.write(os.path.join(dirname, filename))
        zf.close()


# Create the FinalFiles directory
finalPath=("%s/FinalFiles" % datadir)
os.makedirs(finalPath, exist_ok=True)

# Get all directories in the current folder
directories = [dir_name for dir_name in os.listdir(datadir) if os.path.isdir(os.path.join(datadir,dir_name))]
# Iterate over the directories and copy the ones starting with "Converted"
for directory in directories:
    if directory.startswith("Converted"):
        shutil.copytree(os.path.join(datadir,directory), os.path.join(finalPath, directory))

# Compress the FinalFiles directory into a zip file
shutil.make_archive(finalPath, "zip", datadir, "FinalFiles")
