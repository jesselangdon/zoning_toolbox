"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\Export_DDP_to_PDF_v1.py
... revised by Jesse Q3-2023, ... "...\Export Map Series to PDF"
%ws% ==>  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\...

Purpose:  Within the .MXD that produces the official Zoning Maps, customized Python Script Tools.
              The script that exports each map of the Data Driven Pages to both the archive
              and current folders storing .PDF images of the official Zoning Maps.

Usage:  Python Tool Scripts with ArcMap >> DDP_OZmap Toolbar >> "Export DDP to PDF"
             within W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\TR_DDP_v1.mxd
             1) Run the "Twp Rng DDP Refresh" script tool first, to dial-in the preferred Twp & Rng;
             2) Run the "Export DDP to PDF" script tool second, to generate saved .PDF images.

History:   20120329 - Final Debugging and moved into production
            scdwpr, 20120330
"""

import arcpy
import os, os.path, datetime, glob

#--Variables
gpat_continue_cond = arcpy.GetParameterAsText(0)
# gpat_continue = "true" #TEST
# gpat_continue_cond = str(gpat_continue)
aprx = arcpy.mp.ArcGISProject("CURRENT")
# aprx = arcpy.mp.ArcGISProject(r'C:\Users\SCDJ2L\dev\Zoning\OZmapMaker_v0\OZmapMaker_v20230829.aprx') #TEST
map = aprx.listMaps("Premier Data Frame")[0]
lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]
layout = aprx.listLayouts()[1]
if not layout.mapSeries is None:
    ms = layout.mapSeries
else:
    arcpy.AddError("Map series not available in this Pro project! Map series must be created manually in Pro first...")

currentdir = r'\\snoco\gis\plng\images\zoning_current\pdf'
# currentdir = r"C:\Users\SCDJ2L\dev\Zoning\TEST\zoning_current\pdf" #TEST
archivedir = r'\\snoco\Dept_Images\PDS\PDSHistoricalMapsandAerials\HistoricalZoningMaps'
# archivedir = r"C:\Users\SCDJ2L\dev\Zoning\TEST\HistoricalZoningMaps\pdf" #TEST

def get_date_stamp():
    '''Generate date stamp suffix'''
    now = datetime.datetime.now()
    date_stamp = now.strftime("_%Y%m%d")
    arcpy.AddMessage("date_stamp: " + str(date_stamp))
    return date_stamp


def find_all(a_str, sub):
    '''Function for finding a listing of where the '=' occurs'''
    start = 0
    while True:
        start = a_str.find(sub, start)
        if start == -1:
            return
        yield start
        start += len(sub)


def ToDelCurPDF(CurPDF):
    '''Deletes a file with a wildcard (i.e. the existing current .PDF, hence glob)'''
    for fl in glob.glob(CurPDF):
        os.remove(fl)
    return


def ToRemoveExisting(ExistingPDF):
    if os.path.exists(ExistingPDF):
        arcpy.AddMessage("Removing existing (i.e. duplicate) .PDF file: " + str(ExistingPDF))
        os.remove(ExistingPDF)
    return


def isrunning(exe):
    '''Uses Taskkill to perform a hard close on running instance of Adobe Acrobat Reader application'''
    try :
        p = os.popen(r'tasklist /FI "IMAGENAME eq "'+ exe + ' /FO "LIST" 2>&1' , 'r' )
        PID = p.read().split('\n')[2].split(":")[1].lstrip(" ")
        p.close()
        return PID
    except :
        p.close()
        return "None"


def ms_loop():
    '''Loop through map series pages'''
    #--Global variables
    global currentdir, archivedir, currentpath, archivepath

    #--While loop begins
    pg = 1
    while pg <= ms.pageCount:
        #--Change current page name
        ms.currentPageNumber = pg

        #--Get QSTR-name of current page from .MasterPageGrid
        mpg_label = ms.pageRow.LABEL2

        #--Split LABEL per empty spaces
        mpg_parseList = mpg_label.split()

        #--Sectional
        if len(str(mpg_parseList[1])) == 1:     #--Sec. Str - If... 1-digit Section, then add "0"
            a_sec = "0" + str(mpg_parseList[1])
        elif len(str(mpg_parseList[1])) == 2:
            a_sec = str(mpg_parseList[1])       

        a_twp = str(mpg_parseList[2][1:3])      #--Twp. Str

        if len(str(mpg_parseList[3])) > 3:      #--Rng. Str - If... 2-digit Range, then slice by [1:3]
            a_rng = str(mpg_parseList[3][1:3])
        elif len(str(mpg_parseList[3])) == 3:
            a_rng = "0" + str(mpg_parseList[3][1:2])

        if len(str(mpg_parseList[0])) == 3:     #--Qtr. Str -
            a_qtr = str(0)
        elif len(str(mpg_parseList[0])) > 3:
            a_qtr = str(5)
        elif str(mpg_parseList[0]) == "NE":
            a_qtr = str(1)
        elif str(mpg_parseList[0]) == "NW":
            a_qtr = str(2)        
        elif str(mpg_parseList[0]) == "SW":
            a_qtr = str(3)
        elif str(mpg_parseList[0]) == "SE":
            a_qtr = str(4)

        arcpy.AddMessage("Section = {0}, Township = {1}, Range = {2}, Quarter Section = {3}".format(a_sec, a_twp, a_rng, a_qtr))
        fname = a_twp + a_rng + a_sec + a_qtr + str(date_stamp) + ".pdf"
        fname_cur = a_twp + a_rng + a_sec + a_qtr + "_*.pdf"
        arcpy.AddMessage("Archived filename: {}".format(fname))
        arcpy.AddMessage("Current filename: {}".format(fname_cur))
        
        #--Taskkill to perform a hard close on Adobe Acrobat Reader
        PID = isrunning('AcroRd32.exe')
        if PID != "None" :
            os.system(r'taskkill /F /PID ' + PID)
            
        #--Taskkill to perform a hard close on Adobe Acrobat Pro
        PID = isrunning('Acrobat.exe')
        if PID != "None" :
            os.system(r'taskkill /F /PID ' + PID)

        #--Remove existing .PDF with previous date from current folder (with wildcard, hence glob)
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname_cur)

        ToDelCurPDF(currentpath)

        #--Remove existing .PDF with same names from current and archive folders
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname)
        archivepath = os.path.join(archivedir, a_twp + a_rng, fname)
        arcpy.AddMessage("currentpath: " + currentpath)
        arcpy.AddMessage("archivepath: " + archivepath)

        arcpy.AddMessage("Removing existing PDF files from current and archive folders...")
        ToRemoveExisting(currentpath)
        ToRemoveExisting(archivepath)

        #--Export .PDF to archive folder
        arcpy.AddMessage("Exporting PDF to current folder...")
        layout.exportToPDF(currentpath)

        #--Copy .PDF file from current to archive folder
        arcpy.AddMessage("Copying to archive folder...")
        os.system("copy %s %s" % (currentpath, archivepath))

        #--While loop increment page number
        pg = pg + 1

    return


#--Read the .MasterPageGrid definition query to label archive and current folder paths
if gpat_continue_cond == "true":
    date_stamp = get_date_stamp()
    if lyr.supports("DEFINITIONQUERY"):
        # --Definition query of .PageGrid_MasterPageGrid
        dq = lyr.definitionQuery
        arcpy.AddMessage("definition query: " + str(dq))

        if dq != "":
            #--Make list of all map series pages
            dq_list = list(find_all(dq, '='))
            arcpy.AddMessage("dq_list: " + str(dq_list))

            #--(WHILE) Loop through DDP pages
            arcpy.AddMessage("Initiating WHILE loop through map series pages")
            ms_loop()

        else:    #--Def. Qry. count
            arcpy.AddError("No definition query found for .MasterPageGrid!")

    else:    #--if lyr.supports
        arcpy.AddError("The feature class does not support definition queries!")
else:     #--if Continue checkbox not checked
    arcpy.AddError("Unchecked parameter box ... Bailing out of script tool!")
