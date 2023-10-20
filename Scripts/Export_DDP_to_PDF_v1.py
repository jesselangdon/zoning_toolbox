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

# Set global variables

# gpat_layout = arcpy.GetParameterAsText(0)
gpat_layout = "Full Township"
# gpat_continue_cond = arcpy.GetParameterAsText(1)
gpat_continue = "true" #TEST
gpat_continue_cond = str(gpat_continue)
# aprx = arcpy.mp.ArcGISProject("CURRENT")
aprx = arcpy.mp.ArcGISProject(r'C:\Users\SCDJ2L\dev\Zoning\OZmapMaker_v0\OZmapMaker_FullTWP.aprx') #TEST
map = aprx.listMaps("Premier Data Frame")[0]
lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]

# currentdir = r'\\snoco\gis\plng\images\zoning_current\pdf'
currentdir = r"C:\Users\SCDJ2L\dev\Zoning\TEST\zoning_current\pdf" #TEST
# archivedir = r'\\snoco\Dept_Images\PDS\PDSHistoricalMapsandAerials\HistoricalZoningMaps'
archivedir = r"C:\Users\SCDJ2L\dev\Zoning\TEST\HistoricalZoningMaps\pdf" #TEST


def get_layout(user_aprx):
    '''Returns a layout object from the APRX, only if it is one of the two valid layouts'''
    layouts_valid = ["FULL_TWP_DDP_OZmap_v1", "TR_DDP_OZmap_v1"]
    layouts_aprx = user_aprx.listLayouts()
    for layout in layouts_aprx:
        if layout.name in layouts_valid:
            return layout
        else:
            arcpy.AddError(f"There are no valid layouts in {user_aprx.filePath}")


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


def delete_previous_PDF(CurPDF):
    '''Deletes a file with a wildcard (i.e. the existing current .PDF, hence glob)'''
    for fl in glob.glob(CurPDF):
        os.remove(fl)
    return


def delete_duplicate_PDF(ExistingPDF):
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


def get_export_filename_section(mpg_parse_list):
    '''Returns four string variables, if exporting PDF files by section'''
    if len(str(mpg_parse_list[1])) == 1:  # section string - if 1-digit Section, then add "0"
        a_sec = "0" + str(mpg_parse_list[1])
    elif len(str(mpg_parse_list[1])) == 2:
        a_sec = str(mpg_parse_list[1])

    a_twp = str(mpg_parse_list[2][1:3])  # township string
    if len(str(mpg_parse_list[3])) > 3:  # range string - if 2-digit range, then slice by [1:3]
        a_rng = str(mpg_parse_list[3][1:3])
    elif len(str(mpg_parse_list[3])) == 3:
        a_rng = "0" + str(mpg_parse_list[3][1:2])

    if len(str(mpg_parse_list[0])) == 3:  # quarter section string
        a_qtr = str(0)
    elif len(str(mpg_parse_list[0])) > 3:
        a_qtr = str(5)
    elif str(mpg_parse_list[0]) == "NE":
        a_qtr = str(1)
    elif str(mpg_parse_list[0]) == "NW":
        a_qtr = str(2)
    elif str(mpg_parse_list[0]) == "SW":
        a_qtr = str(3)
    elif str(mpg_parse_list[0]) == "SE":
        a_qtr = str(4)

    return a_twp, a_rng, a_sec, a_qtr


def get_export_filename_fullTWP(parse_list):
    '''Returns four string variables, if exporting PDF files by for the full towship'''
    a_sec = r"00"                               # section string
    a_twp = str(parse_list[0][1:3])             # township string
    if len(str(parse_list[1])) > 3:             # range string - if 2-digit range, then slice by [1:3]
        a_rng = str(parse_list[1][1:3])
    elif len(str(parse_list[1])) == 3:
        a_rng = "0" + str(parse_list[1][1:2])
    a_qtr = str(0)                              # quarter section string

    return a_twp, a_rng, a_sec, a_qtr


def kill_acrobat():
    # Taskkill to perform a hard close on Adobe Acrobat Reader
    PID_acro32 = isrunning('AcroRd32.exe')
    if PID_acro32 != "None":
        os.system(r'taskkill /F /PID ' + PID_acro32)
        arcpy.AddWarning("Adobe Acrobat Reader was terminated...")

    # Taskkill to perform a hard close on Adobe Acrobat Pro
    PID_acro = isrunning('Acrobat.exe')
    if PID_acro != "None":
        os.system(r'taskkill /F /PID ' + PID_acro)
        arcpy.AddWarning("Adobe Acrobat Pro was terminated...")

    return


def ms_loop():
    '''Loop through map series pages to export to PDF'''
    # Global variables
    global currentdir, archivedir, currentpath, archivepath

    # While loop begins
    pg = 1
    while pg <= ms.pageCount:
        # Change current page name
        ms.currentPageNumber = pg

        # Get QSTR-name of current page from .MasterPageGrid
        mpg_label = ms.pageRow.LABEL2

        # Split QSTR-name label per empty spaces
        mpg_parse_list = mpg_label.split()

        # Get PLSS string which will be included in PDF export filename
        if gpat_layout == "Full Township":
            a_twp, a_rng, a_sec, a_qtr = get_export_filename_fullTWP(mpg_parse_list)
        elif gpat_layout == "Individual Sections":
            a_twp, a_rng, a_sec, a_qtr = get_export_filename_section(mpg_parse_list)

        arcpy.AddMessage(f"Section: {a_sec}, Township: {a_twp}, Range: {a_rng}, Quarter Section: {a_qtr}")
        fname = f"{a_twp}{a_rng}{a_sec}{a_qtr}{str(date_stamp)}.pdf"
        fname_cur = f"{a_twp}{a_rng}{a_sec}{a_qtr}.pdf"
        arcpy.AddMessage("Archived filename: {}".format(fname))
        arcpy.AddMessage("Current filename: {}".format(fname_cur))
        
        kill_acrobat()

        # Remove existing .PDF with previous date from current folder
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname_cur)
        delete_previous_PDF(currentpath)

        # Remove existing .PDF with same names from current and archive folders
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname_cur)
        archivepath = os.path.join(archivedir, a_twp + a_rng, fname)
        arcpy.AddMessage("currentpath: " + currentpath)
        arcpy.AddMessage("archivepath: " + archivepath)

        arcpy.AddMessage("Removing existing PDF files from current and archive folders...")
        delete_duplicate_PDF(currentpath)
        delete_duplicate_PDF(archivepath)

        # Export .PDF to archive folder
        arcpy.AddMessage("Exporting PDF to current folder...")
        layout.exportToPDF(currentpath)

        # Copy .PDF file from current to archive folder
        arcpy.AddMessage("Copying to archive folder...")
        os.system("copy %s %s" % (currentpath, archivepath))

        # While loop increment page number
        pg = pg + 1

    return


#--Read the .MasterPageGrid definition query to label archive and current folder paths
if gpat_continue_cond == "true":
    date_stamp = get_date_stamp()

    # Get layout from the APRX project, only if it is one of two "valid" OZmap layouts
    layout = get_layout(aprx)
    arcpy.AddMessage(f"layout: {layout.name}")

    if not layout.mapSeries is None:
        ms = layout.mapSeries
    else:
        arcpy.AddError(
            "Map series not available in this project's layout! Map series must be created manually in Pro first...")

    if lyr.supports("DEFINITIONQUERY"):
        # --Definition query of .PageGrid_MasterPageGrid
        dq = lyr.definitionQuery
        arcpy.AddMessage("definition query: " + str(dq))

        if dq != "":
            #--Make list of all map series pages
            dq_list = list(find_all(dq, '='))
            arcpy.AddMessage("dq_list: " + str(dq_list))

            #--Loop through map series pages to export
            arcpy.AddMessage("Initiating loop through map series pages")
            ms_loop()

        else:    #--Def. Qry. count
            arcpy.AddError("No definition query found for .MasterPageGrid!")

    else:    #--if lyr.supports
        arcpy.AddError("The feature class does not support definition queries!")
else:     #--if Continue checkbox not checked
    arcpy.AddError("Unchecked parameter box ... Terminating OZ map export process!")
