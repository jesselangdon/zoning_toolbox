"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\Export_DDP_to_PDF_v1.py
%ws% ==>  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\...

Useage:  Python Tool Scripts with ArcMap >> DDP_OZmap Toolbar >> "Export DDP to PDF"
             within W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\TR_DDP_v1.mxd
             1) Run the "Twp Rng DDP Refresh" script tool first, to dial-in the preferred Twp & Rng;
             2) Run the "Export DDP to PDF" script tool second, to generate saved .PDF images.

Purpose:  Within the .MXD that produces the official Zoning Maps, customized Python Script Tools.
              The script that exports each map of the Data Driven Pages to both the archive 
              and current folders storing .PDF images of the official Zoning Maps.
              
History:  20120329 - Final Debugging and moved into production
          20201104 - Modified for the FULL TOWNSHIP map series

scdwpr, 20120330
"""

import arcpy
import os, os.path, datetime, glob, shutil, sys

#--Variables
gpat_continue = arcpy.GetParameterAsText(0)
gpat_continue_cond = str(gpat_continue)
aprx = arcpy.mp.ArcGISProject("CURRENT")
map = aprx.listMaps("Premier Data Frame")[0]
lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]
layout = aprx.listLayouts()[1]
ddp = layout.mapSeries

currentdir = r'\\snoco\gis\plng\images\zoning_current\pdf'
archivedir = r'\\snoco\Dept_Images\PDS\PDSHistoricalMapsandAerials\HistoricalZoningMaps'

print("aprx is: " + str(aprx))
print("map is: " + str(map))
print("lyr is: " + str(lyr))
print("ddp is: " + str(ddp))

#--Date Stamp Suffix
now = datetime.datetime.now()
date_stamp = now.strftime("_%Y%m%d")
print("date_stamp is: " + str(date_stamp))

#--Find_All Defined Function for a listing of where the '=' occurs
def find_all(a_str, sub):
    start = 0
    while True:
        start = a_str.find(sub, start)
        if start == -1: return
        yield start
        start += len(sub)

#--To delete a file with a wildcard (i.e. the existing current .PDF, hence glob)
def ToDelCurPDF(CurPDF):
    for fl in glob.glob(CurPDF):
        os.remove(fl)

def ToRemoveExisting(ExistingPDF):
    if os.path.exists(ExistingPDF):
        print("Removing Existing (i.e. duplicate) .PFD file: " + str(ExistingPDF))
        os.remove(ExistingPDF)

#--Taskkill to perform a hard close on Adobe Acrobat Reader
def isrunning(exe):
    try :
        p = os.popen(r'tasklist /FI "IMAGENAME eq "'+ exe + ' /FO "LIST" 2>&1' , 'r' )
        PID = p.read().split('\n')[2].split(":")[1].lstrip(" ")
        p.close()
        return PID
    except :
        p.close()
        return "None"

#--Def. -- Loop through DDP pages
def ddp_loop():
    #--Global variables
    global currentdir, archivedir, currentpath, archivepath

    #--While loop begin
    pg = 1
    while pg <= ddp.pageCount:
        #--Change current page name
        ddp.currentPageID = pg

        #--Get QSTR-name of current Page from .MasterPageGrid
        # mpg_label = ddp.pageRow.getValue('LABEL2')   #--Field from DDP layer (i.e. .MasterPageGrid)
        mpg_label = ddp.pageRow.LABEL2
        print("____________________")
        print("mpg_label is: " + mpg_label)

        #--Split LABEL per empty spaces
        mpg_parseList = mpg_label.split()
        #print mpg_parseList

        #--Sectional

        a_sec = r"00"                                       #--Sec. Str

        a_twp = str(mpg_parseList[0][1:3])                  #--Twp. Str

        if len(str(mpg_parseList[1])) > 3:                  #--Rng. Str - If... 2-digit Range, then slice by [1:3]
            a_rng = str(mpg_parseList[1][1:3])
        elif len(str(mpg_parseList[1])) == 3:
            a_rng = "0" + str(mpg_parseList[1][1:2])

        a_qtr = str(0)                                      #--Qtr. Str -
            

        print("a_sec is: " + a_sec)
        print("a_twp is: " + a_twp)    
        print("a_rng is: " + a_rng)
        print("a_qtr is: " + a_qtr)
        fname = a_twp + a_rng + a_sec + a_qtr + str(date_stamp) + ".pdf"
        fname_cur = a_twp + a_rng + a_sec + a_qtr + "_*.pdf"
        print("*** fname is: " + fname)
        print("*** fname_cur is: " + fname_cur)        
        #arcpy.addmessage("OKoooo")
        
        #--Taskkill to perform a hard close on Adobe Acrobat Reader
        PID = isrunning('AcroRd32.exe')
        if PID != "None" :
            os.system(r'taskkill /F /PID ' + PID)
            
        #--Taskkill to perform a hard close on Adobe Acrobat Pro
        PID = isrunning('Acrobat.exe')
        if PID != "None" :
            os.system(r'taskkill /F /PID ' + PID)

        #--Remove Existing .PFD of a previous date from Current folder {with wildcard, hence glob}
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname_cur)
        print("currentpath is: " + currentpath)

        ToDelCurPDF(currentpath)

        #--Remove Existing .PFD of same names from both Current & Archive folders
        currentpath = os.path.join(currentdir, a_twp + a_rng, fname)
        archivepath = '"' + os.path.join(archivedir, a_twp + a_rng, fname) + '"'
        print("currentpath is: " + currentpath)
        print("archivepath is: " + archivepath)

        ToRemoveExisting(currentpath)
        ToRemoveExisting(archivepath)

        #--Export .PDF to Archive folder
        layout.exportToPDF(currentpath)
        print("=-=-=-=-=- To copy Currentpath over to Archivepath")

        #--Copy .PDF file from Current over to Archive folder
        os.system("copy %s %s" % (currentpath, archivepath))

        print("+++++++++ Did it Copy over to Current & Archive Directories???")

        #--While loop-- Increment page number
        pg = pg + 1
    return

#----MAIN BODY (then calls Definition Functions)--------
#--Read the .MasterPageGrid Definition Query to label Archive & Current Folder Paths

if gpat_continue_cond == "true":
    if lyr.supports("DEFINITIONQUERY"):
        dq = lyr.definitionQuery      #-- Def. Qry of .PageGrid_MasterPageGrid
        print("dq is: " + str(dq))

        if dq != "":
            #--Make List of all DataDrivenPages for given 
            dq_list = list(find_all(dq, '='))
            print("dq_list is: " + str(dq_list))

            #--(WHILE) Loop through DDP pages
            print("To do the WHILE loop through DDP pages")
            ddp_loop()

        else:    #--Def. Qry. count
            print("~~~~ No Definition Query found for .MasterPageGrid ~~~~ ")

    else:    #--if lyr.supports
        print(" ~~~~ This Feature Class doesn't support Definition Queries ~~~~")
else:     #--if Continue checkbox
    print("~~~~ Unchecked Parameter Box, ... Bailing out of Script Tool ~~~~ ")
    
#--Clean up variables
del aprx, map, lyr, layout, ddp, currentdir, archivedir, now, date_stamp

"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\ConfigurationTips__DDP_OZmap__toolbar.txt
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Configuration Tips for "DDP_OZmap" toolbar:

1) In ArcCatalog, 
  a) Add New Toolbox (.tbx)  {if need be}
  b) Rt-click on Toolbox (.tbx) >> Add >> Script   {choose the .PY file's path}
    1) Rt-click on Script tool >> General tab >> set Name and Label
    2) ... >> Source tab >> Set "Script File" path; and Check both boxes, esp. "Run Python script in process"
    3 ... >> Parameters tab >> Add Display Name and Data Type...
      a) e.g Export DDP to PDF Properties,"ATTN: This tool will..."; Boolean
      b) e.g. Twp Rng DDP Refresh Properities, "Township {b/w 27 and 32}"; String
        1) Set Property and Values  {Requied; Input; No; 30; blank; None; blank; blank}
      c) e.g. Twp Rng DDP Refresh Properities, "Range {b/w 3 and 11}"; String
        2) Set Property and Values  {Requied; Input; No; 7; blank; None; blank; blank}
      d) Keep all the the Validation tab, click OK.


2) In ArcMap, Customize >> Toolbars tab >>  New... >> DDP_OZmap   {just need to establish once}
  a) ... >> Commands tab >> [Geoprocessing Tools] near bottom of list >> Add Tools... {select the .PY file}
  b) Drag the newly added Tool into the appropriate "DDP_OZmap" Toolbar
    1) Rt-click on the Tool in the Toolbar, check...
      a) Image and Text
      b) Begin New Group

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



======================================
ArcMap Python Window Feedback during Debugging
======================================
mxd is: <geoprocessing Map object object at 0x055F8340>
df is: <geoprocessing Data Frame object object at 0x055FE070>
lyr is: .PageGrid_MASTERPageGrid
ddp is: <geoprocessing Page Layout DDP object object at 0x054A7FE0>
date_stamp is: _20120305
dq is: (TOWNSHIP = 30 AND RANGE = 7) and SECTION < 3
dq_list is: [10, 25]
To do the WHILE loop through DDP pages
____________________
mpg_label is: SEC 1 T30N R7E
cmyk_List is: [u'0', u'50', u'86', u'18']
cmyk_ables are................: "0""50""86""18"
a_sec is: 01
a_twp is: 30
a_rng is: 07
a_qtr is: 0
*** fname is: 3007010_20120305.pdf
~~~~~~~ To Export DDP to PDF
currentpath is: C:\AA_pdf\3007\3007010_20120305.pdf
archivepath is: C:\temp\3007\3007010_20120305.pdf
+++++++++ To Copy fname to Current & Archive Directories
____________________
mpg_label is: SEC 2 T30N R7E
cmyk_List is: [u'0', u'50', u'86', u'18']
cmyk_ables are................: "0""50""86""18"
a_sec is: 02
a_twp is: 30
a_rng is: 07
a_qtr is: 0
*** fname is: 3007020_20120305.pdf
~~~~~~~ To Export DDP to PDF
currentpath is: C:\AA_pdf\3007\3007020_20120305.pdf
archivepath is: C:\temp\3007\3007020_20120305.pdf
+++++++++ To Copy fname to Current & Archive Directories
"""