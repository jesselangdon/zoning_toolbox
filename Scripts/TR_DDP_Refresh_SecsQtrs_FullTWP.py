"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh_SecQtrs_FullTWP.py
        \\snoco\gis\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py
        
        \\snoco\gis\plng\carto\Zoning_OZ_MXD\TESTING\Scripts\TR_DDP_Refresh_SecsQtrsFullTwp.py

Purpose:  Via a Python script tool: 
              * User input of township, range, and OZ sequence number(s)
              * Updates definition query of .PageGrid_MASTERPageGrid layer;
              * Refreshes the map series;
              * Find/replaces the "txtSEQColor" text element with correct color;
              * Generate any OZmap per Qtr-Section, Section, or Full Township scale

Usage:  Python tool scripts with ArcGIS Pro >> DDP_OZmap Toolbar >> "Export DDP to PDF"
             within OZmapMaker.aprx, in the following order...
             1) Run the "Twp Rng DDP Refresh" script tool first, to dial-in the preferred twp & rng;
             2) Run the "Export DDP to PDF" script tool second, to generate and save .PDF files.

History:  20120328 - Renamed and moved "ddp2_cmyktest.py" into production;
          20201104 - Modified Def. Qry. to exclude TWPs, only want Sections and 1/4-Sections; line 95;  ... +  " AND [SECTION] <> 0 " )
          scdwpr, 20120328
          20230829 - Significantly refactored to use Python 3, Def. Qry. per OZSEQNUM, and the arcpy package for ArcGIS Pro, Jesse L.
"""


# Import modules
import arcpy

# Get input
gpat_Twp = arcpy.GetParameterAsText(0)
gpat_Rng = arcpy.GetParameterAsText(1)
oz_string = arcpy.GetParameterAsText(2)
oz_numbers = oz_string.replace(" ", "")
FullTwpTest = str(gpat_Twp) + "-" + str(gpat_Rng)

# Convert input strings to integers
gpat_intTwp = int(gpat_Twp)
gpat_intRng = int(gpat_Rng)
dq = ""

# Crosscheck for making a Full Township map
FULLTWPviaSTRLABEL_List = [7-10, 27-11, 27-12, 27-13, 27-8, 27-9, 28-10, 28-11, 28-12, 28-13, 28-9, 29-10, 29-11, 29-12, 29-13, 29-14, 29-8, 29-9, 30-10, 30-11, 30-12, 30-13, 30-14, 30-15, 30-8, 30-9, 31-10, 31-11, 31-12, 31-13, 31-14, 31-15, 31-16, 31-7, 31-8, 31-9, 32-10, 32-11, 32-12, 32-13, 32-14, 32-15, 32-8, 32-9]

def SetColorTextElem(layout):
    '''Set the color of txtSEQColor to match the particular def. qry. CMYK specs of .MasterPageGrid'''

    #--Global variables
    global dq, cmyk

    #--Current pageID
    ms.currentPageID = 1
    
    #--Get SEQColor of current Page and construct new text to find & replace for Sequence Number
    seq_label = ms.pageRow.CMYK_SEQCOLOR
    cmyk_List = seq_label.split(",")
    
    c_clr = '"' + str(cmyk_List[0]) + '"'
    m_clr = '"' + str(cmyk_List[1]) + '"'
    y_clr = '"' + str(cmyk_List[2]) + '"'
    b_clr = '"' + str(cmyk_List[3]) + '"'
    arcpy.AddMessage("CMYK values: {0}, {1}, {2}, {3}".format(c_clr, m_clr, y_clr, b_clr))
    
    #--Specifiy the layout element to change according to the txtSEQColor text element
    textElem = layout.listElements("TEXT_ELEMENT", "txtSEQColor")
    textElem[0].text = '<CLR cyan = {0} magenta = {1} yellow = {2} black = {3}><dyn type="page" property="ozseqnum"/></CLR>'\
        .format(c_clr, m_clr, y_clr, b_clr)

    return


def setDefinitionQuery(layer, sql):
    layer.updateDefinitionQueries(
        [
            {'name': 'Query 1', 'sql': sql, 'isActive': True}
        ]
    )
    return

#--############################ BEGIN PROCESS #############################

#--Variables
aprx = arcpy.mp.ArcGISProject("CURRENT")
# aprx = arcpy.mp.ArcGISProject(r'C:\Users\SCDJ2L\dev\Zoning\OZmapMaker_v0\OZmapMaker_v20230829.aprx') #TEST
mf = aprx.listMaps("Premier Data Frame")[0]
lyr = mf.listLayers(".PageGrid_MASTERPageGrid")[0]

#layout = aprx.listLayouts()[1]
#ms = layout.mapSeries

lyt_Sec_QtrSec = aprx.listLayouts("TR_DDP_OZmap_v1*")[0]
#layout = aprx.listLayouts()[1]
ms = lyt_Sec_QtrSec.mapSeries

lyt_FullTWP = aprx.listLayouts("FULL_TWP_DDP_OZmap_v1*")[0]


#--CALL TOWNSHIP RANGE FUNCTION. Only perform if township and range is within Snohomish county
if (gpat_intTwp > 26 and gpat_intTwp < 33) and (gpat_intRng > 2 and gpat_intRng < 12):
    dq = "[TOWNSHIP] = {0} AND [RANGE] = {1} AND [OZSEQNUM] IN ({2})".format(gpat_Twp, gpat_Rng, oz_numbers)
    arcpy.AddMessage("Definition query to apply: " + dq)
    if lyr.supports("DEFINITIONQUERY"):
        setDefinitionQuery(lyr, dq)
    ms.refresh()
    layout = lyt_Sec_QtrSec
    SetColorTextElem(layout)
    arcpy.AddMessage("Zoning APRX layout and map series refreshed successfully!")
else: #--if township and range is outside of Snohomish county
    dq = ""
    arcpy.AddError("Township and range are not found in Snohomish County.")


#--Make Full Township Map (if TwpRng is in the hinterlands)
if FullTwpTest in FULLTWPviaSTRLABEL_List:
    dq = "[TOWNSHIP] = {0} AND [RANGE] = {1} AND [QSLABEL] = 'TWP')".format(gpat_Twp, gpat_Rng)
    arcpy.AddMessage("Definition query to apply: " + dq)
    if lyr.supports("DEFINITIONQUERY"):
        setDefinitionQuery(lyr, dq)
    ms.refresh()
    layout = lyt_FullTWP
    SetColorTextElem(layout)
    arcpy.AddMessage("Zoning APRX for Full TWP layout and map series refreshed successfully!... YET, TO DEBUG, SINCE THIS WIPED OUT THE PREVIOUS MAP SERIES... HELP J2L !!!")
else: #--if township and range is outside of Snohomish county
    dq = ""
    arcpy.AddError("Township and range are not found in Snohomish County.")
    
    
    