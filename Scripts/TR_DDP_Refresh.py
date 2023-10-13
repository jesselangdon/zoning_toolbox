"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py
        \\snoco\gis\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py

Purpose:  Via a Python script tool: 
              * User input of township, range, and OZ sequence number(s)
              * Updates definition query of .PageGrid_MASTERPageGrid layer;
              * Refreshes the map series;
              * Find/replaces the "txtSEQColor" text element with correct color;

Usage:  Python tool scripts with ArcGIS Pro >> DDP_OZmap Toolbar >> "Export DDP to PDF"
             within OZmapMaker.aprx
             1) Run the "Twp Rng DDP Refresh" script tool first, to dial-in the preferred twp & rng.
             2) Run the "Export DDP to PDF" script tool second, to generate and save .PDF files.

History:  20120328 - Renamed and moved "ddp2_cmyktest.py" into production;
          20201104 - Modified Def. Qry. to exclude TWPs, only want Sections and 1/4-Sections; line 95;  ... +  " AND [SECTION] <> 0 " )
          scdwpr, 20120328
          20230829 - Significantly refactored to use Python 3 and the arcpy package for ArcGIS Pro, Jesse L.
          20231003 - Updated so that a specific layout "TR_DDP_OZmap_v1" is present, and then pulls the map series from that, Jesse L.
"""


# Import modules
import arcpy

# Get input
gpat_Twp = arcpy.GetParameterAsText(0)
gpat_Rng = arcpy.GetParameterAsText(1)
oz_string = arcpy.GetParameterAsText(2)
oz_numbers = oz_string.replace(" ", "")

# Convert input strings to integers
gpat_intTwp = int(gpat_Twp)
gpat_intRng = int(gpat_Rng)
dq = ""

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
# aprx = arcpy.mp.ArcGISProject(r'\\snoco\gis\plng\carto\Zoning_OZ_MXD\TESTING\OZmapMaker_v0\OZmapMaker_v20230925.aprx') #TEST
map = aprx.listMaps("Premier Data Frame")[0]
lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]
layout = aprx.listLayouts("TR_DDP_OZmap_v1")[0]
ms = layout.mapSeries

#--CALL TOWNSHIP RANGE FUNCTION. Only perform if township and range is within Snohomish county
if (gpat_intTwp > 26 and gpat_intTwp < 33) and (gpat_intRng > 2 and gpat_intRng < 12):
    dq = "[TOWNSHIP] = {0} AND [RANGE] = {1} AND [OZSEQNUM] IN ({2})".format(gpat_Twp, gpat_Rng, oz_numbers)
    arcpy.AddMessage("Definition query to apply: " + dq)
    if lyr.supports("DEFINITIONQUERY"):
        setDefinitionQuery(lyr, dq)
    ms.refresh()
    SetColorTextElem(layout)
    arcpy.AddMessage("Zoning APRX layout and map series refreshed successfully!")
else: #--if township and range is outside of Snohomish county
    dq = ""
    arcpy.AddError("Township and range are not found in Snohomish County.")