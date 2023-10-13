"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh_SecQtrs_FullTWP.py
        \\snoco\gis\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py

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
import json
import arcpy

# Get input
gpat_Twp = arcpy.GetParameterAsText(0)
gpat_Rng = arcpy.GetParameterAsText(1)
twp_rng = f"{gpat_Twp}-{gpat_Rng}"

# Convert input strings to integers
gpat_intTwp = int(gpat_Twp)
gpat_intRng = int(gpat_Rng)
dq = ""

# Import list of townships that need a full layout (i.e. not quarter section layout)
with open(r"..\bin\twp_list.json", 'r') as f:
    twp_list = json.load(f)

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
# aprx = arcpy.mp.ArcGISProject(r'C:\Users\SCDJ2L\dev\Zoning\OZmapMaker_v0\OZmapMaker_v20231011.aprx') #TEST
mf = aprx.listMaps("Premier Data Frame")[0]
lyr = mf.listLayers(".PageGrid_MASTERPageGrid")[0]

# Determine if user-selected township/range should use the full township layout, or quarter section layout
if twp_rng in twp_list:  #FIXME The problem with this is that there are often two features in the map grid where QSLABEL  = 'FULL' AND QSLABEL = 'TWP'
    # use the full township layout
    lyt_full_twp = aprx.listLayouts("FULL_TWP_DDP_OZmap_v1")[0]
    ms = lyt_full_twp.mapSeries
    dq = "[TOWNSHIP] = {0} AND [RANGE] = {1} AND [QSLABEL] = 'TWP'".format(gpat_Twp, gpat_Rng)
    arcpy.AddMessage("Definition query to apply: " + dq)
    if lyr.supports("DEFINITIONQUERY"):
        setDefinitionQuery(lyr, dq)
    ms.refresh()
    SetColorTextElem(lyt_full_twp)
    arcpy.AddMessage("Full township layout and map series refreshed successfully!")
else:
    # use the quarter section layout
    lyt_qrt_sec = aprx.listLayouts("TR_DDP_OZmap_v1")[0]
    ms = lyt_qrt_sec.mapSeries
    dq = f"[TOWNSHIP] = {gpat_Twp} AND [RANGE] = {gpat_Rng} AND [QSLABEL] = 'FULL'"
    arcpy.AddMessage("Definition query to apply: " + dq)
    if lyr.supports("DEFINITIONQUERY"):
        setDefinitionQuery(lyr, dq)
    ms.refresh()
    SetColorTextElem(lyt_qrt_sec)
    arcpy.AddMessage("Quarter section layout and map series refreshed successfully!")