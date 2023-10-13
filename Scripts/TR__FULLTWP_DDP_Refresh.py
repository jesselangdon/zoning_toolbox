"""
File:  W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py
        \\snoco\gis\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\TR_DDP_Refresh.py
        W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\Scripts\ConfigurationTips__DDP_OZmap__toolbar.txt

Purpose:  Via a Python script tool: 
              * User input of Twp & Rng, 
              * Updates Def. Qry of .PageGrid_MASTERPageGrid;
              * Refreshes DDP;
              * Find/replaces the "txtSEQColor" text element with correct color;
              * Refreshes TOC & ActiveView.

Useage:  Python Tool Scripts with ArcMap >> DDP_OZmap Toolbar >> "Export DDP to PDF"
             within W:\plng\carto\Zoning_OZ_MXD\Current_OZmap_MXD\TR_DDP_v1.mxd
             1) Run the "Twp Rng DDP Refresh" script tool first, to dial-in the preferred Twp & Rng;
             2) Run the "Export DDP to PDF" script tool second, to generate saved .PDF images.

History:  20120328 - Renamed and moved "ddp2_cmyktest.py" into production;
            20201104 - Modified Def. Qry. to exclude TWPs, only want Sections and 1/4-Sections; ~ line 96;  ... +  " AND [SECTION] = 0 " );
            20201104 - Turn off def SetColorTextElem() function, since there is no said Text Element in FULL TWP layout; ~ line 116.

scdwpr, 20120328
"""

#--Import system modules
import arcpy

##--GET INPUT--##
#--Get Parameters Twp & Rng from .APRX
gpat_Twp = arcpy.GetParameterAsText(0)
gpat_Rng = arcpy.GetParameterAsText(1)
#--Convert input strings to integers
gpat_intTwp = int(gpat_Twp)
gpat_intRng = int(gpat_Rng)
testTownship = ""
dq = ""

#--Def.--Set the color of txtSEQColor to match the particular Def. Qry. CMYK specs of .MasterPageGrid         
def SetColorTextElem():
    #--Global variables
    global dq, cmyk

    #--Current pageID
    ddp.currentPageID = 1
    
    #--Get SEQColor of current Page and construct new text to find & replace for Sequence Number
    seq_label = ddp.pageRow.getValue('CMYK_SEQCOLOR')   #--Field from DDP layer (i.e. .MasterPageGrid)
    #print "seq_label is: " + seq_label
    cmyk_List = seq_label.split(",")
    print("cmyk_List is: " + str(cmyk_List))
    
    c_clr = '"' + str(cmyk_List[0]) + '"'
    m_clr = '"' + str(cmyk_List[1]) + '"'
    y_clr = '"' + str(cmyk_List[2]) + '"'
    b_clr = '"' + str(cmyk_List[3]) + '"'
    print("c_clr is................: " + str(c_clr))
    print("m_clr is................: " + str(m_clr))
    print("y_clr is................: " + str(y_clr))
    print("b_clr is................: " + str(b_clr))
    print("cmyk_ables is................: " + str(c_clr) + str(m_clr) + str(y_clr) + str(b_clr))
    
    #--Specifiy the layout element to change according to the txtSEQColor text element
    aprx = arcpy.mp.ArcGISProjects("CURRENT")
    layout = aprx.listLayouts()[1]
    textElem = layout.listElements("TEXT_ELEMENT", "textSEQColor")
    #--e.g.  <CLR cyan = "100" magenta = "10" yellow = "100" black = "50"><dyn type="page" property="ozseqnum"/></CLR>
    textElem[0].text = '<CLR cyan = %s magenta = %s yellow = %s black = %s><dyn type="page" property="ozseqnum"/></CLR>' %(c_clr, m_clr, y_clr, b_clr)
    #print "+++++++++ Did the Sequence Color change???"    

#--########################################################################
#--############################ BEGIN PROGRAM #############################
#--########################################################################
#--Variables
aprx = arcpy.mp.ArcGISProject("CURRENT")
map = aprx.listMaps("Premier Data Frame")[0]
lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]
layout = aprx.listLayouts()[1]
ddp = layout.mapSeries

legit_TownshipsList = [2708, 2709, 2710, 2711, 2712, 2713, 2809, 2810, 2811, 2812, 2813, 2908, 2909, 2910, 2911, 2912, 2913, 2914, 3008, 3009, 3010,
                               3011, 3012, 3013, 3014, 3015, 3107, 3108, 3109, 3110, 3111, 3112, 3113, 3114, 3115, 3116, 3208, 3209, 3210, 3211, 3212, 3213, 3214, 3215]

Twp = str(gpat_Twp)

if len(str(gpat_intRng)) == 1:
    Rng = "0" + str(gpat_intRng)
    print(("Rng_1 is: " + Rng))
else:     #--if Rng is 1-character
    Rng = str(gpat_intRng)
    print(("Rng_2 is: " + Rng))

TwpRng = Twp + Rng
print(("TwpRng = " + TwpRng))

 
#--CALL TOWNSHIP RANGE FUNCTION
#--Only Perform in TR is within SnoCo.

if int(TwpRng) in legit_TownshipsList:
    dq = str("[TOWNSHIP] = " + Twp + " AND [RANGE] = " + Rng +  " AND [QSLABEL] = 'TWP' " )
    print(("dq1 is: " + str(dq)))
    
else:       #--if Twp & Rng is outside of SnoCo.
    dq = ""

#--CALL DEFINITION QUERY FUNCTION
if dq != "":
    print("dq2 is:" + dq)
    print("~~~ DQ'ing ??? ~~~")
    #--Only change Def. Qry. if layer supports a Def. Qry.
    for lyr in map.listLayers(".PageGrid_MASTERPageGrid"):
        if lyr.supports("DEFINITIONQUERY"):
            print("~~~ DQ Supporting ??? ~~~")
            lyr.definitionQuery = dq
            print(f"dq3 is: {dq}")
            dq = lyr.definitionQuery
            print(f"dq4 is: {dq}")

    ddp.refresh()
    #SetColorTextElem()  #--It works with TR_DDP_Refresh.py; YET NOT to use with TR__FULLTWP_DDP_Refresh.py!!! ********************************

#--if dq is empty
while dq=="":    
    print ("*** This Twp & Rng is not a legit FULL TOWNSHIP per _MasterPageGrid, please redo ***")
    #--Break out of program
    break

#print " ~~~~ SAVE ISSUE....??? ~~~~"

#--Delete variables
del aprx, map, lyr, layout, ddp, dq, gpat_Twp, gpat_Rng, gpat_intTwp, gpat_intRng, Twp, Rng, TwpRng, legit_TownshipsList
