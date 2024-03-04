﻿"""Title:          OZ Map ToolboxFilename:       OZMapToolbox.pytDescription:    The OZ Map Toolbox is a Python toolbox that is intended to simplify and automate the process of updating PDS OZ Maps,                which is a map series depicting zoning for Snohomish County that is maintained by the Planning and Development Services                (PDS) department. The Toolbox is intended for use with ArcGIS Pro 3.x or higherDependencies:   Python 3, arcpy package, and ArcGIS Pro 3.xAuthor:         Jesse LangdonLast Updated:   1/3/2024 """# -*- coding: utf-8 -*-# import packagesimport osimport datetimeimport jsonimport arcpy# global variablescurrentdir = r"\\snoco\gis\plng\images\zoning_current\pdf"archivedir = r"\\snoco\Dept_Images\PDS\PDSHistoricalMapsandAerials\HistoricalZoningMaps"class Toolbox(object):    def __init__(self):        """OZ Map Toolbox includes tools to update and maintain PDF files that are part of the Snohomish County           Planning and Development Services OZ Map series."""        self.label = "OZ Map Toolbox"        self.alias = "OZ Map Toolbox"        # List of tool classes associated with this toolbox        self.tools = [RefreshMapSeriesTool, ExportMapSeriesTool]class RefreshMapSeriesTool(object):    def __init__(self):        """Define the tool (tool name is the name of the class)."""        self.label = "1 - Refresh OZ Map Series"        self.description = "Refreshes the map series within ArcGIS Pro, either for a whole township, all sections " \                           "within the township, or a subset of sections."        self.canRunInBackground = False    def getParameterInfo(self):        """Define parameter definitions"""        param0 = arcpy.Parameter(            displayName="Township",            name="in_township",            datatype="GPString",            parameterType="Required",            direction="Input")        param0.filter.type = "ValueList"        param0.filter.list = ["27", "28", "29", "30", "31", "32"]        param1 = arcpy.Parameter(            displayName="Range",            name="in_range",            datatype="GPString",            parameterType="Required",            direction="Input")        param1.filter.type = "ValueList"        param1.filter.list = ["3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"]        param2 = arcpy.Parameter(            displayName="Choose a refresh option",            name="refresh_option",            datatype="GPString",            parameterType="Optional",            direction="Input")        param2.filter.type = "ValueList"        param2.filter.list = ["Refresh single map of full township for chosen township and range",                              "Refresh all section/qtr-section maps for chosen township and range",                              "Refresh maps for individual sections"]        param3 = arcpy.Parameter(            displayName="Refresh maps for individual sections using OZ sequence numbers (ex: 27, 28, 29)",            name="ozseqnum",            datatype="GPString",            parameterType="Optional",            direction="Input")        params = [param0, param1, param2, param3]        return params    def isLicensed(self):        """Set whether tool is licensed to execute."""        return True    def updateParameters(self, params):        """Modify the values and properties of parameters before internal        validation is performed.  This method is called whenever a parameter        has been changed."""        # This boolean determines if sections will be updated        if params[2].valueAsText == "Refresh maps for individual sections":            params[3].enabled = True        else:            params[3].enabled = False        return    def updateMessages(self, params):        """Modify the messages created by internal validation for each tool        parameter. This method is called after internal validation."""        return    def execute(self, params, messages):        """The source code of the tool."""        # Local variables        twp, rng, refresh_option_selected, ozseq_str = parse_refresh_params(params)        ozseq = clean_string(ozseq_str)        aprx = arcpy.mp.ArcGISProject("CURRENT")        aprx.closeViews()        # Choose which layout to use        if refresh_option_selected == "Refresh single map of full township for chosen township and range" and found_in_twp_list(twp, rng):            arcpy.AddMessage("Refreshing a single map for the selected township and range...")            lyr, layout, ms, map = get_pro_project_params("FULL_TWP_MapSeries_OZmap")            dq = f"[TOWNSHIP]={twp} AND [RANGE]={rng} AND [QSLABEL] = 'TWP'"            set_def_query(lyr, dq)            layout.openView()            ms.refresh()        elif refresh_option_selected == "Refresh all section/qtr-section maps for chosen township and range":            arcpy.AddMessage("The map series will be refreshed with all available section/qtr-section maps in the township...")            lyr, layout, ms, map = get_pro_project_params("Section_MapSeries_OZmap")            dq = f"[TOWNSHIP]={twp} AND [RANGE]={rng} AND [QSLABEL] <> 'TWP'"            set_def_query(lyr, dq)            layout.openView()            ms.refresh()        elif refresh_option_selected == "Refresh maps for individual sections":            arcpy.AddMessage("Refreshing a subset of sections...")            lyr, layout, ms, map = get_pro_project_params("Section_MapSeries_OZmap")            dq = f"[TOWNSHIP]={twp} AND [RANGE]={rng} AND [OZSEQNUM] IN ({ozseq})"            set_def_query(lyr, dq)            layout.openView()            ms.refresh()        elif refresh_option_selected == "Refresh maps for individual sections" and not found_in_twp_list(twp, rng) and ozseq is None:            arcpy.AddError("The OZ sequence numbers for section maps were not supplied. Add OZ sequence numbers and rerun.")        else:            arcpy.AddWarning("You chose a refresh option that resulted in no updates to the map series...")        # set_seqcolor(map, layout, ms)        arcpy.AddMessage("Zoning APRX layout and map series refreshed successfully!")        return    def postExecute(self, parameters):        """This method takes place after outputs are processed and        added to the display."""        returnclass ExportMapSeriesTool(object):    def __init__(self):        """Define the tool (tool name is the name of the class)."""        self.label = "2 - Export OZ Map Series"        self.description = r"Exports each map in the current map series to PDF files, to both current and archive" \                           r"directories on the \\snoco\gis\plng file server."        self.canRunInBackground = False    def getParameterInfo(self):        """Define parameter definitions"""        param0 = arcpy.Parameter(            displayName="Please choose which layout map series to export",            name="layout",            datatype="GPString",            parameterType="Required",            direction="Input")        param0.filter.type = "ValueList"        param0.filter.list = ["Full township", "Section"]        param1 = arcpy.Parameter(            displayName="ATTENTION: This export will overwrite files in the Current folder. "                        "Please confirm that this is intended...",            name="validate",            datatype="GPBoolean",            parameterType="Required",            direction="Input")        params = [param0, param1]        return params    def isLicensed(self):        """Set whether tool is licensed to execute."""        return True    def updateParameters(self, parameters):        """Modify the values and properties of parameters before internal        validation is performed.  This method is called whenever a parameter        has been changed."""        return    def updateMessages(self, parameters):        """Modify the messages created by internal validation for each tool        parameter.  This method is called after internal validation."""        return    def execute(self, params, messages):        """The source code of the tool."""        layout_choice, validated = parse_export_params(params)        valid_layouts_list = ["FULL_TWP_MapSeries_OZmap", "Section_MapSeries_OZmap"]        if validated:            date_stamp = get_date_stamp()            # Get layout from the APRX project, only if it is one of two "valid" OZmap layouts            layout_name = validate_layout(layout_choice, valid_layouts_list)            lyr, layout, ms, map = get_pro_project_params(layout_name)            if lyr.supports("DEFINITIONQUERY"):                dq = lyr.definitionQuery  # Definition query of .PageGrid_MasterPageGrid                arcpy.AddMessage(f"Definition query: {str(dq)}")                # Loop through map series pages to export                arcpy.AddMessage("Initiating process loop through map series pages...")                ms_loop(ms, layout_choice, date_stamp, layout)            else:  # if lyr.supports not supported                arcpy.AddError("The feature class does not support definition queries!")        else:  # if continue checkbox not checked            arcpy.AddError("Unchecked validation checkbox ... Terminating OZ map export process!")        return    def postExecute(self, parameters):        """This method takes place after outputs are processed and        added to the display."""        return# helper functions -----------------------------------------------------------------------------------------------------def clean_string(list_as_string):    """This function assumes a list is provided as a string. Any spaces between elements are removed."""    if list_as_string is not None:        return list_as_string.replace(" ", "")    else:        arcpy.AddMessage("No OZ sequence number provided...")        returndef parse_twp_json():    """Opens the twp_list.json file and returns a list object."""    with open(os.path.join(os.path.dirname(__file__), "twp_list.json"), 'r') as f:        twp_list = json.load(f)    return twp_listdef parse_refresh_params(params):    """Parses input parameters from the Refresh Map Series Tool into text strings and boolean values."""    township_str = params[0].valueAsText    range_str = params[1].valueAsText    refresh_option = params[2].valueAsText    oz_seq_str = params[3].valueAsText    return township_str, range_str, refresh_option, oz_seq_strdef parse_export_params(params):    """Parses input parameters from the Export Map Series Tool into text strings and boolean values."""    layout_str = params[0].valueAsText    validation = params[1].value    return layout_str, validationdef get_pro_project_params(layout_name):    """Returns a layer, layout, and map series object from the current ArcGIS Pro project."""    active_map = "CURRENT"    aprx = arcpy.mp.ArcGISProject(active_map)    mapx = aprx.listMaps("Premier Data Frame")[0]    lyrx = mapx.listLayers(".PageGrid_MASTERPageGrid")[0]    layoutx = aprx.listLayouts(layout_name)[0]    msx = layoutx.mapSeries    return lyrx, layoutx, msx, mapxdef found_in_twp_list(township, range):    """Checks if the input township and range string values are found within the twp_list.json file."""    twp_list = parse_twp_json()    twp = f"{township}-{range}"    if twp in twp_list:        return True    else:        return Falsedef set_def_query(lyr, dq_new):    """Sets (or resets) the definition query for the input layer"""    lyr.definitionQuery = None    check_def_query(lyr, dq_new) # ensure the defn. query actually returns selected features    if lyr.supports("DEFINITIONQUERY"):        lyr.updateDefinitionQueries(            [                {'name': 'Query 1', 'sql': dq_new, 'isActive': True}            ]        )    returndef check_def_query(lyr, dq):    """Checks if the layer definition query returns a count of feature records > 0"""    arcpy.SelectLayerByAttribute_management(in_layer_or_view=lyr, selection_type="NEW_SELECTION", where_clause=dq)    selected_count = arcpy.GetCount_management(lyr)    if int(selected_count[0]) > 0:        arcpy.AddMessage(f"{selected_count[0]} features selected by the definition query...")    else:        arcpy.AddError(            "No features selected by the definition query! Verify you are using valid OZ sequence numbers!")    returndef set_seqcolor(map, layout, map_series, index_layer_name = ".PageGrid_MASTERPageGrid"):    """Set color of the txtSEQColor element to match the particular CMYK specs of .MasterPageGrid for current page        in the map series. This uses the APRX's layout CIM object to access the text element color."""    # Get the  "Name Field" value for the current page in the map series    current_label = map_series.pageRow.LABEL2 # LABEL2 is the name of the attribute field used to index the map series    # Access the .PageGrid_MASTERPageGrid layer    page_grid_layer = map.listLayers(index_layer_name)[0]    # Query the index layer feature corresponding to the current map series page    qry = f"LABEL2 = '{current_label}'"    with arcpy.da.SearchCursor(page_grid_layer, ["CMYK_SEQCOLOR"], qry) as cursor:        for row in cursor:            cmyk_seqcolor = row[0]            break    # Split the CMYK string values into a list    cmyk_list = [int(value.strip()) for value in cmyk_seqcolor.split(",")]    # Access the page number text element from the layout CIM object    cim_layout = layout.getDefinition('V3')    cim_text_elem = next((e for e in cim_layout.elements if e.name == "txtSEQColor"), None)    # TODO create function to convert CMYK to RGB on-the-fly    # Modify the CIM data (change CMYK color) and apply changes back to text element    if cim_text_elem:        cim_text_elem.graphic.symbol.symbol.symbol.symbolLayers[0].color.values = cmyk_list        layout.setDefinition(cim_layout)        arcpy.AddMessage("Page number color updated...")    else:        arcpy.AddWarning("Text element 'txtSEQColor' not found in layout...")    returndef get_valid_layout(user_aprx, layout_list):    """Returns a layout object from the APRX, only if it is one of the two valid layouts."""    layouts_aprx = user_aprx.listLayouts()    for layout in layouts_aprx:        if layout.name in layout_list:            return layout        else:            arcpy.AddError(f"There are no valid layouts in {user_aprx.filePath}")    returndef validate_layout(layout_selection, layout_list):    """Matches the user-supplied layout choice with a valid layout name."""    if layout_selection == "Full township":        return layout_list[0]    else:        return layout_list[1]def get_date_stamp():    """Generates a date stamp suffix string for the current date"""    now = datetime.datetime.now()    date_stamp = now.strftime("_%Y%m%d")    arcpy.AddMessage("Date stamp: " + str(date_stamp))    return date_stampdef find_all(a_str, sub):    """Function for finding a listing of where the '=' occurs"""    start = 0    while True:        start = a_str.find(sub, start)        if start == -1:            return        yield start        start += len(sub)def delete_matching_pdf(file_path):    """Deletes any existing PDF files in the supplied filepath, based on the embedded 7-digit PLSS value"""    # Extract directory and base file name    directory, file_name = os.path.split(file_path)    # Extract the 7-digit PLSS values    plss_string = file_name.split('_')[0]    # Search and delete matching files    for filename in os.listdir(directory):        if filename.startswith(plss_string) and filename.endswith('.pdf'):            os.remove(os.path.join(directory, filename))            arcpy.AddMessage(f"Found and removed PDF file: {filename}")    returndef isrunning(exe):    """This function uses Taskkill to perform a hard close on running instance of Adobe Acrobat Reader application"""    try:        p = os.popen(r'tasklist /FI "IMAGENAME eq "'+ exe + ' /FO "LIST" 2>&1' , 'r' )        PID = p.read().split('\n')[2].split(":")[1].lstrip(" ")        p.close()        return PID    except:        p.close()        return "None"def get_export_filename(page_trs_q):    """Function returns four string variables, if exporting PDF files for section maps"""    a_twp = page_trs_q[0:2]    a_rng = page_trs_q[2:4]    a_sec = page_trs_q[4:6]    a_qtr = page_trs_q[-1]    return a_twp, a_rng, a_sec, a_qtrdef kill_acrobat():    # Taskkill to perform a hard close on Adobe Acrobat Reader    PID_acro32 = isrunning('AcroRd32.exe')    if PID_acro32 != "None":        os.system(r'taskkill /F /PID ' + PID_acro32)        arcpy.AddWarning("Adobe Acrobat Reader was terminated...")    # Taskkill to perform a hard close on Adobe Acrobat Pro    PID_acro = isrunning('Acrobat.exe')    if PID_acro != "None":        os.system(r'taskkill /F /PID ' + PID_acro)        arcpy.AddWarning("Adobe Acrobat Pro was terminated...")    returndef parse_export_filename(a_twp, a_rng, a_sec, a_qtr, date_stamp):    """Parses a string to be used as a filename for exported PDF files."""    arcpy.AddMessage(f"Section: {a_sec}, Township: {a_twp}, Range: {a_rng}, Quarter Section: {a_qtr}")    fname = f"{a_twp}{a_rng}{a_sec}{a_qtr}{str(date_stamp)}.pdf"    arcpy.AddMessage(f"Export filename:{fname}")    return fnamedef update_export_filepath(a_twp, a_rng, filename):    """Returns a string representing the current and archive filepath of exported PDF files."""    current_fpath = os.path.join(currentdir, a_twp + a_rng, filename)    archive_fpath = os.path.join(archivedir, a_twp + a_rng, filename)    arcpy.AddMessage(f"Export current filepath: {current_fpath} | archive filepath: {archive_fpath}")    return current_fpath, archive_fpathdef export_pdf(layout, currentpath, archivepath):    """Function exports PDFs for map series associated with the ArcGIS Pro layout. Also copies the file to the archive folder."""    # Export .PDF to archive folder    arcpy.AddMessage("Exporting PDF to current folder...")    layout.exportToPDF(currentpath)    # Copy .PDF file from current to archive folder    arcpy.AddMessage("Copying to archive folder...")    os.system("copy %s %s" % (currentpath, archivepath))    returndef check_for_export_selection_mismatch(pg_QSLABEL, layout_selection):    """This function checks if the map in a map series matches the selected ArcGIS Pro layout."""    if pg_QSLABEL != "TWP" and layout_selection == "Section":        return True    elif pg_QSLABEL == "TWP" and layout_selection == "Full township":        return True    else:        return Falsedef ms_loop(ms, layout_selection, date_stamp, layout):    """Loop through map series pages to export each map to PDF files."""    pg = 1    while pg <= ms.pageCount:        # Change current page name        ms.currentPageNumber = pg        # Get attributes from .PageGrid_MASTERPageGrid layer        pg_QSLABEL = ms.pageRow.QSLABEL        pg_parse_list = ms.pageRow.TRS_Q        if check_for_export_selection_mismatch(pg_QSLABEL, layout_selection):            # Get PLSS strings from map series which will be included in PDF export filename            a_twp, a_rng, a_sec, a_qtr = get_export_filename(pg_parse_list)            fname = parse_export_filename(a_twp, a_rng, a_sec, a_qtr, date_stamp)            current_filepath, archive_filepath = update_export_filepath(a_twp, a_rng, fname)            kill_acrobat()            # Remove existing .PDF with same names from current and archive folders            delete_matching_pdf(current_filepath)            export_pdf(layout, current_filepath, archive_filepath)            # While loop increment page number            pg = pg + 1        else:            arcpy.AddError(                "There is a mismatch between the layout selected for export and the map series! "                "Please be sure to run the Refresh the OZ Map Series tool first before rerunning this tool...")            break    return# TESTING# def get_pro_project_params_test(layout_name):#     """This function returns a layer, layout, and map series object from the current ArcGIS Pro project."""#     active_map = r"C:\Users\SCDJ2L\dev\zoning_toolbox\OZmapMaker_v0\OZmapMaker_v20231214.aprx"#     aprx = arcpy.mp.ArcGISProject(active_map)#     map = aprx.listMaps("Premier Data Frame")[0]#     lyr = map.listLayers(".PageGrid_MASTERPageGrid")[0]#     layout = aprx.listLayouts(layout_name)[0]#     ms = layout.mapSeries##     return lyr, layout, ms, map# ## layer, layout, mapseries, mapx = get_pro_project_params_test("Section_MapSeries_OZmap")# set_seqcolor(mapx, layout=layout, map_series=mapseries)