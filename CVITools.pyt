# Project Name:     CVI Tools
# File Name:        CVITools.pyt
# Version:          0.1
# Author:           Jesse Langdon
# Last Update:      6/1/2023
# Description:      ArcGIS Pro Python toolbox with tools that facilitate updating data in the CVI Tool.
# Dependencies:     Python 3.x, arcpy, pandas
# ----------------------------------------------------------------------------------------------------------------------
# -*- coding: utf-8 -*-

# import modules
import os
import sys
import arcpy

# get current working directory of the toolbox
current_dir = arcpy.env.workspace

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the .pyt file)."""
        self.label = "CVI Tools"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [UpdateCVIExcel, UpdateCombinedCVI, UpdateAGOLFeatureLayers]


class UpdateCVIExcel(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "1 - Update Indicator data in the CVI Excel document"
        self.description = "Updates the SnohomishCountyCVI_Tool.xlsx Excel document with new indicator values. This " \
                           "tool only allows the user to overwrite an existing column."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        param0 = arcpy.Parameter(
            displayName="Snohomish County CVI Tool Excel spreadsheet",
            name="input_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        # param0.filter.list = ["xslx"]

        param1 = arcpy.Parameter(
            displayName="CSV file with updated indicator values",
            name="input_csv",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        param1.filter.list = ['txt', 'csv']

        param2 = arcpy.Parameter(
            displayName="Select the index to update",
            name="index_name",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        param2.filter.type = "ValueList"
        param2.filter.list = ["Adaptive Capacity Index",
                              "Sensitivity Index",
                              "Exposure Index"]

        param3 = arcpy.Parameter(
            displayName="Select a data source",
            name="data_source_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            enabled=False)

        param4 = arcpy.Parameter(
            displayName="Select the indicator to update",
            name="indicator_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            enabled=False)

        # Set the parameter dependencies
        param3.parameterDependencies = [param2.name]
        param4.parameterDependencies = [param3.name]

        params = [param0, param1, param2, param3, param4]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, params):
        """Modify the values and properties of parameters before internal
        validation is performed. This method is called whenever a parameter
        has been changed."""

        # import the data source config file
        data_src_dict = json_to_dict(os.path.join(sys.path[0], "csv_data_sources.json"))

        if params[2].altered:
            index = params[2].value
            data_src_list = get_data_src_by_index(data_src_dict, index)

            # populate data source input with list of data sources from config file, based on selected index
            if index:
                params[3].enabled = True
                params[3].filter.type = "ValueList"
                params[3].filter.list = data_src_list
            else:
                params[3].enabled = False
                params[3].value = None

        if params[3].altered:
            data_src = params[3].value
            indicator_list = get_indicators_by_data_src(data_src_dict, index, data_src)
            params[4].enabled = True
            params[4].filter.type = "ValueList"
            params[4].filter.list = indicator_list
        else:
            params[4].enabled = False
            params[4].value = None

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, params, messages):
        """The source code of the tool."""

        # Assign local variables
        spreadsheet_filename = params[0]
        csv_filename = params[1]
        index_name = params[2]
        data_source = params[3]
        indicator_name = params[4]

        update_CVI_excel(spreadsheet_filename, csv_filename, index_name, data_source, indicator_name)

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


class UpdateCombinedCVI(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "2 - Update Combined CVI Feature Class"
        self.description = "Updates the combined CVI feature class in the Snohomish_Climate.gdb file geodatabase."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        params = None
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""

        update_combined_CVI_fc()

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


class UpdateAGOLFeatureLayers(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "3 - Update CVI Hosted Feature Layers on AGOL"
        self.description = "Updates and overwrites hosted feature layers on AGOL for the CVI web map and application."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        params = None
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""

        update_AGOL_feature_layers()

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


### Helper functions

def update_CVI_excel(spreadsheet, csv, index, data_src, indicator):

    import pandas as pd
    import logging

    start_logging()

    # Convert SnohomishCountyCVI_Tool.xlsx sheet to pandas data frame(s)
    excel_df = pd.read_excel(spreadsheet, sheet_name=data_src)

    # Convert CSV file with summarized indicator data into pandas data frame
    csv_df = pd.read_csv(csv)

    # Merge data frames based on common column
    common_column = "Block Group ID"
    merged_df = pd.merge(excel_df, csv_df, on=common_column)

    # Replace indicator column in spreadsheet data frame with new updated data column
    merged_df[indicator] = merged_df[indicator].fillna(merged_df[indicator + "_y"])
    merged_df = merged_df.drop([indicator + "_y"], axis=1)

    # Write update spreadsheet data frame to Excel spreadsheet (replace existing sheet?)
    with pd.ExcelWriter(spreadsheet, engine='openpyxl', mode='a') as writer:
        merged_df.to_excel(writer, sheet_name=data_src, index=False)
        writer.save()

    success_msg = "The {0} indicator was successfully updated.".format(indicator)
    logging.info(success_msg)
    arcpy.AddMessage(success_msg)

    return


def json_to_dict(json_filepath):
    import json

    # parse JSON file as dictionary
    with open(json_filepath) as json_file:
        data_dict = json.load(json_file)
    arcpy.AddMessage("Data dictionary imported...")

    return data_dict


def get_data_src_by_index(dict_obj, index):

    # lists that will hold keys and values from the data dictionary
    first_level_key_list = []
    second_level_key_list = []

    for item in dict_obj:
        for key in item.keys():
            if key == index:
                first_level_key_list.append(key)
                for subitem in item[key]:
                    for subkey in subitem.keys():
                        second_level_key_list.append(subkey)
    return second_level_key_list


def get_indicators_by_data_src(dict_obj, index, data_src):
    for item in dict_obj:
        for key in item.keys():
            if key == index:
                for subitem in item[key]:
                    for subkey in subitem.keys():
                        if subkey == data_src:
                            indicator_list = subitem[subkey]
    return indicator_list


def start_logging():
    import logging
    import datetime

    # Create a datetime stamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Configure the logging settings
    log_filename = f"CVITools_{timestamp}.log"
    logging.basicConfig(filename='log_filename', level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    return

def update_combined_CVI_fc():

    #### Pseudo-code
    # 1. import libraries and data dictionary
    # 2. set arcpy env.workspace to CVI directory
    # 3. user input 1 - CVI file geodatabase (needed from user?)
    # 4. user input 2 - Select attribute fields to update in combined CVI feature class

    # 5. user input 2 - CVI Excel spreadsheet
    # 6. open log file
    # 7. based on user input 2, import Index sheet from Excel spreadsheet as pandas data frames
    # 8. convert Combined CVI feature class into data frame
    # 9. remove all columns from feature class data frame that will be replaced
    # 10. join updated columns to feature class data frame
    # 11. rearrange columns to match original schema
    # 12. convert updated feature class data frame to feature class (overwrite?)
    # 13. write to log file
    # 14. print message to user upon completion

    return

def update_AGOL_feature_layers():

    #### Pseudo-code
    # 1. import libraries (including arcgis library)
    # 3. user input 1 - file path to combined CVI feature class
    # 4. user input 2 - text to update combined CVI hosted feature layer description on AGOL
    # 2. set arcpy env.workspace to CVI directory
    # 5. based on SSO login info, return list of hosted feature layers in the CVI folder
    # 6. user chooses combined CVI hosted feature layer
    # 7. overwrite combined CVI hosted feature layer with local feature class
    # 8. print message to user upon completion

    return


# TESTING
spreadsheet_file = r"\\snoco\gis\plng\carto\CVI\SnohomishCounty_CVI\SnohomishCountyCVI_Tool.xlsx"
csv_file = r"\\snoco\gis\plng\carto\CVI\SnohomishCounty_CVI\test.csv"
index_name = "Exposure Index"
data_source = "BG_CIG_Exposure"
indicator_name = "Increase90DegreeDays"

update_CVI_excel(spreadsheet_file, csv_file, index_name, data_source, indicator_name)
print("Testing complete")