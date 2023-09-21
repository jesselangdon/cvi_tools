# Project Name:     CVI Tools
# File Name:        CVITools.pyt
# Version:          0.1
# Author:           Jesse Langdon
# Last Update:      6/1/2023
# Description:      ArcGIS Pro Python toolbox with tools that facilitate updating data in the CVI Tool.
# Dependencies:     Python 3.x, arcpy, pandas, xlwings (note: the xlwings package may need to be installed manually)
# ----------------------------------------------------------------------------------------------------------------------
# -*- coding: utf-8 -*-

# import modules
import os
import sys
import json
import pandas as pd
import xlwings as xw
import shutil
import datetime
import logging
import arcpy
import traceback

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
        self.label = "1 - Update indicator in CVI Excel spreadsheet"
        self.description = "Updates the SnohomishCountyCVI_Tool.xlsx Excel document with new indicator values. This " \
                           "tool only allows the user to overwrite an existing indicator column."
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
            index = params[2].valueAsText
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
            data_src = params[3].valueAsText
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
        spreadsheet_filename = params[0].valueAsText
        csv_filename = params[1].valueAsText
        data_source = params[3].valueAsText
        indicator_name = params[4].valueAsText
        unique_id = "Block Group ID"

        update_CVI_excel(spreadsheet_filename, csv_filename, data_source, indicator_name, unique_id)

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

def update_CVI_excel(spreadsheet, csv, data_src, indicator, unique_id):
    start_logging()
    logging.info("Running UpdateCVIExcel tool...")

    try:
        today = datetime.date.today().strftime("%Y%m%d")
        spreadsheet_copy = spreadsheet.replace('.xlsx', f'_{today}.xlsx')
        shutil.copy(spreadsheet, spreadsheet_copy)

        df_excel, df_csv = read_data_files(spreadsheet_copy, csv, data_src, unique_id)
        df_excel_cols = list(df_excel.columns)
        df_updated = update_dataframe(df_excel, df_csv, unique_id)

        df_updated = df_updated[df_excel_cols]
        write_to_excel(df_updated, spreadsheet_copy, data_src)

        logging.info(f"The {indicator} indicator was successfully updated.")

    except Exception as e:
        traceback_str = traceback.format_exc()
        error_msg = f"An exception occurred: \n{traceback_str}"
        logging.error(error_msg)
        arcpy.AddError(error_msg)

    return


def json_to_dict(json_filepath):
    with open(json_filepath) as json_file:
        data_dict = json.load(json_file)
    arcpy.AddMessage("Data dictionary imported...")
    return data_dict


def get_data_src_by_index(dict_obj, index):
    second_level_key_list = [subitem[subkey] for item in dict_obj for key in item.keys() if key == index for subitem in item[key] for subkey in subitem.keys()]
    return second_level_key_list


def get_indicators_by_data_src(dict_obj, index, data_src):
    for item in dict_obj:
        for key in item.keys():
            if key == index:
                for subitem in item[key]:
                    for subkey in subitem.keys():
                        if subkey == data_src:
                            return subitem[subkey]


def start_logging():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"CVITools_{timestamp}.log"
    logging.basicConfig(
        filename=log_filename, filemode='w', level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return


def read_data_files(excel_path, csv_path, sheet_name, unique_id):
    df_excel = pd.read_excel(excel_path, sheet_name=sheet_name)
    if df_excel[unique_id].duplicated().sum() != 0:
        raise ValueError(f"Duplicates found in Excel {unique_id} column")

    df_csv = pd.read_csv(csv_path)
    if df_csv[unique_id].duplicated().sum() != 0:
        raise ValueError(f"Duplicates found in CSV {unique_id} column")

    return df_excel, df_csv


def update_dataframe(df_excel, df_csv, unique_id):
    df_excel.set_index(unique_id, inplace=True)
    df_csv.set_index(unique_id, inplace=True)
    df_excel.update(df_csv, join="left", overwrite=True)
    df_excel.reset_index(inplace=True)
    return df_excel


def write_to_excel(df, excel_path, sheet_name):
    app = xw.App(visible=False)
    wb = app.books.open(excel_path)
    target_sheet = wb.sheets[sheet_name]
    target_sheet.range('A2').options(index=False, header=False).value = df
    wb.save()
    wb.close()
    app.quit()
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
spreadsheet_file = r"C:\Users\SCDJ2L\dev\CVI\TEST\SnohomishCountyCVI_Tool.xlsx"
csv_file = r"C:\Users\SCDJ2L\dev\CVI\TEST\slr_parcels_20230920.csv"
index_name = "Exposure Index"
data_source = "BG_CIG_Exposure"
indicator_name = "SeaLevelRise_2Ft_Parcels"
unique_id = "Block Group ID"

update_CVI_excel(spreadsheet_file, csv_file, data_source, indicator_name, unique_id)