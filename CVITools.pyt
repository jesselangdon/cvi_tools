# Project Name:     CVI Tools
# File Name:        CVITools.pyt
# Version:          0.1
# Author:           Jesse Langdon
# Last Update:      9/21/2023
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
        self.tools = [UpdateCVI, UpdateAGOLFeatureLayers]


class UpdateCVI(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "1 - Update indicator in CVI Excel spreadsheet and feature class"
        self.description = "Updates the SnohomishCountyCVI_Tool.xlsx Excel document with new indicator values, as" \
                           "well as the SnohomishCounty_BG_Index_Final feature class."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        param0 = arcpy.Parameter(
            displayName="CVI Tool Excel spreadsheet",
            name="input_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        param0.filter.list = ["xslx"]

        param1 = arcpy.Parmeter(
            displayName="CVI feature class",
            name="input_fc",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Input")

        param2 = arcpy.Parameter(
            displayName="CSV file with updated indicator values",
            name="input_csv",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        param2.filter.list = ['txt', 'csv']

        param3 = arcpy.Parameter(
            displayName="Select the index to update",
            name="index_name",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        param3.filter.type = "ValueList"
        param3.filter.list = ["Adaptive Capacity Index",
                              "Sensitivity Index",
                              "Exposure Index"]

        param4 = arcpy.Parameter(
            displayName="Select a data source",
            name="data_source_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            enabled=False)

        param5 = arcpy.Parameter(
            displayName="Select the indicator to update",
            name="indicator_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            enabled=False)

        # Set the parameter dependencies
        param4.parameterDependencies = [param3.name]
        param5.parameterDependencies = [param4.name]

        params = [param0, param1, param2, param3, param4, param5]
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

        if params[3].altered:
            index = params[3].valueAsText
            data_src_list = get_data_src_by_index(data_src_dict, index)

            # populate data source input with list of data sources from config file, based on selected index
            if index:
                params[4].enabled = True
                params[4].filter.type = "ValueList"
                params[4].filter.list = data_src_list
            else:
                params[4].enabled = False
                params[4].value = None

        if params[4].altered:
            data_src = params[4].valueAsText
            indicator_list = get_indicators_by_data_src(data_src_dict, index, data_src)
            params[5].enabled = True
            params[5].filter.type = "ValueList"
            params[5].filter.list = indicator_list
        else:
            params[5].enabled = False
            params[5].value = None

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, params, messages):
        """The source code of the tool."""

        # Assign local variables
        spreadsheet_filename = params[0].valueAsText
        fc_name = params[1].valueAsText
        csv_filename = params[2].valueAsText
        subindex_name = params[3].valueAsText
        data_source = params[4].valueAsText
        indicator_name = params[5].valueAsText
        unique_id = "Block Group ID"

        update_CVI_excel(spreadsheet_filename, csv_filename, data_source, indicator_name, unique_id)

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

# TODO This helper function will be replaced with the update_cvi function.
def update_CVI_excel(spreadsheet, csv, data_src, indicator, unique_id):
    start_logging()
    logging.info("Running UpdateCVIExcel tool...")

    try:
        today = datetime.date.today().strftime("%Y%m%d")
        spreadsheet_copy = spreadsheet.replace('.xlsx', f'_{today}.xlsx')
        shutil.copy(spreadsheet, spreadsheet_copy)

        df_excel, df_csv = convert_to_dataframe(spreadsheet_copy, csv, data_src, unique_id)
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


# TODO this function is replacing the update_CVI_excel function
def update_CVI(speadsheet, fc, csv, data_src, indicator, uid):
    start_logging()
    logging.info("Running the Update CVI tool...")
    try:
        # make a copy of the CVI spreadsheet to update, appending today's date to the new file name
        spreadsheet_to_update = make_spreadsheet_copy(spreadsheet)

        # convert data files (spreadsheet, csv files) to pandas data frames
        df_cvi_index = pd.read_excel(spreadsheet_to_update, sheet_name="CVI_Index")
        df_indicator = pd.read_excel(spreadsheet_to_update, sheet_name=indicator)
        df_datasrc = pd.read_excel(spreadsheet_to_update, sheet_name=data_src)
        df_csv = pd.read_excel(csv)

        # check for duplicate values in the unique ID column for each data frame
        df_list = [df_cvi_index, df_indicator, df_datasrc, df_csv]
        for df in df_list:
            check_df_column_for_dupes(df, uid)

        # update the indicator columns of the target dataframes based on the new values in the CSV indicator column




    except Exception as e:
        traceback_str = traceback.format_exc()
        error_msg = f"An exception occurred: \n{traceback_str}"
        logging.error(error_msg)
        arcpy.AddError(error_msg)

    return


def make_spreadsheet_copy(spreadsheet):
    today = datetime.date.today().strftime("%Y%m%d")
    spreadsheet_copy = spreadsheet.replace('.xlsx', f'_{today}.xlsx')
    shutil.copy(spreadsheet, spreadsheet_copy)
    return spreadsheet_copy


def check_df_column_for_dupes(df, col_name):
    if df[col_name].duplicated().sum() != 0:
        raise ValueError(f"Duplicate values found in {col_name} column!")
    return


def json_to_dict(json_filepath):
    with open(json_filepath) as json_file:
        data_dict = json.load(json_file)
    arcpy.AddMessage("Data dictionary imported...")
    return data_dict


def get_data_src_by_index(dict_obj, index):
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
                            return subitem[subkey]


def start_logging():
    tool_dir = os.path.dirname(os.path.abspath(__file__))
    log_filename = os.path.join(tool_dir, "CVTTools.log")
    logging.basicConfig(
        filename=log_filename, filemode='a', level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.info("---------------- New Execution ----------------")
    return


# def convert_to_dataframe(excel_path, csv_path, sheet_name, unique_id):
#     df_excel = pd.read_excel(excel_path, sheet_name=sheet_name)
#     if df_excel[unique_id].duplicated().sum() != 0:
#         raise ValueError(f"Duplicates found in Excel {unique_id} column")
#
#     df_csv = pd.read_csv(csv_path)
#     if df_csv[unique_id].duplicated().sum() != 0:
#         raise ValueError(f"Duplicates found in CSV {unique_id} column")
#
#     return df_excel, df_csv


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


def update_combined_CVI_fc(input_fc, spreadsheet, unique_id):
    # TODO This function will be removed. Functionality to update the CVI feature class will be incorporated into a
    # TODO single tool in the toolbox, since all of the user inputs will already be available.

    # Determine file geodatabase path from the feature class path
    workspace = os.path.dirname(input_fc)
    arcpy.env.workspace = workspace

    start_logging()
    logging.info("Running UpdateCombinedCVI tool...")

    # Convert Combined CVI feature class into data frame
    df_cvi = pd.read_excel(io=spreadsheet, sheet_name="CVI_Index", index_col=unique_id)

    # Get list of column headers from CVI_Index spreadsheet
    fields_to_delete = df_cvi.columns.values.tolist()
    fields_to_delete = fields_to_delete.remove(unique_id)

    # Remove fields from Climate Vulnerability feature class that will be replaced
    delete_fields_from_fc(input_fc, fields_to_delete)

    # Convert the combined CVI data frame into a temporary table in the CVI file geodatabase
    temp_table = "excel_df_temp"
    arcpy.TableToTable_conversion(in_rows=df_cvi, out_path=workspace, out_name=temp_table)

    # Join Excel data table with the feature class
    arcpy.JoinField_management(in_data=input_fc, in_field=unique_id, join_table=temp_table, join_field=unique_id)
    arcpy.AddMessage("Attribute fields were successfully replaced in the CVI feature class!")
    return


def delete_fields_from_fc(input_fc, fields_to_del):
    '''Delete all columns from feature class data frame that will be replaced, based on columns in list_col_headers'''
    fc_fields = arcpy.ListFields(input_fc)
    for fc_field in fc_fields:
        if fc_field.name in fields_to_del:
            arcpy.DeleteField_management(input_table=input_fc, Fields=fc_field.name)
            arcpy.AddMessage(f"Attribute field {fc_field.name} removed...")
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
input_fc = r"\\snoco\gis\plng\carto\CVI\SnohomishCounty_CVI\GIS\Snohomish_Climate.gdb\SnohomishCounty_BG_Index_Final"
spreadsheet = r"C:\Users\SCDJ2L\dev\CVI\TEST\SnohomishCountyCVI_Tool_20231213.xlsx"
unique_id = "Block Group ID"
update_combined_CVI_fc(input_fc, spreadsheet, unique_id)