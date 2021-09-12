import csv
import openpyxl as xl
from pathlib import Path


def get_robot_plc_signals(sheet, signals_with_descriptions={}):
    '''
    Returns the extracted signals with descriptions from single xlsm sheet.


    Parameters:
        sheet (Worksheet): The single sheet from which signals and descriptions are to be extracted.
        signals_with_descriptions (dict): Signals with descriptions which are to be updated.


    Returns:
        signals_with_descriptions (dict): Signals and descriptions which get updated.
    '''

    start_plc_signals_row = 7
    end_plc_signals_row = 31

    signal_number_column = 11

    start_outputs_column = 12
    end_outputs_column = 15
    start_inputs_colum = 16
    end_inputs_column = 19

    for row in range(start_plc_signals_row, end_plc_signals_row):

        signal_number = sheet.cell(row, signal_number_column).value
        output_signal_description = ""
        input_signal_description = ""

        for column in range(start_outputs_column, end_outputs_column + 1):

            actual_cell_value = sheet.cell(row, column).value

            if actual_cell_value:
                output_signal_description += str(actual_cell_value) + " "
            else:
                break

        for column in range(start_inputs_colum, end_inputs_column + 1):

            actual_cell_value = sheet.cell(row, column).value

            if actual_cell_value:
                input_signal_description += str(actual_cell_value) + " "
            else:
                break

        if(output_signal_description != ""):

            signal_with_description = {
                "A"+str(signal_number): output_signal_description
            }

            signals_with_descriptions.update(signal_with_description)

        if(input_signal_description != ""):

            signal_with_description = {
                "E"+str(signal_number): input_signal_description
            }

            signals_with_descriptions.update(signal_with_description)

    return signals_with_descriptions
