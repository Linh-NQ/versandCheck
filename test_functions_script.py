import xlwings as xw
import pandas as pd
import glob
import tempfile
import os
import pytest
from Versand_Check import frame
from Versand_Check import return_error_rows_as_string, zellen_bunt_malen, check_feldcode, check_masterid, check_sampleid, check_patid, check_sample_master, check_pflichtfelder, check_discharge_reason, check_datum, check_remarks, check_condition, check_datenreihe, check_versandgrund


# Unit Test für return_error_rows_as_string
def test_return_error_rows_as_string():
    assert return_error_rows_as_string([1]) == '1'
    assert return_error_rows_as_string([1, 3]) == '1, 3'
    assert return_error_rows_as_string([1, 2, 3]) == '1-3'
    assert return_error_rows_as_string([1, 2, 3, 5, 6, 7]) == '1-3, 5-7'
    assert return_error_rows_as_string([1, 2, 3, 5, 6, 7, 9]) == '1-3, 5-7, 9'
    assert return_error_rows_as_string([1, 2, 3, 5, 7, 8]) == '1-3, 5, 7-8'
    assert return_error_rows_as_string([101, 103, 104, 105, 110]) == '101, 103-105, 110'
    

# Helper function to check cell color
def check_cell_color(ws, column_name, row, expected_color):
    global file_name
    column_index = ord(column_name) - ord('A') + 1  # Convert column letter to index
    actual_color = ws.cells(row, column_index).color
    return actual_color == expected_color, f"Cell color mismatch at {column_name}{row}. Expected: {expected_color}, Actual: {actual_color}"

# Unit Test für zellen_bunt_malen
def test_zellen_bunt_malen():
    global file_name
    vorlage_path = 'excel_unit_tests/test1.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']

    # Rot
    zellen_bunt_malen('1', 'master_id', vorlage, ws, (220, 20, 60))
    assert check_cell_color(ws, 'A', 1, (220, 20, 60))

    # Grau
    zellen_bunt_malen('2, 4', 'pat_id', vorlage, ws, (168, 168, 168))
    assert check_cell_color(ws, 'B', 2, (168, 168, 168))
    assert check_cell_color(ws, 'B', 4, (168, 168, 168))

    # Gelb
    zellen_bunt_malen('1-3, 5, 7-8', 'sample_id', vorlage, ws, (238, 232, 170))
    assert check_cell_color(ws, 'C', 1, (238, 232, 170))
    assert check_cell_color(ws, 'C', 2, (238, 232, 170))
    assert check_cell_color(ws, 'C', 3, (238, 232, 170))
    assert check_cell_color(ws, 'C', 5, (238, 232, 170))
    assert check_cell_color(ws, 'C', 7, (238, 232, 170))
    assert check_cell_color(ws, 'C', 8, (238, 232, 170))

    wb.close()
    

# Unit Test für check_feldcode
def test_check_feldcode():
    global felder_f
    vorlage_path = r'excel_unit_tests/test_feldcode.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_feldcode(vorlage, cols, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'K', 1, (220, 20, 60))
    assert check_cell_color(ws, 'M', 1, (220, 20, 60))
    assert check_cell_color(ws, 'N', 1, (220, 20, 60))
    assert check_cell_color(ws, 'O', 1, (220, 20, 60))

    wb.close()



# Unit Test für check_masterid
def test_check_masterid():
    vorlage_path = r'excel_unit_tests/test_masterid.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_masterid(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 4, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))
    assert check_cell_color(ws, 'A', 13, (220, 20, 60))
    assert check_cell_color(ws, 'A', 15, (220, 20, 60))

    wb.close()



# Unit Test für check_sampleid
def test_check_sampleid():
    vorlage_path = r'excel_unit_tests/test_sampleid.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_sampleid(cols, vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 2, (220, 20, 60))
    assert check_cell_color(ws, 'A', 4, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))
    assert check_cell_color(ws, 'A', 10, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'A', 12, (220, 20, 60))

    wb.close()



# Unit Test für check_patid
def test_check_patid():
    vorlage_path = r'excel_unit_tests/test_patid.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)
    # Einlesen der Excel-Tabelle 'Aufbau und Felder'

    error_count = [0]
    error_count_total = [0]

    check_patid(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 5, (220, 20, 60))
    assert check_cell_color(ws, 'A', 6, (220, 20, 60))
    assert check_cell_color(ws, 'A', 7, (220, 20, 60))

    wb.close()



# Unit Test für check_sample_master
def test_check_sample_master():
    vorlage_path = r'excel_unit_tests/test_sample_master.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_sample_master(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'A', 8, (220, 20, 60))
    assert check_cell_color(ws, 'A', 11, (220, 20, 60))
    assert check_cell_color(ws, 'B', 4, (220, 20, 60))
    assert check_cell_color(ws, 'B', 5, (220, 20, 60))
    assert check_cell_color(ws, 'B', 7, (220, 20, 60))
    assert check_cell_color(ws, 'B', 9, (220, 20, 60))

    wb.close()



# Units Test für check_pflichtfelder
def test_check_pflichtfelder():
    vorlage_path = r'excel_unit_tests/test_pflichtfelder.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_pflichtfelder(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'A', 3, (220, 20, 60))
    assert check_cell_color(ws, 'B', 4, (220, 20, 60))
    assert check_cell_color(ws, 'C', 5, (220, 20, 60))
    assert check_cell_color(ws, 'D', 6, (220, 20, 60))
    assert check_cell_color(ws, 'G', 6, (220, 20, 60))
    assert check_cell_color(ws, 'P', 8, (220, 20, 60))
    assert check_cell_color(ws, 'P', 9, (220, 20, 60))
    assert check_cell_color(ws, 'P', 10, (220, 20, 60))
    assert check_cell_color(ws, 'T', 8, (220, 20, 60))
    assert check_cell_color(ws, 'T', 9, (220, 20, 60))
    assert check_cell_color(ws, 'T', 10, (220, 20, 60))
    assert check_cell_color(ws, 'R', 8, (220, 20, 60))
    assert check_cell_color(ws, 'S', 8, (220, 20, 60))
    assert check_cell_color(ws, 'V', 8, (220, 20, 60))
    assert check_cell_color(ws, 'V', 9, (220, 20, 60))
    assert check_cell_color(ws, 'V', 10, (220, 20, 60))

    wb.close()


# Unit Test für check_discharge_reason
def test_check_discharge_reason():
    vorlage_path = r'excel_unit_tests/test_discharge_reason.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_discharge_reason(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'D', 1, (220, 20, 60))
    assert check_cell_color(ws, 'E', 10, (220, 20, 60))
    assert check_cell_color(ws, 'F', 9, (220, 20, 60))
    assert check_cell_color(ws, 'G', 3, (220, 20, 60))
    assert check_cell_color(ws, 'G', 4, (220, 20, 60))
    assert check_cell_color(ws, 'G', 5, (220, 20, 60))
    assert check_cell_color(ws, 'G', 6, (220, 20, 60))
    assert check_cell_color(ws, 'G', 7, (220, 20, 60))
    assert check_cell_color(ws, 'J', 8, (220, 20, 60))
    assert check_cell_color(ws, 'Q', 2, (220, 20, 60))

    wb.close()


# Unit Test für check_datum
def test_check_datum():
    vorlage_path = r'excel_unit_tests/test_datum.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_datum(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'T', 1, (220, 20, 60))
    assert check_cell_color(ws, 'T', 2, (220, 20, 60))
    assert check_cell_color(ws, 'T', 3, (220, 20, 60))
    assert check_cell_color(ws, 'T', 4, (220, 20, 60))
    assert check_cell_color(ws, 'T', 5, (220, 20, 60))
    assert check_cell_color(ws, 'W', 6, (220, 20, 60))

    wb.close()



# Unit Test für check_remarks
def test_check_remarks():
    vorlage_path = r'excel_unit_tests/test_remarks.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_remarks(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'U', 7, (220, 20, 60))

    wb.close()



# Unit Test für check_condition
def test_check_condition():
    vorlage_path = r'excel_unit_tests/test_condition.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_condition(vorlage, 'condition', ws, error_count, error_count_total)
    check_condition(vorlage, 'aliquoteTypeId', ws, error_count, error_count_total)
    assert check_cell_color(ws, 'E', 8, (220, 20, 60))
    assert check_cell_color(ws, 'F', 7, (220, 20, 60))

    wb.close()



# Unit Test für check_datenreihe
def test_check_datenreihe():
    vorlage_path = r'excel_unit_tests/test_datenreihe.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_datenreihe(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'R', 3, (168, 168, 168))
    assert check_cell_color(ws, 'R', 4, (168, 168, 168))
    assert check_cell_color(ws, 'R', 2, (168, 168, 168))
    assert check_cell_color(ws, 'V', 6, (168, 168, 168))

    wb.close()



# Unit Test für check_versandgrund
def test_check_versandgrund():
    vorlage_path = r'excel_unit_tests/test_datenreihe.xlsx'
    vorlage = pd.read_excel(vorlage_path)
    file_name = vorlage_path.split('/')[-1].split('.xlsx')[0]
    vorlage.to_excel('{}_Check.xlsx'.format(file_name), index=False)
    wb = xw.Book('{}_Check.xlsx'.format(file_name))
    ws = wb.sheets['Sheet1']
    cols = list(vorlage.columns)

    error_count = [0]
    error_count_total = [0]

    check_versandgrund(vorlage, ws, error_count, error_count_total)
    assert check_cell_color(ws, 'R', 5, (220, 20, 60))
    assert check_cell_color(ws, 'R', 8, (220, 20, 60))
    assert check_cell_color(ws, 'R', 11, (220, 20, 60))
    assert check_cell_color(ws, 'R', 12, (220, 20, 60))
    assert check_cell_color(ws, 'S', 7, (220, 20, 60))
    assert check_cell_color(ws, 'S', 8, (220, 20, 60))
    assert check_cell_color(ws, 'S', 12, (220, 20, 60))
    assert check_cell_color(ws, 'V', 5, (220, 20, 60))
    assert check_cell_color(ws, 'V', 7, (220, 20, 60))
    assert check_cell_color(ws, 'V', 8, (220, 20, 60))
    assert check_cell_color(ws, 'V', 10, (220, 20, 60))
    assert check_cell_color(ws, 'V', 11, (220, 20, 60))
    assert check_cell_color(ws, 'V', 12, (220, 20, 60))

    wb.close()