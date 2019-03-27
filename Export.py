#!/usr/bin/env python3
from os import path, getlogin
from sys import argv
import configparser
import pandas as pd
from numpy import nan
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.formatting.rule import FormulaRule
import re
import PIL


def process_data(inp):
    """
    The main program. Takes file path, reads file, processes it and returns a finished dataframe.
    :param inp:
    :return: Dataframe
    """
    samples = read_file(inp)
    endo = read_endo_target(inp)
    samples = endo_cleanup(samples, endo)                # Remove endos
    ctrl_info = read_ctrl_targets(inp, samples)          # get targets for control. must be before ctrls removed
    samples, ctrls = separate_ctrls(samples)             # Separate controls from samples.
    samples['Assay Type'] = caller(samples, assay_type)  # set assay type
    samples = samples.sort_values(by=['Assay Type', 'Target', 'Sample'])  # Sort rows, must be before inserting formulas
    samples = add_formulas(samples, inp, ctrl_info)
    samples = pd.concat([samples, ctrls], sort=True)[cols_order]    # Insert ctrls to end of file, sort + remove columns
    samples['Cт'] = pd.to_numeric(samples['Cт'], errors='coerce')   # CT floats are str because of 'Undetermined'
    samples['Cт'] = samples['Cт'].fillna("Undetermined")            # ^ This causes ugly excel files. This fixes that.
    return samples


def read_file(inp):
    """
    Reads the file given as input and returns a dataframe. Only accepts columns in the first 9 of cols_list.
    :param inp:
    :return: Dataframe
    """
    params = [(14, "\t"), (15, "\t"), (14, ","), (15, ",")]  # list of parameters to try
    for header, sep in params:
        try:
            return pd.read_csv(inp, header=header, sep=sep, usecols=cols_order[:9])
        except ValueError:
            continue
    else:
        raise ValueError("That file doesnt look right.\nIt is either missing a column we need, or it is not an export" +
                         " file.\nCheck that you ticked 'Results' when exporting.\nColumns needed: " +
                         ", ".join(cols_order[:9]) + ".")


def read_assay_file():
    """
    Reads the assay file and returns a dataframe. The assay file path is specified in config.ini
    :return: dataframe
    """
    try:
        assays = pd.read_csv(config['File paths']['assays'], sep='\t')  # reads file.
    except FileNotFoundError:
        assays = input("Can't find the Assays file! If it has moved, please update config.ini with its new location."
                       "\nPress Enter to quit.")
        quit()
    # noinspection PyTypeChecker
    assays['Variant'] = assays['Variant'].str.lower()   # sets variant col to lower-case
    return assays.set_index('Variant', drop=False)  # set variant col as index, keeping variant as a column, returns.


def read_ctrl_targets(inp, samples):
    """Reads CSV to get control name, uses control name to get a list of targets that control applies to,
        Calling set() removes duplicates"""
    with open(inp, "r") as file:
        for i in range(12):
            next(file)
        ctrl_name = file.readline().split(sep='=')[1].strip().lower()
    return set(samples[samples.Sample.str.lower() == ctrl_name]['Target'].tolist()), ctrl_name


def read_endo_target(inp):
    with open(inp, "r") as file:
        for i in range(4):
            next(file)
        return file.readline().split(sep='=')[1].strip().lower()


def read_formulas():
    """
    Loads excel formulas used from a file on the team drive, calls get_formula_sub to format them and returns the
    finished formulas
    :return: String - formulas with {0} instead of row number.
    """
    try:
        wb = load_workbook(config['File paths']['Formulas'])
    except FileNotFoundError:
        wb = input("Can't find the Formulas file! If it has moved, please update config.ini with its new location."
                   "\nPress Enter to quit")
        quit()
    formula_sheet = wb['Sheet1']
    genf_sub = get_formula_sub(formula_sheet.cell(row=2, column=11).value)
    assayf_sub = get_formula_sub(formula_sheet.cell(row=2, column=17).value)
    confirmf_sub = get_formula_sub(formula_sheet.cell(row=2, column=18).value)
    return genf_sub, assayf_sub, confirmf_sub


def get_formula_sub(formula):
    """
    Takes in a formula, splits it and replaces the row number with a sub string {0} for later use.
    :param formula:
    :return: String - formula with {0} instead of row number.
    """
    from openpyxl.formula import Tokenizer
    formula_sub = "="
    for t in Tokenizer(formula).items:          # Tokenizer is part of openpyxl
        if t.subtype is 'RANGE':                # if token is a cell reference
            t.value = t.value[0] + "{0}"        # replace row number with {0} e.g F3 > F{0}
            formula_sub += t.value              # add it to formula_sub
        else:
            formula_sub += t.value              # else do nothing and add it to formula_sub
    return formula_sub


def endo_cleanup(samples, endo):
    """
    Removes endogenous control results. Adds a field to samples that indicates if the corresponding endo was omitted.
    :param samples:     Dataframe
    :param endo:        String - name of the endogenous control
    :return:            Dataframe
    """
    endos = samples[samples.Target.str.lower() == endo][['Well ', 'Omitted ']]  # get df of well + omitted for endo
    endos = endos.rename(index=str, columns={'Omitted ': 'Omitted_endo'})  # rename endos omitted column
    samples = samples[samples.Target.str.lower() != endo]  # Remove endo from samples
    samples = pd.merge(samples, endos, on='Well ', how='inner')  # Merge endos with samples using Well as the key
    return samples


def separate_ctrls(samples):
    """regular expression to pattern match a mouse or blast in the form ~PMGB11.2a or M02983000
       (letter a-z 3-4 times, number 1-3 times, full stop, number 1-2 times, letter a-z) OR (letter m, number x 8) """
    regex = '[a-z]{3,4}\\d{1,3}\\.\\d{1,2}[a-z]|m\\d{8}'

    ctrls = samples.loc[~samples['Sample'].str.lower().str.contains(regex)]  # Move controls to new dataframe
    samples = samples.loc[samples['Sample'].str.lower().str.contains(regex)]  # Remove controls from df
    ctrls = ctrls.sort_values(by=['Target', 'Sample'])
    return samples, ctrls


def assay_type(line):
    """
    Determines if the assay type is LoA or qPCR. Add new assays to the odd_assays list in all caps.
    :param line: pd.Series - Horizontal row from dataframe
    :return: String
    """
    target = line['Target'].lower()
    if target in assay_frame['Variant']:
        return assay_frame.loc[target, 'Type']
    elif "_wt" in target or "_ce" in target:
        return "LoA"
    else:
        return "Unknown"


def add_formulas(samples, inp, ctrl_info):
    """ Adds formulas and extra columns. Row number is added by string formatting based on df['index']
        Adding formulas rather than doing the logic in python allows the user to make adjustments."""

    samples['index'] = range(2, samples.shape[0] + 2)  # Create an index that corresponds to excel row number
    barcode = path.basename(inp).split("_")[0].upper()
    ctrl_targets, ctrl_name = ctrl_info
    samples['Assay Name'] = caller(samples, assay_name)  # must be before is_transgene is called.
    columns_add = {'Mouse':      samples['Sample'], 'Plate Barcode': barcode,
                   'Allele':     nan, 'Locked': nan,  'Comment':       nan, 'Name': nan,
                   'Compare':    nan, 'Gender': nan,
                   'Het Control?':  caller(samples, het_flag, ctrl_name, ctrl_targets),
                   'X-Linked?':  caller(samples, is_transgene),
                   'RQ   ':      caller(samples, rq_add_zero, ctrl_targets),
                   'Genotype':   caller(samples, genf),
                   'Result':     caller(samples, assayf),
                   'Confirmed':  caller(samples, confirmf),
                   }  # dict of columns we need to add and their values
    for col_name in columns_add:  # Add the columns in columns_add, and set its value respectively.
        samples[col_name] = columns_add[col_name]
    return samples


def caller(samples, form_or_func, *args):
    """
    Applies a function or formula to the dataframe.
    Formula uses .format to insert row number.
    :param samples: Dataframe
    :param form_or_func: formula or function
    :param args: variable number of arguments for function
    :return:
    """
    try:
        return samples.apply(lambda line: pd.Series([form_or_func.format(line['index'])]), axis=1)  # formula
    except AttributeError:
        return samples.apply(lambda line: pd.Series([form_or_func(line, *args)]), axis=1)           # function


def assay_name(line):
    """
    Checks Target against a list of common mistakes and corrects them.
    :param line: pd.Series - Horizontal row from dataframe
    :return: String
    """
    target = line['Target'].lower()
    if line['Target'].lower() in assay_frame['Variant']:
        return assay_frame.loc[target, 'Assay']
    else:
        return line['Target']


def rq_add_zero(line, ctrl_targets):
    """Adds 0 to RQ if the control applies to that target. Does not add 0 if endo has been omitted."""

    if line['Omitted_endo']:    # TODO remove this and move to formula?
        return nan                                                  # If endo is omitted, insert NaN.
    if line['Cт'] == "Undetermined":                                # If CT is Undetermined
        if line['Target'] in ctrl_targets:
            return 0                                                # Add 0 to RQ if target also applies to the control
        else:
            return nan                                              # If not add NaN
    else:
        return line['RQ   ']                                        # Keep same RQ value.


def het_flag(line, ctrl_name, ctrl_targets):
    """
    If the control in export is a het, and applies to that target, return "Yes"
    :param line:            A line of the dataframe
    :param ctrl_name:       String - the control
    :param ctrl_targets:    Set of targets the control applies to.
    :return:                String - "Yes" if cond is met
    """
    if "het" in ctrl_name and line['Target'] in ctrl_targets:
            return "Yes"


def is_transgene(line):
    """
    If the assay is a transgene assay, returns Transgene.
    :param line: pd.Series - Horizontal row from dataframe
    :return: String.
    """
    target = line['Assay Name'].upper()
    if "_TG" in target:
        return "Transgene"


def get_sheet_name(plate):  # todo: improve this.
    """
    Parses the file name and shortens it to <32 chars so it can be used as the sheet name in excel.
    :param plate:
    :return:
    """
    plate_barcode = r'^c0000\d{5}' + r'|^sl000\d{5}' + r'|^\d{5}'  # allow misspellings???
    rex = re.compile(plate_barcode, re.IGNORECASE)  # todo needed or not?
    plate2 = plate.split(sep='_')  # separate on underscores
    try:
        plate2.remove('data')
    except ValueError:
        pass
    users = {'jb40', 'db11', 'es16', 'sa24', 'dg4', 'er1', 'db7', getlogin()}  # todo: read this in? better way?
    # I cant think of a way to regex user log in names without getting gene names that also fit that pattern
    user = []
    plates = []
    assays_etc = []
    plates_small = []

    for i in plate2:
        if i in users:
            user.append(i)
            continue
        p = rex.match(i)
        if p:
            plates.append(p.group())
        else:
            assays_etc.append(i)
    for i in plates:
        f = i.replace('SL000', '')
        s = f.replace('C0000', '')[:5]
        plates_small.append(s)

    final = '_'.join(plates + assays_etc + user)
    if len(final) >= 31:
        final = '_'.join(plates_small + assays_etc + user)
    if len(final) >= 31:
        final = '_'.join(plates_small[0:1] + assays_etc + user)
    if len(final) >= 31:
        final = '_'.join(plates_small[0:1] + assays_etc)
    if len(final) >= 31:
        final = final[:31]

    return final


def to_xlsx(out, dataframe, main_file=''):
    """
    Exports to xlsx file and formats it with correct column widths, top row freeze pane and conditional formatting
    based on genotype.
    :param out: Input file path (.txt)
    :param dataframe:
    :param main_file: path to main xlsx file
    :return: Output file path
    """

    sheet = get_sheet_name(path.split(path.splitext(out)[0])[1])
    xlsx_file = path.splitext(out)[0] + '.xlsx'
    if path.isfile(xlsx_file):
        main_file = xlsx_file
    if main_file:
        book = load_workbook(main_file)
        writer = pd.ExcelWriter(main_file, engine='openpyxl')
        writer.book = book
        if sheet in book.sheetnames:
            for i in range(1, 100):
                sheet1 = sheet + str(i)
                if len(sheet1) > 31:
                    sheet1 = sheet[:30] + str(i)
                if len(sheet1) > 31:
                    sheet1 = sheet[:29] + str(i)
                if sheet1 not in book.sheetnames:
                    sheet = sheet1
                    break
    else:
        out = xlsx_file
        writer = pd.ExcelWriter(path.splitext(out)[0] + '.xlsx', engine='openpyxl')
    dataframe.to_excel(writer, sheet_name=sheet, index=False, freeze_panes=(1, 0))  # Write dataframe to excel
    """Formatting for a pretty output"""
    wb = writer.book
    ws = wb[sheet]

    col_width = {'A': 5, 'C': 11, 'J': 11, 'K': 9.14, 'N': 12.57, 'O': 10, 'P': 18, 'R': 10, 'S': 15.14, 'W': 11.43,
                 'Y': 13.43}

    for col in col_width:  # todo Split this out to allow some non centered col width adjustments? or dict slicing?
        ws.column_dimensions[col].width = col_width[col]
        for cell in ws[col]:
            cell.alignment = Alignment(horizontal='center')

    conditions = {'Het': PatternFill(patternType='solid', bgColor='DCE6F0'),
                  'Hom': PatternFill(patternType='solid', bgColor='B8CCE4'),
                  'Hemi': PatternFill(patternType='lightUp', bgColor='DCE6F0', fgColor='B8CCE4'),
                  'Fail': PatternFill(patternType='solid', bgColor='FFC7CE'),
                  'Retest': PatternFill(patternType='solid', bgColor='FFC7CE')}
    for genotype in conditions:
        ws.conditional_formatting.add('K2:K' + str(dataframe.shape[0] + 2),
                                      FormulaRule(formula=['NOT(ISERROR(SEARCH("' + genotype + '",K2)))'],
                                                  stopIfTrue=True, fill=conditions[genotype]))
    wb.active = ws
    writer.save()  # Save xlsx.
    if main_file:
        print('Sheet added: ' + sheet)
    return main_file if main_file else out


def auto_to_xl(file, main_file=''):
    try:
        dfr = process_data(file)  # Formats data and inserts formulas.
        out = to_xlsx(file, dfr, main_file=main_file)
        if not main_file:
            print("Export processing complete: " + path.split(out)[1])
        return out
    except PermissionError as e:
        print(e)
        print("You already have an export of this file open. Close it and re-try.")
    except ValueError as e:  # if file is missing cols or is not an export - raised in read_file()
        print(e)


cols_order = ['Well ', 'Omitted ', 'Sample', 'Target', 'Reporter', 'RQ   ', 'Cт', 'ΔCт', 'ΔΔCт', 'Mouse',
              'Genotype', 'Allele', 'Locked', 'Plate Barcode', 'Assay Type', 'Assay Name', 'Result',
              'Confirmed', 'Comment', 'Name', 'Compare', 'Gender', 'Het Control?', 'X-Linked?', 'Omitted_endo']

config = configparser.ConfigParser()
config.read(path.dirname(argv[0]) + '/config.ini')           # read config.ini from same folder the script is in.
assay_frame = read_assay_file()                              # Reads Assay info from file
genf, assayf, confirmf = read_formulas()                     # Reads Formulas from file
print("Assays and formulas loaded.")
