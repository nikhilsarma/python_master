'''
Id:          "$Id: zn_differ.py,v 1.9 2013/03/13 15:43:23 jerson.chua Exp $"
Description: 
Test:     
Category: Util
'''
from qz.data.qztable_csv import load
from qz.core import bobfns
from qz.data.qztable_utils import cmp
from risk.zinc.util.fileutil import saveQZTableAsCSV

import qztable

DEFAULT_DIFF_LIMIT=10000

def db_friendly_column_name(column_name):
    return column_name.strip().replace('[', '').replace(']', '').replace('.', '')
    
def fileToQztable(filename, delimiter='\t', skip_header=0):
    with open(filename) as file1:
        for i in range(skip_header):
            file1.readline()
        
        # load() assumes the first line in the file is the header. That's why we need to extract the headers here.
        fields = [db_friendly_column_name(e) for e in file1.readline().split(delimiter)]
        
        # skip_footer isn't working as described in http://stackoverflow.com/questions/3761103/using-genfromtxt-to-import-csv-data-with-missing-values-in-numpy
        # using invalid_raise=False to work around the issue
        return load(filename, readFields=False, fields=fields, delimiter=delimiter, skip_header=skip_header+1, invalid_raise=False)


def generate_diff_row(keys, row, column_idx, field_name, value1, value2):
    line = []

    for key in keys:
        key_val = row[column_idx[key]]
        line.append(key_val)
        
    line.append(field_name)
    line.append(value1)
    line.append(value2)
    
    return line
    
def equals(val1, val2, epsilon):
    try:
        fval1 = float(val1)
        fval2 = float(val2)
        return abs(fval1 - fval2) <= epsilon
    except:
        return val1 == val2

def diff_report(keys, columns, diffs, diff_limit, epsilon):
    schema = qztable.Schema(keys + ['Field', 'LValue', 'RValue'], ['string'] * (len(keys) + 3))
    report_tbl = qztable.Table(schema)
    
    all_columns = diffs.getSchema().columnNames
    
    column_idx = {}
    for idx, col in enumerate(all_columns):
        column_idx[col] = idx
    
    diff_cnt = 0
    for row in diffs:
        diff_cnt = diff_cnt + 1

        for col in columns:
            if col in keys:
                continue
                
            value1 = row[column_idx[col]]
            col_ = '%s_0' % col
            value2 = row[column_idx[col_]]

            line = None
            if type(value1) == float and type(value2) == float:
                if abs(value1 - value2) > epsilon:
                    line = generate_diff_row(keys, row, column_idx, col, value1, value2)
            elif not equals(value1, value2, epsilon):
                line = generate_diff_row(keys, row, column_idx, col, value1, value2)
                
            if line:
                report_tbl.append(line)
        
        if diff_cnt >= diff_limit:
            break
            
    return report_tbl            

def diffSchemas(l_schema, r_schema):
    l_cols = set(l_schema.columnNames)
    r_cols = set(r_schema.columnNames)
    
    only_in_l_table_cols = list(l_cols - r_cols)
    only_in_r_table_cols = list(r_cols - l_cols)
    
    l_table_col_dict = {}
    for idx, col in enumerate(l_schema.columnNames):
        l_table_col_dict[col] = l_schema.columnTypes[idx]
    
    r_table_col_dict = {}
    for idx, col in enumerate(r_schema.columnNames):
        r_table_col_dict[col] = r_schema.columnTypes[idx]
    
    type_mismatch_tbl = qztable.Table(qztable.Schema(['Field', 'LType', 'RType'], ['string'] * 3))
    
    for col in l_cols & r_cols:
        if l_table_col_dict[col] <> r_table_col_dict[col]:
            type_mismatch_tbl.append((col, l_table_col_dict[col], r_table_col_dict[col]))
            
    return only_in_l_table_cols, only_in_r_table_cols, type_mismatch_tbl

def diffQZTables(l_table, r_table, keycols, ignorecols=None, diff_limit=DEFAULT_DIFF_LIMIT, epsilon=0.0001, strict_schema=False):
    '''            
    :param l_table    
    :param r_table    
    :param keycols          list of key columns (comma-delimited) e.g. TradeId, BookMapId
    :param ignorecols       list of columns not to compare (comma-delimited) e.g. LastEditDate, Callable
    :param diff_limit       number of rows with mismatching columns that are reported in output screen
    :param epsilon          floating comparison tolerance
    :param strict_schema    If False, all columns where that type didn't watch will be excluded in the comparison (added to ignorecols)
                            If True, comparison will only be performed if the schemas of the two tables match

    Note: Everything is done in-memory so if your files are extremely big, don't use this tool
    '''
    if not keycols:
        raise RuntimeError("key is a required field")
        
    keys = [db_friendly_column_name(e) for e in keycols.split(',')]

    # columns to compare
    columns = list(l_table.getSchema().columnNames)
    
    if ignorecols:
        ignores = [db_friendly_column_name(e) for e in ignorecols.split(',')]
    else:
        ignores = []

    only_in_l_table_cols, only_in_r_table_cols, col_type_mismatch_tbl = diffSchemas(l_table.getSchema(), r_table.getSchema())

    if l_table.getSchema() != r_table.getSchema():
        print 'Schemas are not matching'
        print 'Fields only in left table: %s' % (str(only_in_l_table_cols))
        print 'Fields only in right table: %s' % (str(only_in_r_table_cols))
        print 'Fields with mismatch types'
        print col_type_mismatch_tbl

    if not strict_schema and any((only_in_l_table_cols, only_in_r_table_cols, col_type_mismatch_tbl)):
        print 'Running with strict_schema=False. Fields with mismatch types and fields that only exist in one of the two tables will not be compared'
        ignores.extend(only_in_l_table_cols)
        ignores.extend(only_in_r_table_cols)
        ignores.extend(col_type_mismatch_tbl.colToList(col='Field'))
    
    columns = list(set(columns) - set(ignores))
    
    lhsdiffs, diffs, rhsdiffs = cmp(l_table, r_table, keys=keys, columns=columns)
    
    if diffs:
        field_by_field_diffs = diff_report(keys, columns, diffs, diff_limit, epsilon)
    else:
        field_by_field_diffs = None
            
    return lhsdiffs, rhsdiffs, field_by_field_diffs, only_in_l_table_cols, only_in_r_table_cols, col_type_mismatch_tbl

def diff(filename1, filename2, keycols, ignorecols=None, delimiter='\t', skip_header=0, diff_limit=DEFAULT_DIFF_LIMIT, epsilon=0.0001, 
         generate_csv_report=True, csv_report_dir='c:/temp', strict_schema=False):

    tbl1 = fileToQztable(filename1, delimiter=delimiter, skip_header=skip_header)
    tbl2 = fileToQztable(filename2, delimiter=delimiter, skip_header=skip_header)
    
    print '%s has %d rows' % (filename1, tbl1.nRows())
    print '%s has %d rows' % (filename2, tbl2.nRows())
    print '\n'
    
    lhsdiffs, rhsdiffs, field_by_field_diffs, only_in_l_table_cols, only_in_r_table_cols, col_type_mismatch_tbl = diffQZTables(tbl1, tbl2, keycols, 
                                                                                ignorecols=ignorecols, diff_limit=diff_limit, 
                                                                                epsilon=epsilon, strict_schema=strict_schema)
                 
    if lhsdiffs:
        print 'Rows in %s that are not in %s' % (filename1, filename2)
        print lhsdiffs
        print '\n'

    if rhsdiffs:
        print 'Rows in %s that are not in %s' % (filename2, filename1)
        print rhsdiffs
        print '\n'
    
    if field_by_field_diffs:
        print 'Total number of fields that are not matching: %d' % len(field_by_field_diffs)
        print 'Only the first %d records with mismatching fields(s) are reported. To increase the number of records, change the value of diff_limit.' \
            % (diff_limit)
        print '\n'
        print field_by_field_diffs
                 
    if generate_csv_report:
        from os.path import basename, splitext
        
        csv_files = []
        
        if lhsdiffs:
            only_in_file1 = '%s/only_in_%s_ldiff.csv' % (csv_report_dir, splitext(basename(filename1))[0])
            csv_files.append(only_in_file1)
            saveQZTableAsCSV(only_in_file1, lhsdiffs)

        if rhsdiffs:
            only_in_file2 = '%s/only_in_%s_rdiff.csv' % (csv_report_dir, splitext(basename(filename2))[0])
            csv_files.append(only_in_file2)
            saveQZTableAsCSV(only_in_file2, rhsdiffs)
        
        if field_by_field_diffs:
            diff_file = '%s/diff_%s_vs_%s.csv' % (csv_report_dir, splitext(basename(filename1))[0], splitext(basename(filename2))[0])
            csv_files.append(diff_file)
            saveQZTableAsCSV(diff_file, field_by_field_diffs)
        
        if csv_files:
            print 'The following diff reports have been generated: %s' % str(csv_files)

def run(filename1, filename2, keycols, ignorecols=None, delimiter='\t', skip_header=0, diff_limit=DEFAULT_DIFF_LIMIT, epsilon=0.0001, 
        generate_csv_report=True, csv_report_dir='c:/temp', strict_schema=False):
    diff(filename1=filename1, filename2=filename2, keycols=keycols, ignorecols=ignorecols, skip_header=skip_header, diff_limit=diff_limit, epsilon=epsilon, 
         generate_csv_report=generate_csv_report, csv_report_dir=csv_report_dir, strict_schema=strict_schema)
    
def main():
    bobfns.run(run)
