import openpyxl as x
import openpyxl.styles as xstyles

fz_path = '/Users/chrisprobst/Desktop/Fz.xlsx'
f4_path = '/Users/chrisprobst/Desktop/F4.xlsx'
input_path = '/Users/chrisprobst/Desktop/EEG_Auswertung.xlsx'
output_path = '/Users/chrisprobst/Desktop/EEG_Auswertung_generated.xlsx'

################################################################################
################################################################################

fzWB = x.load_workbook(fz_path)
f4WB = x.load_workbook(f4_path)
inputWB = x.load_workbook(input_path)

################################################################################
################################################################################

redFill = xstyles.PatternFill(start_color='FFFF0000',
                              end_color='FFFF0000',
                              fill_type='solid')

orangeFill = xstyles.PatternFill(start_color='FFFFC000',
                                 end_color='FFFFC000',
                                 fill_type='solid')

row_beta1_offset = 7
row_beta3_offset = 8
row_percentage_offset = 10
row_step = 12

################################################################################
############################### Build the cache ################################
################################################################################

row_index_cache = {}
column_index_cache = {}
result_rows = inputWB.active.rows[:]

def build_caches():
    column_names = result_rows[0][2:]
    data_rows = result_rows[1:]

    for i, row in enumerate(data_rows):
        row_index_cache[row[0].value] = i+1

    for i, column in enumerate(column_names):
        column_index_cache[column.value] = i+2

# Build caches
build_caches()

################################################################################
############################### Insert value into output #######################
################################################################################

def insert_value_into_output(row_index, column_index, percentage, new_value):
    if row_index not in row_index_cache:
        print("[VP%d] Could not find row: %d" % (row_index, row_index))
        return

    if column_index not in column_index_cache:
        print("[VP%d] Could not find column: %s" % (row_index, column_index))
        return

    i = row_index_cache[row_index]
    j = column_index_cache[column_index]
    c = result_rows[i][j]
    c.value = new_value

    if percentage > 40:
        c.fill = redFill
    elif percentage > 30:
        c.fill = orangeFill

################################################################################
############################### Copy from F to output ##########################
################################################################################

def copy_from_f_to_output(category, wb):
    print('Processing %s...' % category)

    for worksheet in wb.worksheets:
        row_index = int(worksheet.title[2:4])
        row_offset = 5
        rows = worksheet.rows[:]

        while True:
            if len(rows) < row_offset:
                break

            table_name = rows[row_offset][0]
            table_beta1_mean = rows[row_offset+row_beta1_offset][2]
            table_beta3_mean = rows[row_offset+row_beta3_offset][2]
            table_percentage = rows[row_offset+row_percentage_offset][0]

            table_name_value = table_name.value
            table_name_value = table_name_value.split(':')[1].split('(')[0][:-1]
            table_name_value = table_name_value.replace('.', '')

            if 'Baseline' in table_name_value:
                table_name_value = table_name_value.replace('Baseline', 'B')

            if 'Cue' not in table_name_value:
                table_beta1_mean_value = table_beta1_mean.value
                table_beta3_mean_value = table_beta3_mean.value
                table_percentage_value = table_percentage.value

                percentage = float(table_percentage_value.split('%')[0].strip())

                lb_column_index = category + 'LB_' + table_name_value
                lb_cell_value = float(table_beta1_mean_value.strip())

                hb_column_index = category + 'HB_' + table_name_value
                hb_cell_value = float(table_beta3_mean_value.strip())

                insert_value_into_output(row_index, lb_column_index, percentage, lb_cell_value)
                insert_value_into_output(row_index, hb_column_index, percentage, hb_cell_value)

            row_offset += row_step

################################################################################
############################### Main app #######################################
################################################################################

copy_from_f_to_output("F4", f4WB)
copy_from_f_to_output("FZ", fzWB)

################################################################################
################################################################################

inputWB.save(output_path)
