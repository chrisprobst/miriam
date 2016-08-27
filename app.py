import openpyxl as x
import openpyxl.styles as xstyles

fz_path = '/Users/chrisprobst/Desktop/Fz.xlsx'
f4_path = '/Users/chrisprobst/Desktop/F4.xlsx'
input_path = '/Users/chrisprobst/Desktop/EEG_Auswertung.xlsx'
output_path = '/Users/chrisprobst/Desktop/EEG_Auswertung_generated.xlsx'

fzWB = x.load_workbook(fz_path)
f4WB = x.load_workbook(f4_path)
inputWB = x.load_workbook(input_path)

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

def insert_cell_into_output(id, percentage, column, new_value):
    main_sheet = inputWB.active
    column_names = main_sheet.rows[0]
    data_rows = main_sheet.rows[1:]

    for row in data_rows:
        if len(row) == 0:
            continue

        row_id = int(row[0].value)

        if row_id == id:
            for j, column_name in enumerate(column_names):
                if column_name.value != column:
                    continue

                c = row[j]
                c.value = new_value

                if percentage >= 40:
                    c.fill = redFill
                elif percentage >= 30:
                    c.fill = orangeFill

                return

            print("[VP%d] Could not find column: %s" % (id, column))
            return

    print("[VP%d] Could not id: %d" % (id, id))
    return

def copy_from_f_to_output(category, wb):
    for worksheet in wb.worksheets:
        id = int(worksheet.title[2:4])
        row_offset = 5

        while True:
            if len(worksheet.rows) < row_offset:
                break

            table_name = worksheet.rows[row_offset][0]
            table_beta1_mean = worksheet.rows[row_offset+row_beta1_offset][2]
            table_beta3_mean = worksheet.rows[row_offset+row_beta3_offset][2]
            table_percentage = worksheet.rows[row_offset+row_percentage_offset][0]

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

                lb_column_name = category + 'LB_' + table_name_value
                lb_cell_value = float(table_beta1_mean_value.strip())

                hb_column_name = category + 'HB_' + table_name_value
                hb_cell_value = float(table_beta3_mean_value.strip())

                insert_cell_into_output(id, percentage, lb_column_name, lb_cell_value)
                insert_cell_into_output(id, percentage, hb_column_name, hb_cell_value)

            row_offset += row_step

copy_from_f_to_output("F4", f4WB)
copy_from_f_to_output("FZ", fzWB)

inputWB.save(output_path)
