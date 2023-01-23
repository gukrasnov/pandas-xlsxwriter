import time
import pandas
import datetime

start_time = time.monotonic()

path_xlsx = 'test_1.xlsx'

worksheet = pandas.read_excel( path_xlsx, sheet_name = 0 )
df = pandas.DataFrame( worksheet )

path_xlsx = 'test_11.xlsx'

if path_xlsx == True:
    with pandas.ExcelWriter( path_xlsx, mode = 'a', engine = 'openpyxl', if_sheet_exists = 'overlay' ) as writer:
        df.to_excel( writer, sheet_name = 'Лист1', index = False, header = False, startrow = 1, freeze_panes = ( 1, 1 ) )
else:
    with pandas.ExcelWriter( path_xlsx, engine = 'xlsxwriter' ) as writer:
        df.to_excel( writer, sheet_name = 'Лист1', index = False, header = True, freeze_panes = ( 1, 1 ) )
        workbook = writer.book
        worksheet = writer.sheets[ 'Лист1' ]
        worksheet.set_row_pixels( 0, 100 )
        worksheet.autofilter( 0, 0, 0, len( df.columns ) - 1 )
        header_format = workbook.add_format( {
                                                'bold': True,
                                                'text_wrap': True,
                                                'align': 'center',
                                                'valign': 'vcenter',
                                                'bg_color': '#00B0F0',
                                                'border': 1
                                            } )
        for col_num, value in enumerate( df.columns.values ):
            worksheet.write( 0, col_num, value, header_format )

end_time = time.monotonic()
print( 'Время исполнения кода: ', datetime.timedelta( seconds = end_time - start_time ) )
