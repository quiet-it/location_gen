from locations import *
from settings import *

# if EVEN_ODD == 'odd':
#     odd_worksheet = Locations()
#     odd_worksheet.main()
#     odd_worksheet.close_workbook()
# else:
worksheet = Locations()
    # even_worksheet = Locations()
worksheet.main()
# worksheet.main(sheet_name=['Odd_worksheet'])
    # even_worksheet.main(1, ['Odd_worksheet'])

    # odd_worksheet.close_workbook()
worksheet.close_workbook()
