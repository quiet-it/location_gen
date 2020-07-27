#
FIRST_LETTER = 'B'

#NUMBER OF SECTIONS
SECTION_START = 3
SECTION_END = 26

#FIRST LIST ELEMENT = START
#SECOND = END
#THIRD IS (0 = EVEN, 1 = ODD, 2 = NONE)
SECTION_RANGE = [
                [3,6],
                [7,7],
                [7,7],
                [8,10],
                [11,14],
                [15,18],
                [19,19],
                [19,19],
                [20,25]
                ]
#START AND END OF LOCATIONS
LOCATION_RANGE = [
                 [273,304,2],
                 [274,304,0],
                 [273,351,1],
                 [273,352,2],
                 [273,320,2],
                 [273,352,2],
                 [274,352,0],
                 [273,367,1],
                 [273,368,2]
                 ]
# LOCATION_START = 161
# LOCATION_END = 252

EVEN_ODD = 'odd'


EXCEPTION_RANGE = []
#Position related settings and row headers.
LOW_POSITION = True
CLIENT = 'Mack Weldon'
RESTRICTION = 'Single SKU'
LOCUS = 'No'

#EXCEL FILE SETTINGS
# REWRITE = True #IF TRUE IT WILL CREATE NEW FILE, IF FALSE IT WILL ADD NEW SHEET TO EXISTING FILE
REWRITE = False #IF TRUE IT WILL CREATE NEW FILE, IF FALSE IT WILL ADD NEW SHEET TO EXISTING FILE
FILE_NAME='Mack_Weldon_Front_Sec.xlsx'

# SHEET_1_NAME = ''
# SHEET_2_NAME = ''
# if EVEN_ODD == 'odd':
#     SHEET_1_NAME='odd_locations'
# elif EVEN_ODD == 'all':
#     SHEET_1_NAME = 'odd_locations'
#     SHEET_2_NAME =  'even_locations'
# else:
    # SHEET_2_NAME = 'even_locations'
SHEET_NAME='Front_section'
COLUMN_SIZE = 22




## TODO: Separate for loop in dedicated function
# in section_writer function so odd and even numbers
# can go along
