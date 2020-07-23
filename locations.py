import xlsxwriter
import os
from settings import *



class Locations():

    def __init__(self):
        self.locations = {}
        self.workbook = None
        self.worksheet = None
        self.collumn_counter = 2
        self.x = None # Location iteration
        self.inner_counter = None # Gaylords and Bins iterator
        self.odd_or_even = None #If even location print only bin locations

    def file_handle(self, rewrite=True):

        if REWRITE == True:
            print('WENT HERE')
            if os.path.isfile('./'+FILE_NAME):
                try:
                    os.remove(FILE_NAME)
                except OSError:
                    print('!_____ERROR_____!')
                    print("You probably havn't close the file. Close it and try again." )
                    exit(1)
                self.open_xlsx_file(True)
            else:
                self.open_xlsx_file(True)
        else:
            self.open_xlsx_file(False)





    def main(self): #If passed, makes numbers even
        self.file_handle()
        self.set_table_header(['LOCATIONS',	'LOCATION TYPE', 'CLIENT', 'HIGH/LOW', 'RESTRICTION', 'LOCUS YES/NO'])
        for number in range(SECTION_START, SECTION_END+1): #range method doesn't include last number, so +1 added
            self.odd_or_even = 0 if number % 2 == 0 else 1
            self.section_writer(start=LOCATION_START, end=LOCATION_END, isle=number, low_or_high=True)
        LOW_POSITION = False
        for number in range(SECTION_START, SECTION_END+1): #range method doesn't include last number, so +1 added
            self.odd_or_even = 0 if number % 2 == 0 else 1
            self.section_writer(start=LOCATION_START, end=LOCATION_END, isle=number, low_or_high=False)
            # self.section_writer(start=LOCATION_START + 1, end=LOCATION_END, isle=number)

    def close_workbook(self):
        self.workbook.close()

    # def location_writer(self, inner_counter):





    def section_writer(self, start, end, isle, low_or_high):
        x = LOCATION_START
        # step = [6, -5, 9, -3, 7, -3, 5, -1] #Pattern: Gaylordx1, Gaylordx1, Gaylordx1, BinStackx5
        step = [1, 5, 1, 3, 1, 3, 1, 1] #Pattern: Gaylordx1, Gaylordx1, Gaylordx1, BinStackx5
        # [ 6, 4, 4, 2]
        #[7, 5, 5, 3]
        while x < LOCATION_END: #counts within the range
            location_type = 'Gaylord Location'
            inner_counter = 1
            position_letter = 'A' if low_or_high == True else 'F'

            for number in step: #counts within each section
                # print('x=>',str(x))
                if inner_counter == 7 or inner_counter == 8:
                    # print('inner_counter=>'+str(inner_counter))
                    location_type = 'Bin locations'
                    # print(LOW_POSITION)
                    letter_set =  ['A','B','C','D','E'] if low_or_high == True else ['F','G','H','I','J']
                    # print(str(letter_set))
                    for each_letter in letter_set:
                        final_location_name = self.location_concat(isle, str(x), each_letter)
                        n = str(self.collumn_counter) #collumn incremination to string

                        self.worksheet.write('A'+n, final_location_name)
                        self.worksheet.write('B'+n, location_type)
                        self.worksheet.write('C'+n, CLIENT)
                        self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
                        self.worksheet.write('E'+n, RESTRICTION)
                        self.worksheet.write('F'+n, LOCUS)

                        self.collumn_counter+=1
                else:
                    if self.odd_or_even == 0:
                        pass
                    else:
                        n = str(self.collumn_counter) #collumn incremination to string
                        final_location_name = self.location_concat(isle, str(x), position_letter)

                        self.worksheet.write('A'+n, final_location_name)
                        self.worksheet.write('B'+n, location_type)
                        self.worksheet.write('C'+n, CLIENT)
                        self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
                        self.worksheet.write('E'+n, RESTRICTION)
                        self.worksheet.write('F'+n, LOCUS)

                        self.collumn_counter+=1
                    inner_counter+=1
                x += number #increasing count range



    def location_concat(
        self,
        isle,
        location_number,
        position_letter,
        ):
        isle_string = str(isle)
        isle_number_length = isle_string.split()
        whole_location = ''
        if len(isle_number_length) == 1:
            isle_string = '0' + isle_string
        whole_location = FIRST_LETTER + '-' + isle_string +  '-' + location_number + '-' + position_letter

        return whole_location

    def set_table_header(self, list_of_headers):
        set_bold = self.workbook.add_format({'bold': True})
        counter = 0
        for header in list_of_headers:
            # print(header)
            self.worksheet.set_column(0, counter, COLUMN_SIZE)
            self.worksheet.write(0, counter, header, set_bold )
            counter+=1

    def open_xlsx_file(self, rewrite):
        self.workbook = xlsxwriter.Workbook(FILE_NAME)
        self.worksheet = self.workbook.add_worksheet(SHEET_NAME)
        if rewrite == True:
            with open(FILE_NAME, 'x') as f:
                f.close()


#JSON OUTPUT:

# self.locations[final_location_name] =  {
#         'location' : final_location_name,
#         'location_type' : location_type,
#         'client' : CLIENT,
#         'position_real' : 'Low' if LOW_POSITION == True else 'High',
#         'restriction' : RESTRICTION,
#         'locus' : LOCUS
#     }
