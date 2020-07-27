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
        self.location_number = None
        self.each_section_range = None #Takes each range | Determine if even or odd numbers required in locations
        self.each_location_range = None #Takes each range | Determine if even or odd numbers required in locations
        self.step = [1, 5, 1, 3, 1, 3, 1, 1]

        self.section_number = None #Each section element in for loop, main() method

        self.isle_number = None
    def file_handle(self, rewrite=True):

        if REWRITE == True:
            # print('WENT HERE')
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




    #
    def main(self): #If passed, makes numbers even
        self.file_handle()
        self.set_table_header(['LOCATIONS',	'LOCATION TYPE', 'CLIENT', 'HIGH/LOW', 'RESTRICTION', 'LOCUS YES/NO'])
        for each_section_range in SECTION_RANGE:
            # print(each_section_range)
            # print(each_section_range)
            # each_section_range[0] - START
            # each_section_range[1] - END
            self.each_section_range = each_section_range
            for number in range(each_section_range[0], each_section_range[1] + 1 if each_section_range[0] != each_section_range[1] else 0): #range method doesn't include last number, so +1 added
                print(number)
                self.isle_number = number
                self.section_writer(isle=self.isle_number, low_or_high=True, location_index=SECTION_RANGE.index(each_section_range))

            #FOR HIGH LOCATIONS
            LOW_POSITION = False
            self.each_section_range = each_section_range
            for number in range(each_section_range[0], each_section_range[1] + 1 if each_section_range[0] != each_section_range[1] else 0): #range method doesn't include last number, so +1 added

                self.isle_number = number
                self.section_writer(isle=self.isle_number, low_or_high=False, location_index=SECTION_RANGE.index(each_section_range))


    def section_writer(self, isle, low_or_high, location_index):
        # print(self.each_location_range)
        # step = [1, 5, 1, 3, 1, 3, 1, 1] #Pattern: Gaylordx1, Gaylordx1, Gaylordx1, BinStackx5 for odd and even
        # [ 6, 4, 4, 2]
        #[7, 5, 5, 3]
        # print(isle)
        start = LOCATION_RANGE[location_index][0]
        end = LOCATION_RANGE[location_index][1]
        even_or_odd = LOCATION_RANGE[location_index][2]
        self.x = start
        # print(start)
        # print(end)
        # print(self.each_location_range)
        while self.x <= end: #counts within the range
            # print('HERE')
            # location_type = 'Gaylord Location'
            self.inner_counter = 1
            position_letter = 'A' if low_or_high == True else 'F'

            even_GL = True if self.isle_number % 2 == 0 else False
            # print(isle)
            self.section_writer_handler(isle, position_letter, low_or_high, even_or_odd=even_or_odd, even_GL=even_GL)


    def section_writer_handler(self, isle, position_letter, low_or_high, even_or_odd, even_GL):
        location_type = 'Gaylord Location'
        for number in self.step: #counts within each section
            # print('x=>',str(x))
            if self.inner_counter == 7 or self.inner_counter == 8:
                print(isle)
                # if even_or_odd != 0 and self.x % 2 != 0:
                #     pass
                # else:
                # print('inner_counter=>'+str(inner_counter))
                location_type = 'Bin locations'
                # print(LOW_POSITION)
                letter_set =  ['A','B','C','D','E'] if low_or_high == True else ['F','G','H','I','J']
                # print(str(letter_set))
                for each_letter in letter_set:
                    final_location_name = self.location_concat(isle, str(self.x), each_letter)
                    n = str(self.collumn_counter) #collumn incremination to string

                    self.worksheet.write('A'+n, final_location_name)
                    self.worksheet.write('B'+n, location_type)
                    self.worksheet.write('C'+n, CLIENT)
                    self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
                    self.worksheet.write('E'+n, RESTRICTION)
                    self.worksheet.write('F'+n, LOCUS)

                    self.collumn_counter+=1
            else:
                # if self.odd_or_even == 0: #IF isle is even, skip printing Gaylords
                #     pass
                # else:
                if even_GL == True:
                    pass
                else:
                    # print(self.isle_number)
                    # print(even_GL)
                    print(isle)
                    # if even_or_odd != 1 and self.x % 2 != 1:
                    #     pass
                    # else:
                    n = str(self.collumn_counter) #collumn incremination to string
                    final_location_name = self.location_concat(isle, str(self.x), position_letter)

                    self.worksheet.write('A'+n, final_location_name)
                    self.worksheet.write('B'+n, location_type)
                    self.worksheet.write('C'+n, CLIENT)
                    self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
                    self.worksheet.write('E'+n, RESTRICTION)
                    self.worksheet.write('F'+n, LOCUS)

                    self.collumn_counter+=1
                    self.inner_counter+=1
            self.x += number #increasing count range
            # print(self.x)

    def location_concat(
        self,
        isle,
        location_number,
        position_letter,
        ):
        isle_string = str(isle)
        isle_number_length = list(isle_string)
        whole_location = ''
        # print(len(isle_number_length))
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

    def close_workbook(self):
        self.workbook.close()

#JSON OUTPUT:

# self.locations[final_location_name] =  {
#         'location' : final_location_name,
#         'location_type' : location_type,
#         'client' : CLIENT,
#         'position_real' : 'Low' if LOW_POSITION == True else 'High',
#         'restriction' : RESTRICTION,
#         'locus' : LOCUS
#     }


def writer_handler(self, number, low_or_high, range_index):
    self.location_number = number
    self.odd_or_even = 0 if number % 2 == 0 else 1
    # for each_location in LOCATION_RANGE[range_index]:
        # print(each_location_range)
    # if each_location_range[2] == 0:
    #     # print(number)
    #     if number % 2 == 0:
    #         self.each_location_range = each_location_range
    # elif each_location_range[2] == 1:
    #     # print(number)
    #     if number % 2 == 1:
    #         self.each_location_range = each_location_range
    #     pass
    # self.each_location_range = each_location_range
    # print(each_location_range)
    # self.section_writer(start=self.each_location_range[0], end=self.each_location_range[0], isle=self.isle_number, low_or_high=low_or_high)
    self.section_writer(start=LOCATION_RANGE[range_index][0], end=LOCATION_RANGE[range_index][1], isle=self.isle_number, low_or_high=low_or_high)



# for number in step: #counts within each section
#     # print('x=>',str(x))
#     if inner_counter == 7 or inner_counter == 8:
#         # print('inner_counter=>'+str(inner_counter))
#         location_type = 'Bin locations'
#         # print(LOW_POSITION)
#         letter_set =  ['A','B','C','D','E'] if low_or_high == True else ['F','G','H','I','J']
#         # print(str(letter_set))
#         for each_letter in letter_set:
#             final_location_name = self.location_concat(isle, str(x), each_letter)
#             n = str(self.collumn_counter) #collumn incremination to string
#
#             self.worksheet.write('A'+n, final_location_name)
#             self.worksheet.write('B'+n, location_type)
#             self.worksheet.write('C'+n, CLIENT)
#             self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
#             self.worksheet.write('E'+n, RESTRICTION)
#             self.worksheet.write('F'+n, LOCUS)
#
#             self.collumn_counter+=1
#     else:
#         if self.odd_or_even == 0: #IF isle is even, skip printing Gaylords
#             pass
#         else:
#             n = str(self.collumn_counter) #collumn incremination to string
#             final_location_name = self.location_concat(isle, str(x), position_letter)
#
#             self.worksheet.write('A'+n, final_location_name)
#             self.worksheet.write('B'+n, location_type)
#             self.worksheet.write('C'+n, CLIENT)
#             self.worksheet.write('D'+n, 'Low' if low_or_high == True else 'High')
#             self.worksheet.write('E'+n, RESTRICTION)
#             self.worksheet.write('F'+n, LOCUS)
#
#             self.collumn_counter+=1
#         inner_counter+=1
#     x += number #increasing count range
