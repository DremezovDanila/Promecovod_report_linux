from pyModbusTCP.client import ModbusClient
from pyModbusTCP import utils
#from fpdf import FPDF
import fpdf
import time
import datetime
import os
import subprocess
from typing import Any
from threading import Thread
from calendar import monthrange
import math
from definitions import ROOT_DIR, root_dir

print(root_dir, type(root_dir))
print(ROOT_DIR, type(ROOT_DIR))

#  Set new directory with fonts for FPDF library.
fpdf.set_global("SYSTEM_TTFONTS", root_dir + "/" + "etc/fonts")


#  Create new class for PDF document based on original class from FPDF lib.
class Pdf1(fpdf.FPDF):

    #  Method creating report in PDF.
    #  table_data comprises list of data divided into columns and rows without '/n' in the end of rows.
    def create_report(self, report_data: list[list[str]], report_params: list, root_dir: str = '/home'):
        if type(report_data) != list or len(report_data) <= 1:
            #  data variable includes table head and table contents.
            report_data = [['Data', 'is', 'wrong'], [0, 0, 0]]
        date_and_time = datetime.datetime.now()
        #  Initialise list of file head.
        report_head = report_data.pop(0)
        #  Initialise list of file content.
        report_content = report_data
        #  Formatting pdf document.
        left_margin, top_margin, right_margin = 5, 5, -5
        margins = (left_margin, top_margin, right_margin)
        row_height = 6.2
        self.set_margins(left_margin, top_margin, right_margin)
        cur_x, cur_y = left_margin, top_margin
        self.set_xy(cur_x, cur_y)
        #  Create the first page.
        self.add_page()
        #  Download the font supporting Unicode and set it. *It has to be added before used.
        #self.add_font(family="NotoSans", style="", fname= root_dir + "/" + "etc/fonts/NOTOSANS-REGULAR.TTF", uni=True)
        self.add_font(family="NotoSans", style="", fname=root_dir + "/" + "etc/fonts/ARIALUNI.TTF", uni=True)
        #  Filling in general information.
        #  First line with current date on the left and company name on the right.
        self.set_font(family="NotoSans", size=14)
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size,
                  txt="{:%d.%m.%Y}".format(date_and_time), ln=0, align="L")
        self.cell(w=(self.w - (margins[0] + abs(margins[2]))) / 2, h=self.font_size, txt='ООО "Промэковод"', ln=0,
                  align="R")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Site name.
        self.set_font("NotoSans", size=20)
        self.cell(self.w - 20, self.font_size * 1.5, txt=report_params[1], ln=0, align="C")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Table title.
        self.set_font("NotoSans", size=14)
        self.cell(self.w - 20, self.font_size * 2, txt=f"Отчет по суточным расходам воды и электроэнергии", ln=0,
                  align="C")
        cur_x = left_margin
        cur_y += 10
        self.set_xy(cur_x, cur_y)
        #  Drawing a table.
        #  Find widths of columns in table using specified font.
        self.set_font("NotoSans", size=14)
        #  Minimum column width is 20, else the meaning of a number of lines is rough.
        #  Adding some width to meanings of table data.
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                if j == 0:
                    report_content[i][j] = report_content[i][j].rjust(3, " ")
                elif j == 1:
                    report_content[i][j] = report_content[i][j].rjust(12, " ")
                else:
                    #  Check column width coefficient for adequacy.
                    if int(site_params[7]) < 0:
                        column_width_coef = 0
                    elif int(site_params[7]) > 30:
                        column_width_coef = 30
                    else:
                        column_width_coef = int(site_params[7])
                    report_content[i][j] = report_content[i][j].rjust(6 + column_width_coef, " ")
        #  col_width equals to the width of the largest meaning in a respective column of the table.
        report_elem_width = list()
        for i in range(len(report_content)):
            report_elem_width.append([])
            for j in range(len(report_head)):
                report_elem_width[i].append(self.get_string_width(report_content[i][j]))
        #  Find the longest lines in columns and assign these values to head titles length.
        #  Reverse massive of contents length (columns into rows).
        report_elem_width_rev = [[0 for j in range(len(report_content))] for i in range(len(report_head))]
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                report_elem_width_rev[j][i] = report_elem_width[i][j]
        #  Finding widths of column name cells.
        report_head_width = list()
        for i in range(len(report_head)):
            #  Indents from edges of the cell equal 2.
            report_head_width.append(max(report_elem_width_rev[i]) + 4.0)
        #  Find how many lines head titles take.
        report_head_lines_num = list()
        report_head_len = list()
        for i in range(len(report_head)):
            report_head_len.append(self.get_string_width(report_head[i]) + 2.0)
            #  How many lines a string takes.
            if 0 < report_head_width[i] < report_head_len[i]:
                report_head_lines_num.append(math.ceil(report_head_len[i] / report_head_width[i]))
            else:
                report_head_lines_num.append(1)
        #  Draw a table head.
        #  Calculate how many Line Feed characters are used in every head cell relating to the highest head cell.
        for i in range(len(report_head)):
            self.multi_cell(w=report_head_width[i],
                            h=row_height,
                            txt="\n" + report_head[i] + "\n" * 2 + "\n" * (max(report_head_lines_num) - report_head_lines_num[i]),
                            border=1,
                            align='C')
            cur_x += report_head_width[i]
            self.set_xy(cur_x, cur_y)
        cur_x = left_margin
        #  Make top indent according to number of lines in a head cell.
        cur_y += row_height * (max(report_head_lines_num) + 2)
        self.set_xy(cur_x, cur_y)
        for i in range(len(report_content)):
            for j in range(len(report_head)):
                self.multi_cell(w=report_head_width[j], h=row_height, txt=report_content[i][j], border=1, align='C')
                cur_x += report_head_width[j]
                self.set_xy(cur_x, cur_y)
            cur_x = left_margin
            cur_y += row_height
            self.set_xy(cur_x, cur_y)


    #  Override footer method so that it numerates pages.
    def footer(self):
        self.set_y(-15)
        self.set_x(10)
        self.set_font(family="NotoSans", size=12)
        self.cell(w=self.w - 20, h=10, txt='Страница %s из ' % self.page_no() + '{nb}', align='C')


#  Function for checking config file and adding site parameters to the list of site parameters.
def get_sites_params(config_path: str = ROOT_DIR / "configs/config"):
    #  Check config text file and get the list of row sites data.
    if os.path.isfile(config_path):
        config_file_exist = True
        with open(config_path, "r", encoding="UTF-8") as config_file:
            config_list = config_file.readlines()
            config_list_len = len(config_list)
            sites_row_params = list()
            for num in range(config_list_len):
                if config_list[num].split(';')[0] == "site name":
                    if num + 1 <= config_list_len:
                        site_row_params_list = config_list[num + 1].split(';')
                        site_row_params_list.pop(-1)
                        if len(site_row_params_list) == 8:
                            sites_row_params.append(site_row_params_list)
    else:
        config_file_exist = False

    #  Check correctness of row sites data and create a list of sites data.
    if config_file_exist and len(sites_row_params) > 0:
        config_params_off = False
    else:
        config_params_off = True

    sites_params = list()
    if not config_params_off:
        for site_row_params in sites_row_params:
            site_params_off = False
            site_params = list()
            try:
                site_params.append(site_row_params[0])  #  Site name.
                site_params.append(site_row_params[1])  #  Site name in Russian.
                site_params.append(site_row_params[2])  #  Report text file directory.
                site_params.append(site_row_params[3])  #  IP address and TCP port.
                site_ip_str = site_row_params[3].split(':')[0]
                int(site_row_params[3].split(':')[1])
                site_ip_row_list = site_ip_str.split('.')
                if len(site_ip_row_list) == 4:
                    for el in site_ip_row_list:
                        if 0 < int(el) < 255:
                            pass
                        else:
                            site_params_off = True
                            break
                else:
                    site_params_off = True
                site_params.append(int(site_row_params[4]))  #  Parameter number.
                site_params.append(site_row_params[5])  #  PDF report directory.
                site_params.append(site_row_params[6])  #  Columns names.
                if len(site_row_params[6].split(":")) != site_params[4]:
                    site_params_off = True
                site_params.append(int(site_row_params[7]))  #  Column width coefficient (int, 0 - 30).
            except:
                site_params_off = True
            finally:
                if not site_params_off:
                    sites_params.append(site_params)

    return config_params_off, sites_params


#  Global variables.
finish_main_process: bool = False
restore_report_start: bool = False
restore_report_site_name: str = ""
restore_report_date: str = ""
print_pdf_start: bool = True
print_pdf_site_name: str = ""
print_pdf_report_date: str = ""
command: str = ""

#  Cyclic function implementing console for operating the program.
def operate_program():
    global finish_main_process
    global restore_report_start
    global restore_report_site_name
    global restore_report_date
    global print_pdf_start
    global print_pdf_site_name
    global print_pdf_report_date
    global command

    while not finish_main_process:
        print("Enter your command:\n")
        command = input()
        #  Finish main process.
        if command == "finish" or command == "Finish" or command == "FINISH":
            finish_main_process = True
            command = ""
        #  Restore data.
        elif command == "restore" or command == "Restore" or command == "RESTORE":
            restore_report_site_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            restore_report_date = str(input("Year and month (ex.:2023_05): "))
            restore_report_start = True
            command = ""
        #  Print data.
        elif command == "print" or command == "Print" or command == "PRINT":
            print_pdf_site_name = str(input("Name of object (ex.:vzu_borodinsky): "))
            print_pdf_report_date = str(input("Year and month (ex.:2023_05): "))
            print_pdf_start = True
            command = ""


#  Get current date and time.
cur_datetime = datetime.datetime.now()
cur_date = cur_datetime.date()
cur_time = cur_datetime.time()
cur_time_hour = cur_time.hour
cur_time_min = cur_time.minute
cur_time_sec = cur_time.second
cur_date_day = datetime.datetime.today().day
cur_date_month = datetime.datetime.today().month
cur_date_year = datetime.datetime.today().year
report_cur_datetime = datetime.datetime.now()
#  Log parameters.
logs_num_line = 0
logs_separator = " " * 4

#  Set new directory with fonts for FPDF library.
#fpdf.set_global("SYSTEM_TTFONTS", root_dir + "/" + "etc/fonts")

#  Start another thread for processing console commands.
th = Thread(target=operate_program, args=())
th.start()

if not os.path.isdir(ROOT_DIR / "logs"):
    os.makedirs(ROOT_DIR / "logs")

if not os.path.isdir(ROOT_DIR / "reports"):
    os.makedirs(ROOT_DIR / "reports")

if not os.path.isdir(ROOT_DIR / "pdf reports"):
    os.makedirs(ROOT_DIR / "pdf reports")

with open(ROOT_DIR / "logs/logs", "w", encoding="UTF-8") as log:
    log.write("logs start...\n")

#  Main loop.
if __name__ == "__main__":
    while not finish_main_process:
        # Get current date and time.
        cur_datetime = datetime.datetime.now()
        cur_date = cur_datetime.date()
        cur_time = cur_datetime.time()
        cur_time_hour = cur_time.hour
        cur_time_min = cur_time.minute
        cur_time_sec = cur_time.second
        cur_date_day = datetime.datetime.today().day
        cur_date_month = datetime.datetime.today().month
        cur_date_year = datetime.datetime.today().year
        #  Check config text file and get the list of sites data.
        config_params_off, sites_params = get_sites_params()

        #  Start polls once a day at specified time (next day after report get formed in PLC). PLC draws report at
        #  23:59:59 and client polls start several minutes later.
        #if cur_time_hour == 0 and cur_time_min == 5 and cur_time_sec == 0:
        if cur_time_sec == 0:
            #  Get date of previous day (report).
            if cur_date_day == 1 and cur_date_month == 1:
                report_last_year = cur_date_year - 1
                report_last_month = 1
                report_last_day = 31
            elif cur_date_day == 1 and cur_date_month != 1:
                report_last_year = cur_date_year
                report_last_month = cur_date_month - 1
                #  Last day of previous month.
                report_last_day = monthrange(int(report_last_year), int(report_last_month))[1]
            else:
                report_last_year = cur_date_year
                report_last_month = cur_date_month
                report_last_day = cur_date_day - 1
            #  Last day datetime (when last report got formed).
            report_last_datetime = report_cur_datetime
            #  New day datetime (when next report get formed).
            report_cur_datetime = datetime.datetime.now()

            #  Check config text file and get the list of sites data.
            config_params_off, sites_params = get_sites_params()

            #  ModbusTCP polling.
            for site_params in sites_params:
                #  Get number of days in forming report month.
                site_report_days_in_month = monthrange(int(report_last_year), int(report_last_month))[1]
                #  Create report list of a site.
                site_report_list2 = [['' for column in range(len(site_params[6].split(":")) + 2)]
                                     for row in range(site_report_days_in_month + 1)]
                #  Get column names from configuration file.
                site_report_col_names = ["№", "Дата"]
                site_report_col_names = site_report_col_names + site_params[6].split(":")
                #  Fill the first line in a site report.
                for col_num in range(len(site_report_col_names)):
                    site_report_list2[0][col_num] = site_report_col_names[col_num]
                #  Fill the first two columns in a site report.
                for day in range(site_report_days_in_month):
                    site_report_list2[day + 1][0] = f"{day + 1}"
                    site_report_list2[day + 1][1] = f"{(day + 1):02}.{report_last_month:02}.{report_last_year}"
                #  Get a directory for report text file of a site.
                site_report_dir = ROOT_DIR / f"{site_params[2]}/{site_params[0]}_{report_last_year}_{report_last_month:02}"
                #  Initialise MBTCP connection.
                site_mbtcp_client = ModbusClient(host=site_params[3].split(":")[0],
                                                 port=int(site_params[3].split(":")[1]),
                                                 unit_id=1,
                                                 timeout=8.0,
                                                 auto_open=True,
                                                 auto_close=True)

                #  Poll parameters for object.
                poll_list2 = []
                poll_list2_real = []
                is_poll_ok = True
                for param_num in range(site_params[4]):
                    #  Write year, month, parameter number.
                    poll_w_request = \
                        site_mbtcp_client.write_multiple_registers(0,
                                                                  [int(report_last_year),
                                                                             int(report_last_month),
                                                                             param_num + 1])
                    if not poll_w_request:
                        is_poll_ok = False
                        break
                    #  Read parameter values for month.
                    poll_r_request = site_mbtcp_client.read_input_registers(0, site_report_days_in_month * 2)
                    if poll_r_request:
                        #  Draw a list of registers.
                        poll_list2.append(poll_r_request)
                    else:
                        is_poll_ok = False
                        break
                #  If polls succeeded then format data.
                if is_poll_ok:
                    for poll_list in poll_list2:
                        #  Conversion list of 2int16 to list of long32.
                        poll_r_request_long_32 = utils.word_list_to_long(poll_list, big_endian=True)
                        #  Conversion of list of long32 to list of real.
                        poll_r_request_list_real = []
                        for elem in poll_r_request_long_32:
                            poll_r_request_list_real.append(utils.decode_ieee(int(elem)))
                        #  Add every list of reals to list2 of reals.
                        poll_list2_real.append(poll_r_request_list_real)
                    #  Convert rows to columns in list2 of real values.
                    poll_list2_real_rows = len(poll_list2_real)
                    if poll_list2_real_rows > 0:
                        poll_list2_real_columns = len(poll_list2_real[0])
                    else:
                        poll_list2_real_columns = 0
                    poll_list2_real_converted = [[0.0 for col in range(poll_list2_real_rows)]
                                                 for row in range(poll_list2_real_columns)]
                    for row_num in range(poll_list2_real_rows):
                        for col_num in range(poll_list2_real_columns):
                            poll_list2_real_converted[col_num][row_num] = poll_list2_real[row_num][col_num]
                    #  Fill a site report by polled data.
                    for day in range(site_report_days_in_month):
                        for param_num in range(site_params[4]):
                            site_report_list2[day + 1][param_num + 2] = f"{format(poll_list2_real_converted[day][param_num], '.1f')}"
                    #  Write data into report text file.
                    with open(site_report_dir, "w", encoding="UTF-8") as site_report_file:
                        for row in site_report_list2:
                            for elem in row:
                                site_report_file.write(f"{elem};")
                            site_report_file.write("\n")
                else:
                    with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"{site_params[0]} site ModbusTCP poll error.\n")

            #  PDF document printing.
            for site_params in sites_params:
                site_report_dir = ROOT_DIR / f"{site_params[2]}/{site_params[0]}_{report_last_year}_{report_last_month:02}"
                if os.path.isfile(site_report_dir):
                    with open(site_report_dir, "r", encoding="UTF-8") as site_report_file:
                        site_report_data_list = site_report_file.readlines()
                        site_report_data_list2 = list()
                        for row in site_report_data_list:
                            site_report_data_list2.append(row.split(";")[:-1])
                    #  Declare instance of Pdf1.
                    report_pdf = Pdf1()
                    #  Set {nb} alias for PDF document footer.
                    report_pdf.alias_nb_pages()
                    report_pdf.create_report(site_report_data_list2, site_params, root_dir)
                    if os.path.isdir(f"/mnt/ics/Отчетность по работе станций/{site_params[1]}"):
                        report_pdf.output(f"/mnt/ics/Отчетность по работе станций/{site_params[1]}/{site_params[1]}_{report_last_year}_{report_last_month:02}.pdf")
                    else:
                        if os.path.isdir(ROOT_DIR / f"{site_params[5]}/{site_params[1]}"):
                            report_pdf.output(ROOT_DIR / f"{site_params[5]}/{site_params[1]}/{site_params[1]}_{report_last_year}_{report_last_month:02}.pdf")
                        else:
                            os.makedirs(ROOT_DIR / f"{site_params[5]}/{site_params[1]}")
                            report_pdf.output(ROOT_DIR / f"{site_params[5]}/{site_params[1]}/{site_params[1]}_{report_last_year}_{report_last_month:02}.pdf")
            time.sleep(10.0)
        time.sleep(1.0)

        #  Print PDF report from text file.
        if print_pdf_start:
            print_pdf_start = False
            print_report_input_data_ok = False
            #  Check input console data for relevance.
            try:
                print_pdf_report_date_list = print_pdf_report_date.split("_")
                print_pdf_report_date_year = print_pdf_report_date_list[0]
                print_pdf_report_date_month = print_pdf_report_date_list[1]
                print_pdf_report_date_year_int = int(print_pdf_report_date_year)
                print_pdf_report_date_month_int = int(print_pdf_report_date_month)
            except:
                with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                               f"Incorrect input data for printing PDF report.\n")
            else:
                if len(print_pdf_report_date_list) == 2 and \
                        len(print_pdf_report_date_year) == 4 and \
                        (len(print_pdf_report_date_month) == 2 or len(print_pdf_report_date_month) == 1) and \
                        2022 < print_pdf_report_date_year_int < 2040 and \
                        (1 <= print_pdf_report_date_month_int <= 12):
                    print_report_input_data_ok = True
                else:
                    with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"Incorrect date or site name for printing PDF report.\n")

            if print_report_input_data_ok:
                #  Get path to text file.
                print_report_data_path = ROOT_DIR / f"reports/{print_pdf_site_name}_{print_pdf_report_date_year_int}_{print_pdf_report_date_month_int:02}"
                if os.path.isfile(print_report_data_path):
                    with open(print_report_data_path, "r", encoding="UTF-8") as print_report_file:
                        print_report_data_list = print_report_file.readlines()
                        print_report_data_list2 = list()
                        for row in print_report_data_list:
                            print_report_data_list2.append(row.split(";")[:-1])

                    config_params_off, sites_params = get_sites_params()
                    print_report_site_params = list()
                    if not config_params_off:
                        for site_params in sites_params:
                            if site_params[0] == print_pdf_site_name:
                                print_report_site_params = site_params
                                break

                    if print_report_site_params != list():
                        #  Declare instance of Pdf1.
                        print_pdf_file = Pdf1()
                        #  Set {nb} alias for PDF document footer.
                        print_pdf_file.alias_nb_pages()
                        print_pdf_file.create_report(print_report_data_list2, print_report_site_params, root_dir)
                        if os.path.isdir(f"/mnt/ics/Отчетность по работе станций/{print_report_site_params[1]}"):
                            print_pdf_file.output(f"/mnt/ics/Отчетность по работе станций/{print_report_site_params[1]}/{print_report_site_params[1]}_{print_pdf_report_date_year_int}_{print_pdf_report_date_month_int:02}.pdf")
                        else:
                            if os.path.isdir(ROOT_DIR / f"{print_report_site_params[5]}/{print_report_site_params[1]}"):
                                print_pdf_file.output(ROOT_DIR / f"{print_report_site_params[5]}/{print_report_site_params[1]}/{print_report_site_params[1]}_{print_pdf_report_date_year_int}_{print_pdf_report_date_month_int:02}.pdf")
                            else:
                                os.makedirs(ROOT_DIR / f"{print_report_site_params[5]}/{print_report_site_params[1]}")
                                print_pdf_file.output(ROOT_DIR / f"{print_report_site_params[5]}/{print_report_site_params[1]}/{print_report_site_params[1]}_{print_pdf_report_date_year_int}_{print_pdf_report_date_month_int:02}.pdf")


        #  Restore report of certain date from PLC.
        if restore_report_start:
            restore_report_start = False
            restore_report_input_data_ok = False
            #  Check input console data for relevance.
            try:
                restore_report_date_list = restore_report_date.split("_")
                restore_report_date_year = int(restore_report_date_list[0])
                restore_report_date_month = int(restore_report_date_list[1])
            except:
                with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                    logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                f"Incorrect input data for restoring report.\n")
            else:
                if len(restore_report_date_list) == 2 and \
                        2022 < restore_report_date_year < 2040 and \
                        (1 <= restore_report_date_month <= 12):
                    restore_report_input_data_ok = True
                else:
                    with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                    f"Incorrect date or site name for restoring report.\n")

            #  Find site parameters.
            is_restore_report_site_exist = False
            restore_report_site_params = list()
            for site_params in sites_params:
                if site_params[0] == restore_report_site_name:
                    is_restore_report_site_exist = True
                    for param in site_params:
                        restore_report_site_params.append(param)
                    break

            if is_restore_report_site_exist and restore_report_input_data_ok:
                #  Poll data from site PLC.
                #  Get number of days in forming report month.
                restore_report_days_in_month = monthrange(restore_report_date_year, restore_report_date_month)[1]
                #  Create report list of a site.
                restore_report_list2 = [['' for column in range(len(restore_report_site_params[6].split(":")) + 2)]
                                        for row in range(restore_report_days_in_month + 1)]
                #  Get column names from configuration file.
                restore_report_col_names = ["№", "Дата"]
                restore_report_col_names = restore_report_col_names + restore_report_site_params[6].split(":")
                #  Fill the first line in a site report.
                for col_num in range(len(restore_report_col_names)):
                    restore_report_list2[0][col_num] = restore_report_col_names[col_num]
                #  Fill the first two columns in a site report.
                for day in range(restore_report_days_in_month):
                    restore_report_list2[day + 1][0] = f"{day + 1}"
                    restore_report_list2[day + 1][1] = f"{(day + 1):02}.{restore_report_date_month:02}.{restore_report_date_year}"
                #  Get a directory for report text file of a site.
                site_report_dir = ROOT_DIR / f"{restore_report_site_params[2]}/{restore_report_site_params[0]}_{restore_report_date_year}_{restore_report_date_month:02}"
                #  Initialise MBTCP connection.
                site_mbtcp_client = ModbusClient(host=restore_report_site_params[3].split(":")[0],
                                                 port=int(restore_report_site_params[3].split(":")[1]),
                                                 unit_id=1,
                                                 timeout=8.0,
                                                 auto_open=True,
                                                 auto_close=True)

                #  Poll parameters for object.
                poll_list2 = []
                poll_list2_real = []
                is_poll_ok = True
                for param_num in range(restore_report_site_params[4]):
                    #  Write year, month, parameter number.
                    poll_w_request = \
                        site_mbtcp_client.write_multiple_registers(0,
                                                                    [int(restore_report_date_year),
                                                                    int(restore_report_date_month),
                                                                    param_num + 1])
                    if not poll_w_request:
                        is_poll_ok = False
                        break
                    #  Read parameter values for month.
                    poll_r_request = site_mbtcp_client.read_input_registers(0, restore_report_days_in_month * 2)
                    #  Draw a list of registers.
                    if poll_r_request:
                        poll_list2.append(poll_r_request)
                    else:
                        is_poll_ok = False
                        break
                #  If polls succeeded then format data.
                if site_mbtcp_client.last_error == 0:
                    for poll_list in poll_list2:
                        #  Conversion list of 2int16 to list of long32.
                        poll_r_request_long_32 = utils.word_list_to_long(poll_list, big_endian=True)
                        #  Conversion of list of long32 to list of real.
                        poll_r_request_list_real = []
                        for elem in poll_r_request_long_32:
                            poll_r_request_list_real.append(utils.decode_ieee(int(elem)))
                        #  Add every list of reals to list2 of reals.
                        poll_list2_real.append(poll_r_request_list_real)
                    #  Convert rows to columns in list2 of real values.
                    poll_list2_real_rows = len(poll_list2_real)
                    if poll_list2_real_rows > 0:
                        poll_list2_real_columns = len(poll_list2_real[0])
                    else:
                        poll_list2_real_columns = 0
                    poll_list2_real_converted = [[0.0 for col in range(poll_list2_real_rows)]
                                                 for row in range(poll_list2_real_columns)]
                    for row_num in range(poll_list2_real_rows):
                        for col_num in range(poll_list2_real_columns):
                            poll_list2_real_converted[col_num][row_num] = poll_list2_real[row_num][col_num]
                    #  Fill a site report by polled data.
                    for day in range(restore_report_days_in_month):
                        for param_num in range(restore_report_site_params[4]):
                            restore_report_list2[day + 1][param_num + 2] = f"{format(poll_list2_real_converted[day][param_num], '.1f')}"
                    #  Write data into report text file.
                    with open(site_report_dir, "w", encoding="UTF-8") as restore_report_file:
                        for row in restore_report_list2:
                            for elem in row:
                                restore_report_file.write(f"{elem};")
                            restore_report_file.write("\n")
                else:
                    with open(ROOT_DIR / "logs/logs", "a", encoding="UTF-8") as logs:
                        logs.write(f"{datetime.datetime.now().strftime('%Y.%m.%d  %H:%M:%S')}{logs_separator}"
                                   f"{restore_report_site_params[0]} site ModbusTCP poll error.\n")


