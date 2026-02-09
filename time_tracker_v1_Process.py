from time import sleep
from openpyxl import load_workbook, Workbook
from datetime import date
from datetime import timedelta as td
from calendar import monthrange
from my_module.cellcontrol import Cell, Member
import numpy as np

# class "MyTime" which deals with the time descriptions
class MyTime:
    def __init__(self, time):
        '''time is given by the 24-hour format with "hhmm".'''
        self._time = time
        self._minute = time % 100
        self._hour = int((time-self._minute)/100)
    def __add__(self, other):
        hour = self.hour + other.hour
        minute = self.minute + other.minute
        if minute >= 60:
            minute -= 60
            hour += 1
        return MyTime(100*hour+minute)
    def __sub__(self, other):
        hour = self.hour - other.hour
        minute = self.minute - other.minute
        if minute < 0:
            minute += 60
            hour -= 1
        return MyTime(100*hour+minute)
    def __eq__(self, other):
        return self.time == other.time
    def __lt__(self, other):
        return self.time < other.time
    def __le__(self, other):
        return self.time <= other.time
    def __repr__(self):
        return f'{self.hour:02d}:{self.minute:02d}'
    def __str__(self):
        return self.__repr__()
    def __int__(self):
        return self.time
    @property
    def hour(self):
        return self._hour
    @property
    def minute(self):
        return self._minute
    @property
    def minutes(self):
        return self._hour*60+self._minute
    @property
    def time(self):
        return self._time
    @property
    def hour_ceil(self):
        result = self.hour
        if self.minute > 30:
            result += 1
        elif self.minute > 0:
            result += 0.5
        return result
    @property
    def hour_floor(self):
        result = self.hour
        if self.minute >= 30:
            result += 0.5
        return result
    @property
    def day_ceil(self):
        if self.hour <= 4:
            return 0.5
        else:
            return 1
        
# Parameters
if_late = {'遲':0}
day_off = {'特':0,'病':1,'事':2,'其他':3}
time1,time2,time3,time4,time5,time6 = MyTime(800),MyTime(1200),MyTime(1300),MyTime(1700),MyTime(1800),MyTime(2030)
labels = ['早上班','早下班','晚上班','晚下班','早遲','早假','午遲','午假']
statistic_labels = ['工作天','加班1.33','加班1.67','遲到(分)','遲到(天)','特休(天)','病假(時)','事假(時)','其他(時)']

# functions
def rapid_enter(key):
    '''Return the time according to the given key'''
    if key == 'a':
        return [time1,time2,time3,time4,-1,-1,-1,-1]
    elif key == 's':
        return [time1,time2,time3,time6,-1,-1,-1,-1]
    elif key == 'd':
        return [time1,time2,time3,time5,-1,-1,-1,-1]

def day_off_name(idx):
    if idx in [0,1,2,3]:
        return list(day_off.keys())[idx]
    
def late_name(idx):
    if idx == 0:
        return '遲'

def hour_error(hour, work_day):
    '''Check whether the data is valid or not
    Return True if the hour is invalid
    
    What this function can do :
    1. make sure nonzero hours are increasing
    2. make sure nonzero hours occupies correct positions
    3. make sure there's a day off information if there are empty hours
    
    Developer notice : this function cannot detect ALL the error concerning day off types'''
    if type(hour) != list:
        return True # hour must be a list
    elif hour[5] not in [-1,0,1,2,3] or hour[7] not in [-1,0,1,2,3]:
        return True # invalid day off input
    elif hour[4] not in [-1,0] or hour[6] not in [-1,0]:
        return True # invalid late input
    list_of_None = []
    if MyTime(0) in hour[:4]:
        list_of_None = list(filter(lambda x:hour[x] == MyTime(0), range(4)))
        if list_of_None not in [[0,1,2,3],[1,2],[0,1],[2,3]]:
            return True # invalid null time
    if not work_day:
        if hour[4:] == [-1,-1,-1,-1]:
            return False # valid holiday
        else:
            return True # holiday can't be late or be off
    elif list_of_None == [0,1,2,3] and hour[4] == hour[6] == -1 and (hour[5] != -1 and hour[7] != -1 or hour[5] == hour[7] == -1):
        return False # day off or no input
    elif list_of_None == [1,2] and hour[0] < hour[3] and hour[5] == hour[6] == -1:
        return False # office member standard form
    elif list_of_None == [0,1] and hour[2] < hour[3] and hour[4] == -1 and hour[5] != -1 and hour[7] != 0:
        return False # morning off
    elif list_of_None == [2,3] and hour[0] < hour[1] and hour[6] == -1 and hour[5] != 0 and hour[7] != -1:
        return False # noon off
    elif list_of_None == [] and hour[0] < hour[1] <= hour[2] < hour[3] and hour[5] != 0 and hour[7] != 0:
        return False
    return True

def daily_result(member_type, hour, work_day):
    if hour == [MyTime(0),MyTime(0),MyTime(0),MyTime(0),-1,-1,-1,-1]: # null input
        return None
    extra_extra, extra, late, off = 0,0,0,[0,0,0,0]
    if work_day:
        clear_list = []
        if 0 < int(hour[0]) <= 700: #早到
            if member_type != 'd':
                extra_extra += (time1-hour[0]).hour_floor
        elif time1 < hour[0]:
            if hour[4] == 0: # 遲到
                late += (hour[0]-time1).minutes
                clear_list.append(4)
            elif hour[5] != -1:
                off[hour[5]] += (hour[0]-time1).hour_ceil
                clear_list.append(5)
            elif hour[5] in [-1,0]:
                return False
        elif int(hour[0]) == 0:
            if hour[5] == 0:
                off[0] += 0.5
                clear_list.append(5)
            else:
                off[hour[5]] += 4
                clear_list.append(5)
        if 0 < int(hour[1]) < 1200:
            if hour[5] in [-1,0]:
                return False
            else:
                off[hour[5]] += (time2-hour[1]).hour_ceil
                clear_list.append(5)
        elif hour[1] == hour[2] != MyTime(0) and member_type != 'd': # one hour extra work at noon
            extra_extra += 1
        elif 1230 <= int(hour[1]) <= 1300 and member_type != 'd':
            extra_extra += min((hour[1]-time2).hour_floor,1)
        elif 1300 < int(hour[1]): # time out of range
            return False
        if int(hour[2]) == 0:
            if member_type != 'd':
                if int(hour[3]) != 0:
                    return False
                elif hour[7] == 0:
                    off[0] += 0.5
                else:
                    off[hour[7]] += 4
                clear_list.append(7)
        elif hour[1] == hour[2]:
            pass
        elif 1200 < int(hour[2]) <= 1230:
            if member_type != 'd':
                extra_extra += 0.5
        elif time3 < hour[2]:
            if hour[6] == -1 and member_type == 'b':
                pass
            elif hour[6] == 0:
                late += (hour[2]-time3).minutes
                clear_list.append(6)
            elif hour[7] == -1:
                return False
            else:
                off[hour[7]] += (hour[2]-time3).hour_ceil
                clear_list.append(7)
        if 0 < int(hour[3]) < 1700:
            if hour[7] in [-1,9]:
                return False
            else:
                off[hour[7]] += (time4-hour[3]).hour_ceil
                clear_list.append(7)
        elif 1740 <= int(hour[3]):
            if hour[3] < MyTime(1900) or member_type == 'e': #沒吃晚餐 加班從1700開始
                extra += (hour[3]-time4).hour_floor
            else:
                extra += (hour[3]-MyTime(1730)).hour_floor
        elif int(hour[3]) == 0 and member_type == 'd':
            if hour[7] == 0:
                off[0] += 0.5
            else:
                off[hour[7]] += 4
            clear_list.append(7)
        for i in clear_list:
            hour[i] = -1
    else: # 假日加班
        if MyTime(0) not in hour[:2]:
            extra += min(hour[1].hour_floor,12)-min(hour[0].hour_ceil,(hour[0]-MyTime(5)).hour_ceil)
        if MyTime(0) not in hour[2:4]:
            extra += min(hour[3].hour_floor,18)-min(hour[2].hour_ceil,(hour[2]-MyTime(5)).hour_ceil)
    if hour[4:] != [-1,-1,-1,-1]: # some input informations aren't used, need to check if error occurred.
        return False
    else:
        result = [0]
        if work_day:
            result[0] += 1-off[0]-(off[1]+off[2]+off[3])/8
        if extra > 2:
            result.append(2)
            result.append(extra-2)
        else:
            result.append(extra)
            result.append(0)
        result[-2] += extra_extra
        if late:
            result.append(late)
            result.append(1)
        else:
            result.append(0)
            result.append(0)
        for i in off:
            result.append(i)
        return result

def read_record(record_cell):
    pass
    
#def special_days(member_info, year):
#    mb_year, mb_month, mb_day = member_info[1:]
#    if member_info[2:] == [2,29]:
#        mb_day = 28
#    proportion = (date(year,12,31)-date(year,mb_month,mb_day)).days/(date(year,12,31)-date(year-1,12,31)).days
#    if year == mb_year and mb_month >= 7:
#        return 0
#    elif year >= mb_year:
#        year_delta = year - mb_year
#        if year_delta <= 8:
#            if year_delta == 0:
#                proportion -= 0.5
#            day_delta = spec_off_list[year_delta+1]-spec_off_list[year_delta]
#            return ((day_delta*proportion+spec_off_list[year_delta])*10).__round__()/10
#        elif year_delta <= 24:
#            print(member_info)
#            print(((proportion+year_delta+6)*10).__round__()/10)
#            annual_statistic[i[0]][11] += i[5]
#            annual_statistic[i[0]][12] = annual_statistic[i[0]][10]-annual_statistic[i[0]][11]
#            return ((proportion+year_delta+6)*10).__round__()/10
#        else:
#            return 30
#    else:
#        return "錯誤"

# read members and sort
member_cell = Member('員工資料.xlsx')
members = member_cell.members # members == [["a99","何宇智",2003,1,2,3], ["a98","何致泉",2007,7,2,7], ...]
member_dict = member_cell.special_days_dict # member_dict == ["a99":["何宇智",3], "a98":["何致泉",7], ...]
members.sort(key=lambda x:x[0]) # Sort members by their numbers
member_cell.to(2,1)
for i in members:
    for j in i:
        member_cell.write(j)
        member_cell.right()
    member_cell.move(1,-6)
member_cell.save()

today = date.today()
year = today.year
#spec_off_list = [0,3,7,10,14,14,15,15,15,15]
while True:
    try:
        cell = Cell(f'{year}打卡紀錄.xlsx', False)
        holiday_cell = Cell(f'{year}假日表.xlsx', False)
        # read members and sort
        current_year = (year == today.year)
        if current_year:
            member_cell = Member('員工資料.xlsx', False) # Load the latest version of members
            current_member_cell = Cell(f'{year}員工資料.xlsx', False)
            # After having loaded, update the latest version to the current year's member file
            # Delete the stored current year member for later update
            current_member_cell.to(2,1)
            while current_member_cell.value:
                for i in range(6):
                    current_member_cell.write(data=False, fill_color=False)
                    current_member_cell.right()
                current_member_cell.move(1,-6)
        else:
            member_cell = Member(f'{year}員工資料.xlsx', False) # Load the previous version of members
        members = member_cell.members # members == [["a99","何宇智",2003,1,2,3], ["a98","何致泉",2007,7,2,7], ...]
        member_dict = member_cell.special_days_dict # member_dict == ["a99":["何宇智",3], "a98":["何致泉",7], ...]
        members.sort(key=lambda x:x[0]) # Sort members by their numbers
        member_cell.to(2,1)
        if current_year:
            current_member_cell.to(2,1)
        for i,member in enumerate(members):
            color = 'grey' if i % 2 == 0 else False
            for j in member:
                member_cell.write(j, fill_color=color, align="center")
                member_cell.right()
                if current_year:
                    current_member_cell.write(j, fill_color=color, align='center')
                    current_member_cell.right()
            member_cell.move(1,-6)
            if current_year:
                current_member_cell.move(1,-6)
        member_cell.save(close=True)
        if current_year:
            current_member_cell.save(close=True)
    except:
        print(f"The file for year {year} doesn't exist.\n")
        break
    cell.down()
    annual_statistic = dict() # annual_statistic == ["a01":[0,0,0,0,0,0,0,0,0,0,3,0,3], "a02":[0,0,0,0,0,0,0,0,0,0,7,6,1], ...]
    quit_member = [] # quit_member == ["a01","a02",...]
    new_member = [] # new_member == [["a99","何宇智"], ["a98","何致泉"], ...]
    annual_member = [] # annual_member == [["a99","何宇智"], ["a98","何致泉"], ...]
    while cell.value: # collect annual member data and erase all annual statistic
        member_id = cell.value
        annual_member.append([member_id]) # worked or working member of the year
        cell.write(data=False, fill_color=False)
        cell.right()
        annual_member[-1].append(cell.value)
        annual_statistic[member_id] = []
        cell.write(data=False, fill_color=False)
        cell.right()
        for i in range(13): # erase all annual statistic (labels won't be erased)
            cell.write(data=False, fill_color=False)
            annual_statistic[member_id].append(0)
            cell.right()
        if not member_cell.find(member_id):
            if current_year: # quit the job in the current year
                quit_member.append(member_id)
            annual_statistic[member_id][10] = "離職"
        else:
            annual_statistic[member_id][10] = member_dict[member_id][1]
        cell.move(1,-15)
    for i in members: #find new members of the year
        if i[0] not in annual_statistic.keys() and int(i[2]) <= year:
            new_member.append(i[0:2])
            annual_member.append(i[0:2])
            annual_statistic[i[0]] = []
            for j in range(13):
                annual_statistic[i[0]].append(0)
            annual_statistic[i[0]][10] = member_dict[i[0]][1]
    annual_member.sort(key=lambda x:x[0])
    print(f'Processing {year} data\n')
    for month in range(1,13):
        new_month = True if today.month < month and current_year else False
        weekday, days = monthrange(year, month)
        holiday_cell.to(7*(month-1-((month-1)%4))/4+2,((month-1)%4)*8+((weekday+1)%7)+2)
        holiday = []
        work_days = days
        for day in range(days):
            if holiday_cell.value == 0:
                holiday_cell.write(fill_color="red")
                holiday.append(day)
                work_days -= 1
            else:
                holiday_cell.write(fill_color="green")
            holiday_cell.right()
            if holiday_cell.current_pos[1]%8 == 1:
                holiday_cell.move(1,-7)
        cell.to_sheet(month)
        statistic = [] # statistic == [["a99", "何宇智",0,0,0,0,0,0,0,0,0], ...]
        month_data = [] # month_data == [[["a99","何宇智"],[False,"0800","1200","1300","1700",False,False,False,False],...], ...]
        for i in new_member: # initialize empty slot for all new members of current month
            if member_dict[i[0]][1] < year or member_dict[i[0]][2] <= month:
                print(f'{month}')
                statistic.append([i[0],i[1],0,0,0,0,0,0,0,0,0])
                month_data.append([i])
                for j in range(days):
                    month_data[-1].append([False for k in range(9)])

        
#############################################################################################
        while cell.value: # read all data and duplicate them
            cell.set_position()
            if new_month and cell.value in quit_member:# member to be quit --> no need to duplicate data
                pass
            else:
                month_data.append([[cell.value]])
                statistic.append([cell.value])
                member_type = cell.value[0]
                cell.right()
                month_data[-1][-1].append(cell.value) # member name
                statistic[-1].append(cell.value)
                cell.down()
                for i in range(9):
                    statistic[-1].append(0)
                for day in range(days):
                    work_day = bool(day not in holiday)
                    month_data[-1].append([False])
                    # If an error of rapid enter occurred, cell will erase "錯誤".
                    # If other error occurred, cell will erase and put a new "錯誤" back, since the wrong hour is still on it.
                    ######## rapid enter ########
                    if cell.value and cell.value != '錯誤':
                        hour = rapid_enter(cell.value)
                        if hour == None:
                            for i in range(4):
                                month_data[-1][-1].append(False)
                        else:
                            for i in range(4):
                                month_data[-1][-1].append(hour[i].time)
                        for i in range(4):
                            month_data[-1][-1].append(False)
                    ################
                    else: 
                        hour = []
                        for i in range(4):
                            cell.right()
                            try:
                                ######## rapid hour enter ########
                                if 0 < cell.value <= 59 and i in [0,2]:
                                    if i == 0:
                                        hour.append(MyTime(int(700+cell.value)))
                                        month_data[-1][-1].append(700+cell.value)
                                    else:
                                        hour.append(MyTime(int(1200+cell.value)))
                                        month_data[-1][-1].append(1200+cell.value)
                                elif 0 <= cell.value <= 59 and i in [1,3]:
                                    if i == 1:
                                        hour.append(MyTime(int(1200+cell.value)))
                                        month_data[-1][-1].append(1200+cell.value)
                                    else:
                                        hour.append(MyTime(int(1700+cell.value)))
                                        month_data[-1][-1].append(1700+cell.value)
                                ################
                                else:
                                    hour.append(MyTime(int(cell.value)))
                                    month_data[-1][-1].append(cell.value)
                            except TypeError:
                                hour.append(MyTime(0))
                                month_data[-1][-1].append(False)
                        for i in range(4):
                            cell.right()
                            if type(cell.value) == int:
                                hour.append(cell.value)
                                if i % 2 == 0:
                                    month_data[-1][-1].append(late_name(cell.value))
                                else:
                                    month_data[-1][-1].append(day_off_name(cell.value))
                            elif cell.value == None:
                                hour.append(-1)
                                month_data[-1][-1].append(False)
                            else:
                                if i % 2 == 0:
                                    hour.append(if_late[cell.value])
                                else:
                                    hour.append(day_off[cell.value])
                                month_data[-1][-1].append(cell.value)
                        cell.left(8)
                    if hour_error(hour, work_day):
                        month_data[-1][-1][0] = '錯誤'
                    else:
                        month_data[-1][-1][0] = False
                        result = daily_result(member_type, hour, work_day)
                        if type(result) == bool:
                            month_data[-1][-1][0] = '錯誤'
                        elif type(result) == list:
                            for i in range(9):
                                statistic[-1][i+2] += result[i]
                    cell.down()
            cell.pin()
            ######## clear data in the file ########
            for i in range(days+1):
                for j in range(10):
                    cell.write(data=False, bold=False, fill_color=False)
                    cell.right()
                cell.move(1,-10)
            cell.pin()
            cell.right(10)
            if cell.value == None:
                cell.to(column=1)
                cell.down(days+2)
            ################
        ######## clear statistic ########
        cell.to(2,32)
        while cell.value:
            for i in range(13):
                cell.write(data=False, fill_color=False)
                cell.right()
            cell.move(1,-13)
        ################

############################################################################
        cell.to(1,1)
        month_data.sort(key=lambda i:i[0][0]) # sort all members according to their number
        for data in month_data: # write all data to the file
            cell.set_position()
            cell.write(data[0][0], bold=True, align="center", fill_color="bright_blue") # member number
            cell.right()
            cell.write(data[0][1], bold=True, align="center", fill_color="bright_blue") # member name
            for n,label in enumerate(labels):
                if n%2 == 0:
                    color = "yellow"
                else:
                    color = False
                cell.right()
                cell.write(label, bold=True, align="center", fill_color=color)
            cell.move(1,-9)
            for day in range(days):
                if day in holiday:
                    cell.write(f'*{day+1}', bold=True, fill_color="pink", align="center")
                elif data[day+1][0] == "錯誤":
                    cell.write(day+1, bold=True, fill_color="red", align="center")
                else:
                    cell.write(day+1, bold=True, fill_color=False, align="center")
                for block in data[day+1]:
                    cell.right()
                    cell.write(block, align="center")
                cell.move(1,-9)
            cell.pin()
            cell.right(10)
            if cell.current_pos[1] == 31:
                cell.to(column=1)
                cell.down(days+2)
        cell.to(2,32)
        statistic.sort(key=lambda x:x[0])
        for n,i in enumerate(statistic): # write month statistic to the right of current sheet
            if n%2 == 0:
                color=False
            else:
                color="bright_yellow"
            work = False
            for j,k in enumerate(i): # add month statistic to annual statistic and print the month statistic into file
                if j in [0,1]:
                    cell.write(data=k, align="center", bold=True, fill_color=color)
                else:
                    cell.write(data=k, align="center", fill_color=color)
                cell.right()
                if j >= 2:
                    annual_statistic[i[0]][j-1] += k
                    if j == 2 and k != 0:
                        work = True
            if work:
                annual_statistic[i[0]][0] += 1
            try:
                annual_statistic[i[0]][11] += i[7]
                annual_statistic[i[0]][12] = annual_statistic[i[0]][10]-annual_statistic[i[0]][11]
                if annual_statistic[i[0]][12] <= -1:
                    cell.write(annual_statistic[i[0]][12], align="center", fill_color="red")
                else:
                    cell.write(annual_statistic[i[0]][12], align="center", fill_color=color)
                    work_day_count = int(i[2]+i[7]+sum(i[8:])*0.125)
                    if work_day_count == work_days:
                        cell.right()
                        cell.write(data='完成', align="center", fill_color="green")
                        cell.left()
            except TypeError: #don't need to calculate quit member's special days
                cell.write(data="離職", align="center", fill_color="green")
            cell.move(1,-11)
        cell.save()
    cell.to_sheet(0)
    cell.to(2,1)
    for n,member in enumerate(annual_member):
        if n%2 == 0:
            color=False
        else:
            color="grey"
        cell.write(member[0], bold=True, align="center", fill_color=color)
        cell.right()
        cell.write(member[1], bold=True, align="center", fill_color=color)
        for i in annual_statistic[member[0]]:
            cell.right()
            cell.write(i, align="center", fill_color=color)
            if cell.value == "離職":
                cell.write(fill_color="green")
            elif cell.value == "錯誤":
                cell.write(fill_color="orange")
            elif cell.value <= -1:
                cell.write(fill_color="red")
        cell.move(1,-14)
    cell.save()
    holiday_cell.save()
    year -= 1
print('檔案處理成功')
sleep(2)