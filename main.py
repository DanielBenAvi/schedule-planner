from tracemalloc import start
import xlrd  # טיפול בקבצי אקסל
# from Course import Course
# from Schedual import Schedual
import PySimpleGUI as gui

class Course():
    def __init__(self, name, start, end, lecturer, score) -> None:
        self.name = name
        self.start = start
        self.end = end
        self.lecturer = lecturer
        self.score = score

    def __str__(self) -> str:
        weekday = {1:'Sun',2:'Mon',3:'Tue',4:'Wed',5:'Thu',6:'Fri'}
        return f'{weekday[int(self.start/100)]}: {int(self.start%100)}:00-{int(self.end%100)-1}:50 ->  {self.name} ,   {self.lecturer.__str__()},   score = {self.score}'

class Schedual():

    def __init__(self) -> None:
        self.courses = []
        self.schedualScore = 0
        self.valid = True

    def addCourse(self, course) -> None:
        self.courses.append(course)
        self.courses.sort(key=lambda x: x.start)  # order after adding

    def calcScore(self) -> int:
        for course in self.courses:
            self.schedualScore += course.score
        return self.schedualScore

    def isValid(self,MAX_BREAK,FREE_DAY,TOP_HOUR,MIN_HOUR_PER_DAY) -> bool:
        for index in range(0, len(self.courses)-1):
            # remove not valid
            if(self.courses[index].end > self.courses[index + 1].start):
                return False
            # remove if have breaks longer than 2 hours
            if(int(self.courses[index].start/100) == int(self.courses[index + 1].start/100)):
                if((int(self.courses[index+1].start%100) - int(self.courses[index].end%100)) > MAX_BREAK):
                    return False

        for index in range(0, len(self.courses)):
            # remove end after 20:00
            if(self.courses[index].end%100>TOP_HOUR):
                return False
            # remove fridays
            if(int(self.courses[index].start/100)>(FREE_DAY-1)):
                return False
        
        # remove days with only 3 hours
        for day in range(1,7):
            sum = 0
            for index in range(0, len(self.courses)):
                if(int(self.courses[index].start/100)==day):
                    sum = sum + int(int(self.courses[index].end%100)-int(self.courses[index].start%100))
            if(sum <= MIN_HOUR_PER_DAY and sum > 0):
                return False

        return True

    def __str__(self) -> str:
        str = '------\n'
        str += f'score = {self.schedualScore}\n'
        for course in self.courses:
            str += course.__str__() + '\n'
        return str



'''
    קבועים
'''
WEEKDAYS_WIDTH = 35
WEEKDAYS_HIGHT = 3


"""
    פונקציה ליצירת מילון של כל מרצה והציון שלו
"""
def lecturrers_to_dict(lecturer_sheet, lecturers={}):
    for row in range(1, lecturer_sheet.nrows):
        lecturers[lecturer_sheet.cell_value(row, 0)] = int(
            lecturer_sheet.cell_value(row, 1))
    return lecturers

"""
    פונקציה ליצירת רשימה של כלל הקורסים האפשריים
"""
def courses_to_list(Course, classes_sheet, lecturers, courses=[]):
    for row in range(1, classes_sheet.nrows):
        courses.append(Course(classes_sheet.cell_value(row, 0),
                              classes_sheet.cell_value(row, 1),
                              classes_sheet.cell_value(row, 2),
                              classes_sheet.cell_value(row, 3),
                              lecturers.get(classes_sheet.cell_value(row, 3))))
    return courses

"""
    אלגוריתם למיון כלל המערכות האופציונליות והורדה ע"פ תנאים מגבילים אשר נמצאים במחלקת המערכת
"""
def all_sub_groups(arr,maxBreak,freeDay,maxHour,minPerDay, schedual_list=[]):
    n = len(arr) 
    indices = [0 for i in range(n)]
    while (1):
        s = Schedual()
        for i in range(n):
            s.addCourse(arr[i][indices[i]])
        if (s.isValid(maxBreak,freeDay,maxHour,minPerDay)):
            s.schedualScore = s.calcScore()
            schedual_list.append(s)
        next = n - 1
        while (next >= 0 and (indices[next] + 1 >= len(arr[next]))):
            next -= 1
        if (next < 0):
            return schedual_list
        indices[next] += 1
        for i in range(next + 1, n):
            indices[i] = 0
    return schedual_list

"""
    מיון כל הרשימות ע"פ קורס
"""
def order_class_by_type(must_take, courses, courses_new=[]):
    for x in must_take:
        tmp = []
        for cl in courses:
            if cl.name == x:
                tmp.append(cl)
        courses_new.append(tmp)
    return courses_new

"""
    הדפסת כל הרשימות
"""
def print_all_max(schedual_list, max_attr):
    for x in schedual_list:
        if x.schedualScore == max_attr.schedualScore:
            print(x)

"""
    פונקציה ליצירת רשימה של כל הקורסים שחובה לקחת
"""
def mustTake(sheet, must=[]):
    for row in range(1,sheet.nrows):
        must.append(str(sheet.cell_value(row,0)))
    return must

def main(maxBreak,freeDay,maxHour,minPerDay,file):
    # מיקום התיקיה
    location = file 
    wb = xlrd.open_workbook(location)

    # יצירת קורסים שחייבים לקחת
    must_sheet = wb.sheet_by_index(2)
    must_take = mustTake(must_sheet)
    lecturer_sheet = wb.sheet_by_index(0)
    classes_sheet = wb.sheet_by_index(1)

    # קבלת מרצים מקובץ אקסל
    lecturers = lecturrers_to_dict(lecturer_sheet)

    # קבלת קורסים מקובץ אקסל
    courses = courses_to_list(Course, classes_sheet, lecturers)

    # יצירת רשימה של כל המערכות ע"פ רשימת תנאים
    schedual_list = all_sub_groups(order_class_by_type(must_take, courses),maxBreak,freeDay,maxHour,minPerDay)

    # מיון המערכות ע"פ ציון
    schedual_list.sort(key=lambda x: x.schedualScore, reverse=True)

    # מציאות המערכת בעלת הציון הגבוה ביותר 
    # max_attr = max(schedual_list, key=attrgetter('schedualScore'))
    # הדפסת כל המערכות בעלות הציון המקסימלי
    # print_all_max(schedual_list, max_attr)
    return(schedual_list)

def create_window(theme):
    gui.theme(theme)
    gui.set_options(font='Courier 10')

    column = [
        [gui.Text("Day: "),gui.Text("Sun",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH)),gui.Text("Mon",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH)),gui.Text("Tue",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH)),gui.Text("Wed",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH)),gui.Text("Thu",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH)),gui.Text("Fri",justification='center',expand_x=True,size=(WEEKDAYS_WIDTH))],
            [gui.Text("08:00"),gui.Text("",expand_x=True,justification='center',key='-108-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-208-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-308-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-408-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-508-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-608-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("09:00"),gui.Text("",expand_x=True,justification='center',key='-109-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-209-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-309-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-409-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-509-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-609-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("10:00"),gui.Text("",expand_x=True,justification='center',key='-110-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-210-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-310-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-410-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-510-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-610-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("11:00"),gui.Text("",expand_x=True,justification='center',key='-111-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-211-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-311-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-411-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-511-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-611-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("12:00"),gui.Text("",expand_x=True,justification='center',key='-112-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-212-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-312-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-412-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-512-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-612-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("13:00"),gui.Text("",expand_x=True,justification='center',key='-113-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-213-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-313-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-413-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-513-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-613-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("14:00"),gui.Text("",expand_x=True,justification='center',key='-114-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-214-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-314-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-414-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-514-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-614-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("15:00"),gui.Text("",expand_x=True,justification='center',key='-115-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-215-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-315-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-415-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-515-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-615-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("16:00"),gui.Text("",expand_x=True,justification='center',key='-116-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-216-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-316-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-416-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-516-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-616-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("17:00"),gui.Text("",expand_x=True,justification='center',key='-117-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-217-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-317-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-417-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-517-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-617-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("18:00"),gui.Text("",expand_x=True,justification='center',key='-118-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-218-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-318-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-418-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-518-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-618-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("19:00"),gui.Text("",expand_x=True,justification='center',key='-119-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-219-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-319-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-419-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-519-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-619-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("20:00"),gui.Text("",expand_x=True,justification='center',key='-120-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-220-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-320-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-420-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-520-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-620-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("21:00"),gui.Text("",expand_x=True,justification='center',key='-121-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-221-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-321-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-421-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-521-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-621-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("22:00"),gui.Text("",expand_x=True,justification='center',key='-122-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-222-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-322-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-422-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-522-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-622-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)],
            [gui.Text("23:00"),gui.Text("",expand_x=True,justification='center',key='-123-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-223-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-323-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-423-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-523-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2),gui.Text("",expand_x=True,justification='center',key='-623-',size=(WEEKDAYS_WIDTH,WEEKDAYS_HIGHT),relief=gui.RELIEF_SOLID,border_width=2)]]

    layout = [
            [gui.Text("max Break:",justification='center',expand_x=True),gui.Text("Free Day:",justification='center',expand_x=True),gui.Text("Max Hour:",justification='center',expand_x=True),gui.Text("Min Hour Per Day:",justification='center',expand_x=True)],
            [gui.Slider(orientation ='horizontal', range=(0,10),default_value=1,expand_x=True,key='-maxBreak-'),gui.Slider(orientation ='horizontal', range=(1,7),default_value=6,expand_x=True,key='-freeDay-'),gui.Slider(orientation ='horizontal', range=(8,23),default_value=20,expand_x=True,key='-maxHour-'),gui.Slider(orientation ='horizontal', range=(1,10),default_value=3,expand_x=True,key='-minPerDay-')],
            [gui.Button("<<",expand_x=True,key='-<-'),gui.Button("Start",expand_x=True,key='-start-'),gui.Button("Load",expand_x=True,key='-load-'),gui.Button(">>",expand_x=True,key='->-')],
            [gui.Text("Calendar",justification='center',expand_x=True,pad=(5,10),key='-score-',font='Ariel 20'),gui.Text("Choose a file: "), gui.FileBrowse(key='-file-'),gui.Button('Submit',key='-submit-')],
            [gui.Column(column, scrollable=True, vertical_scroll_only=True,expand_y=True,expand_x=True)]]
    return gui.Window(program_name,layout,resizable=True,icon='icon.ico')



def loadBoard(window, schedual,nummber):
    window['-score-'].update(f'schedual nummber = {nummber+1}, score = {schedual.schedualScore}')
    for course in schedual.courses:
        for i in range(int(course.start),int(course.end)):
            tmp = str(course.name)+'\n'+str(course.lecturer)
            window[f'-{i}-'].update(tmp)


def reset_board(window):
    for day in range(1,7):
        window['-score-'].update('')
        for hour in range(8,24):
            window[f'-{day}{hour if hour > 9 else str(0)+str(hour)}-'].update('')

if __name__ == "__main__":
    current_schedual = 0
    program_name = 'Calendar'
    theme_engine = ['menu',[]]
    window = create_window('Black').Finalize()
    window.Maximize()
    
    window['-load-'].update(disabled=True)
    window['-<-'].update(disabled=True)
    window['->-'].update(disabled=True)
    window['-start-'].update(disabled=True)

    file = ''

    while True:
        event, values = window.read()

        if event == gui.WIN_CLOSED:
            break
        
        if event in theme_engine[1]:
            window.close()
            window = create_window(event)
        if event in ['-submit-']:
            window['-start-'].update(disabled=False)
            file = values['-file-']

        if event in ['-start-']:
            schedual_list = main(values['-maxBreak-'],values['-freeDay-'],values['-maxHour-'],values['-minPerDay-'],file)
            window['-load-'].update(disabled=False)

        # איפוס הלוח
        if event in ['-load-']:
            loadBoard(window, schedual_list[current_schedual],current_schedual)
            window['-load-'].update(disabled=True)
            window['-<-'].update(disabled=False)
            window['->-'].update(disabled=False)

        if event in ['-<-']:
            if(current_schedual > 0):
                current_schedual-=1
                reset_board(window)
                loadBoard(window, schedual_list[current_schedual],current_schedual)

        if event in ['->-']:
            current_schedual+=1
            reset_board(window)
            loadBoard(window, schedual_list[current_schedual],current_schedual)
        
    window.close()

