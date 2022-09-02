import xlrd  # טיפול בקבצי אקסל
from operator import attrgetter # מציאת אובייקט מקסלימלי באמצעות תכונה מסויימת
from Course import Course
from Schedual import Schedual

"""
    יצירת הגבלות
"""
MAX_BREAK = 1
FREE_DAY = 6
TOP_HOUR = 20
MIN_HOUR_PER_DAY = 3

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
def all_sub_groups(arr, schedual_list=[]):
    n = len(arr) 
    indices = [0 for i in range(n)]
    while (1):
        s = Schedual()
        for i in range(n):
            s.addCourse(arr[i][indices[i]])
        if (s.isValid(MAX_BREAK,FREE_DAY,TOP_HOUR,MIN_HOUR_PER_DAY)):
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

def main():
    # מיקום התיקיה
    location = "C:/Users/danie/documents/Python/calendar_project/DataBase.xls"
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
    schedual_list = all_sub_groups(order_class_by_type(must_take, courses))

    # מיון המערכות ע"פ ציון
    schedual_list.sort(key=lambda x: x.schedualScore, reverse=True)

    # מציאות המערכת בעלת הציון הגבוה ביותר 
    max_attr = max(schedual_list, key=attrgetter('schedualScore'))

    # הדפסת כל המערכות בעלות הציון המקסימלי
    print_all_max(schedual_list, max_attr)

if __name__ == "__main__":
    main()

