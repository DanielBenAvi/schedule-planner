import xlrd  # excel file handler
from operator import attrgetter # max element of list by attribute

from Course import Course
from Schedual import Schedual



def lecturrers_to_dict(lecturer_sheet, lecturers={}):
    for row in range(1, lecturer_sheet.nrows):
        lecturers[lecturer_sheet.cell_value(row, 0)] = int(
            lecturer_sheet.cell_value(row, 1))
    return lecturers


def courses_to_list(Course, classes_sheet, lecturers, courses=[]):
    for row in range(1, classes_sheet.nrows):
        courses.append(Course(classes_sheet.cell_value(row, 0),
                              classes_sheet.cell_value(row, 1),
                              classes_sheet.cell_value(row, 2),
                              classes_sheet.cell_value(row, 3),
                              lecturers.get(classes_sheet.cell_value(row, 3))))
    return courses


def all_sub_groups(arr, schedual_list=[]):
    n = len(arr)
    indices = [0 for i in range(n)]
    while (1):
        s = Schedual()
        for i in range(n):
            s.addCourse(arr[i][indices[i]])
        if (s.isValid()):
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


def order_class_by_type(must_take, courses, courses_new=[]):
    for x in must_take:
        tmp = []
        for cl in courses:
            if cl.name == x:
                tmp.append(cl)
        courses_new.append(tmp)
    return courses_new


def print_all_max(schedual_list, max_attr):
    for x in schedual_list:
        if x.schedualScore == max_attr.schedualScore:
            print(x)


def main():
    location = "C:/Users/danie/documents/Python/calendar_project/DataBase.xls"
    wb = xlrd.open_workbook(location)
    must_take = ['Numeric LE', 'Numeric P', 'Algorithem LE', 'Algorithem P','Methods in software engineering LE','Methods in software engineering P','human computer interfaces LE','human computer interfaces P','Data Security LE','Data Security P','computer embedded systems LE','computer embedded systems LA','Mobile applications LE','Mobile applications P']
    lecturer_sheet = wb.sheet_by_index(0)
    classes_sheet = wb.sheet_by_index(1)

    '''
    get lecturres from excel file
    '''
    lecturers = lecturrers_to_dict(lecturer_sheet)

    '''
    get courses from excel file
    '''
    courses = courses_to_list(Course, classes_sheet, lecturers)

    '''
    create list of all valid schedual
    '''
    schedual_list = all_sub_groups(order_class_by_type(must_take, courses))

    '''
    sort the list of all valid schedual by score
        '''
    schedual_list.sort(key=lambda x: x.schedualScore, reverse=True)

    '''
    find max score
    '''
    max_attr = max(schedual_list, key=attrgetter('schedualScore'))

    '''
    print all max score
    '''
    print_all_max(schedual_list, max_attr)

if __name__ == "__main__":
    main()

