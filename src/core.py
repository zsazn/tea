#! /usr/bin/env python
#-*- encoding = UTF-8 -*-
from __future__ import division
__author__ = 'Mark Z. Zhou'
__version__ = '1.0'
import os
import xlrd # using xlrd version 0.9.3
from xlwt import Workbook # using xlwt version 0.7.5

os.system('title, Teaching Evaluation Report Generator')
os.system("mode con cols=120 lines=40")

print '                                  _________________________________________________ '
print '                                 |  _____________________________________________  |'
print '                                 | |                                             | |'
print '                                 | |     Teaching Evaluation Report Generator    | |'
print '                                 | |                                             | |'
print '                                 | |                 Version: 1.0                | |'
print '                                 | |                                             | |'
print '                                 | |             All Rights Reserved             | |'
print '                                 | |            College of Technology            | |'
print '                                 | |              Purdue University              | |'
print '                                 | |                                             | |'
print '                                 | |           Updated on Apr. 9, 2014           | |'
print '                                 | |_____________________________________________| |'
print '                                 |_________________________________________________|'

rawdata = xlrd.open_workbook("data_rev.xls") # Read the raw data from file

crndata = xlrd.open_workbook("crn_rev.xls") # Read the CRN related data from file

n = rawdata.nsheets # Number of the sheets in the workbook

shnum = range(n) # A list of index. the length of the list equals to the number of the sheets

sh = [] # A list that stores all the sheets in data workbook

shrows = [] # A list that stores the numbers of total rows in each sheet in data workbook

shcols = [] # A list that stores the numbers of total columns in each sheet data workbook

crs_all = [] # crs_all is a list that stores all course numbers

# Get the sheets/rows/columns lists of the raw data
for num in range(n):
    sh.append(rawdata.sheet_by_index(num))
    shrows.append(rawdata.sheet_by_index(num).nrows)
    shcols.append(rawdata.sheet_by_index(num).ncols)

# Get the first sheet as well as the rows and columns in crn workbook
sh_crn = crndata.sheet_by_index(0) # An object refers to the first sheet in crn workbook (with section info)
shrows_crn = crndata.sheet_by_index(0).nrows # A list that stores the total number of rows in the first sheet of crn workbook
shcols_crn = crndata.sheet_by_index(0).ncols # A list that stores the total number of columns in the first sheet of crn workbook

#sh_crn_1 = crndata.sheet_by_index(1) # An object refers to the second sheet in crn workbook (without section info)
#shrows_crn_1 = crndata.sheet_by_index(1).nrows # A list that stores the total number of rows in the second sheet of crn workbook
#shcols_crn_1 = crndata.sheet_by_index(1).ncols # A list that stores the total number of columns in the second sheet of crn workbook

for i in range(1, shrows_crn):
    crs_all.append(str(sh_crn.cell_value(i, 2)))

# Get and validate the user's query of name ([last name, first name])

def getQueryName_last_name(): # get and validate the user's input of last name
    queryName_tmp = raw_input("\n1. Which instructor would you like to query?\n   Please input the LAST name of this instructor:\n").title() # Input the query by faculty names
    faculty = dict()
    for num in shnum:
        for i in range(1, shrows[num]):
            key = sh[num].cell_value(i, 5)
            value = sh[num].cell_value(i, 6)
            if key not in faculty.keys():
                faculty[key] = [value]
            else:
                faculty[key].append(value)
    for key, value in faculty.iteritems():
        value = list(set(value))
        faculty[key] = value
                    
    if queryName_tmp in faculty.keys():
#       print queryName_tmp
        return queryName_tmp, faculty[queryName_tmp]        
    else:
        print "\n   The instructor " + queryName_tmp + " does not exist. Please check your input.\n"
        return getQueryName_last_name()       

queryName_res = getQueryName_last_name() # get a querying name dictionary based on input last name {last name: [first name 1, first name 2, ...]}

queryName_last = queryName_res[0] # not sharing then get the last name

def getQueryName_first_name(namelist): # get and validate user's input of first name
    queryName_first_input = raw_input("\nPlease input the FIRST name of the instructor you'd like to querying:\n").title()     
    if queryName_first_input in queryName_res[1]:
        queryName_first_tmp = queryName_first_input
        return queryName_first_tmp
    else:
        print "\nIt seems a typo... Please check your input.\n"
        return getQueryName_first_name(queryName_res)
    
# multiple instructors are sharing a last name
if len(queryName_res[1]) > 1:
    print "\nWe have found", len(queryName_res[1]), "instructors' last name is", queryName_last + ":\n" # sharing then list the names
    for i in range(len(queryName_res[1])):
        print str(queryName_res[1][i]), queryName_res[0], "\n"
    queryName_first = getQueryName_first_name(queryName_res[1]) # get the first name
    queryName_valid = [queryName_last, queryName_first] # get the name [last name, first name] 
else:
    queryName_first = queryName_res[1][0]
    queryName_valid = [queryName_last, queryName_first] # get the name [last name, first name]

queryName_full = queryName_valid[1] + ' ' + queryName_valid[0] # set the name in the form of 'First Last' for displaying prints
            
# Get all indices of a certain querying string in a list
def indices(lst, element):
    result = []
    offset = -1
    while True:
        try:
            offset = lst.index(element, offset + 1)
        except ValueError:
            return result
        result.append(offset + 1)

# Convert semester text, e.g. from "SU10" to "summer 2010"
def convertSemester(semester):
    getConversion = ""
    if semester != "":
        if semester[0].lower() == 'f':
            getConversion = semester.replace('A', 'all 20')
        elif semester[0].lower() == 's' and semester[1].lower() == 'p':
            getConversion = semester.replace('P', 'pring 20')
        elif semester[0].lower() == 's' and semester[1].lower() == 'u':
            getConversion = semester.replace('U', 'ummer 20')
        else:
            getConversion = ""
    return getConversion

# Get the row indices of a querying name (queryName_valid) in the (num)th list
def query(queryName_valid, num):
    faculty = list()
    checkRow = list()
    for i in range(1, shrows[num]):
        faculty.append(sh[num].cell_value(i, 5)) # insert all faculties' last names in a list "faculty"
    checkRow_last = indices(faculty, queryName_valid[0]) # get the rows where the querying faculty's last name are
    
    for j in checkRow_last:
        if sh[num].cell_value(j, 6) == queryName_valid[1]: # get the rows where the querying faculty's first name are from the rows of his/her last name
            checkRow.append(j)        
    return checkRow

# Get the semester from the first column of a single sheet
# def checkSemester(num):
#     checkRow = query(queryName_valid, num)  
    # detect if the querying faculty name is in this sheet
#     if len(checkRow) == 0:  
#         return sh[num].name  # querying faculty is not found 
#     else:
#         semester = sh[num].cell_value(1, 0)  #querying faculty is in the list
#     return semester # return the semester

# Get all semesters when the querying instructor taught in
def getAvailableSemester(queryName_valid):
    semester = list()
    for num in shnum:
        for i in range(1, shrows[num]):
            if sh[num].cell_value(i, 5) == queryName_valid[0] and sh[num].cell_value(i, 6) == queryName_valid[1]:
                semester.append(sh[num].cell_value(1, 0))
                break
        continue
#            else:
#                print "\nThis instructor has not taught in " + convertSemester(sh[num].cell_value(1, 0)) + "."
    return semester

semestertmp = getAvailableSemester(queryName_valid) # a list stores all semesters the querying instructor has taught

print "\n"
print "\nThis instructor has taught in:\n"

for item in semestertmp:
    print "   " + convertSemester(item) + "\n" # convert print output from SU10 to summer 2010

# Validate the format of the semester input
def validateSemester():
    semester = raw_input("\n2. So which semester of courses would you like to query?\n   Please input the semester:\n").lower()
    asemester = semester[0].upper() + semester[1].upper() + semester[-2] + semester[-1] # reverse conversion
    vali = list()
    for num in shnum:
        vali.append(convertSemester(sh[num].cell_value(1, 0)).lower())
    if semester in vali:
        return [asemester, vali.index(semester)]
    else:
        print "\nIt seems a typo... Please check your input.\n"
        return validateSemester()
asemester = validateSemester() # format semester as [semester abbreviation, numth sheet] : e.g. ['SP13', 12] 
# print type(asemester[1])

for item in semestertmp:
    if asemester[0] == item:
        print "\n3. Alright, the courses which", queryName_full, "has taught in ", convertSemester(item), " are: \n"
        break
theSemester = convertSemester(asemester[0]) # format semester: e.g. Spring 2013

# Get the course list based on queryName_valid and semester; queryName_valid is global
def getAvailableCourse(semester):
    i = 0
    course = []
    # locate a semester (locate a sheet)
    for item in sh:
        if asemester[0] == item.cell_value(1, 0):
            i = sh.index(item)
            break
    # find all courses
    for j in range(1, shrows[i]):
        if sh[i].cell_value(j, 5) == queryName_valid[0] and sh[i].cell_value(j, 6) == queryName_valid[1]:
            course.append(sh[i].cell_value(j, 1))
    return set(course) # remove the repeated elements
crstmp = getAvailableCourse(asemester[0])
for item in crstmp:
    print str(item[:-2]) + "\n"

# Validate input
def validateCourse():
    course_tmp = raw_input("\n4. Which course would you like to look at?\n   Please choose one from the above to input: \n").upper()
    course = course_tmp
    course1 = course + "00" # e.g. BCM38500
    course2 = course + "01" # e.g. BCM38501
    if course1 in crstmp:
        return course1
    elif course2 in crstmp:
        return course2
    else:
        print "\nIt seems a typo... Please check your input.\n"
        return validateCourse()
courseNum = validateCourse()
# print courseNum

# Get the department name
def checkDept(queryName_valid, courseNum):
    num = asemester[1]    
    for i in range(1, shrows[num]):
        if sh[num].cell_value(i, 5) == queryName_valid[0] and sh[num].cell_value(i, 6) == queryName_valid[1]:
            if sh[num].cell_value(i, 1) == courseNum:
                return sh[num].cell_value(i, 3)
                break
            
dept = checkDept(queryName_valid, courseNum)
#print dept

# Get the course list in the [dept] department
def deptCourseList(dept):
    num = asemester[1]
    courseList_tmp = list()
    for i in range(1, shrows[num]):
        if sh[num].cell_value(i, 3) == dept:
            courseList_tmp.append(sh[num].cell_value(i, 1))
    courseList = list(set(courseList_tmp))
    return courseList

deptCrsLst = deptCourseList(dept)

## Get course-section pairs for the [dept] department
#def getDeptCrsPair(courseList):
#    pair = dict()
#    for i in range(1, shrows[asemester[1]]):
#        for item in courseList:              
#            if item == sh[asemester[1]].cell_value(i, 1):
#                if item not in pair.keys():
#                    pair[item] = [sh[asemester[1]].cell_value(i, 2)]
#                else:
#                    pair[item].append(sh[asemester[1]].cell_value(i, 2))
#    for key, value in pair.iteritems():
#        value = list(set(value))
#        pair[key] = value
#    return pair
#deptCrsPair = getDeptCrsPair(deptCrsLst)
##print deptCrsPair

#print deptCrsLst

## get department name for course number for print ouput
#if courseNum[:-5] == 'TECH':
#    dept1 = 'CTA'
#elif courseNum[:-5] == 'CNIT':
#    dept1 = 'C&IT'
#elif courseNum[:-5] == 'ASM':
#    dept1 = 'ABE'
#else:
#    dept1 = courseNum[:-5]


# rspnsNumOneCrs = list()

# Get instructor core for one querying course in the querying semester
def i(courseNum, queryName_valid):
    global rspnsNumOneCrs 
    rspnsNumOneCrs = list() # get response numbers of this course
    i = 0
    counter = 0
    for j in range(1, shrows[asemester[1]]):
        if courseNum == str(sh[asemester[1]].cell_value(j, 1)) and queryName_valid[0] == str(sh[asemester[1]].cell_value(j, 5)) and queryName_valid[1] == str(sh[asemester[1]].cell_value(j, 6)):
            counter += 1
            rspnsNumOneCrs.append(counter)
            if courseNum == '':
                i += 0
            else:
                i += sh[asemester[1]].cell_value(j, 13)
#    print counter
    return round((i / counter), 2)
iout = i(courseNum, queryName_valid) # the instructor core for the querying course in the querying semester
print "\nThe full report for", queryName_full, "is generating. Please press ENTER to review a brief one.\n", raw_input()
print "\n1. The Instructor Core of", queryName_full, "in", theSemester, "is", str(iout) + ".\n"
#print type(rspnsNumOneCrs[-1])

# Get course core for the querying semester
# args ensures the one-to-one course-instructor correspondence
def c(courseNum, queryName_valid):
    c = 0
    counter = 0
    for j in range(1, shrows[asemester[1]]):
        if courseNum == str(sh[asemester[1]].cell_value(j, 1)) and queryName_valid[0] == str(sh[asemester[1]].cell_value(j, 5)) and queryName_valid[1] == str(sh[asemester[1]].cell_value(j, 6)):
            counter += 1
            if courseNum == '':
                c += 0
            else:
                c += sh[asemester[1]].cell_value(j, 12)
    return round((c / counter), 2)          
cout = c(courseNum, queryName_valid) # the course core in the querying semester

print "2. The Course Core of", courseNum[:-2], "that taught by", queryName_full, "in", theSemester, "is", str(cout) + ".\n"
print "\nPlease press ENTER to continue...\n", raw_input() # this line just add a user control to the command line output

# Get the college-average INSTRUCTOR core for each semester
# caic --> College Average Instructor Core
def caic(num):
    caic = 0
    for i in range(1, shrows[num]):
        if sh[num].cell_value(i, 13) == '':
            caic += 0
        else:
            caic += sh[num].cell_value(i, 13)
    return round((caic / (shrows[num] - 1)), 2)

# Get the college-average COURSE core for each semester
# cacc --> College Average Course Core
def cacc(num):
    cacc = 0
    for i in range(1, shrows[num]):
        if sh[num].cell_value(i, 12) == '':
            cacc += 0
        else:
            cacc += sh[num].cell_value(i, 12)
    return round((cacc / (shrows[num] - 1)), 2)

# Get the department-average INSTRUCTOR core for each semester
# daic --> Department Average Instructor Core
def daic(dept, num):
    daic = 0
    depts = dict()
    checkRow = list()
    for i in range(1, shrows[num]):
        depts[i] = sh[num].cell_value(i, 3)
    for key, value in depts.iteritems():
        if value == dept:
            checkRow.append(key)
    for j in checkRow:
        if sh[num].cell_value(j, 13) == '':
            daic += 0
        else:
            daic += sh[num].cell_value(j, 13)
    if len(checkRow) == 0:
        return 0
    else:
        return round((daic / len(checkRow)), 2)

# Get the department-average COURSE core for each semester
# dacc --> Department Average Course Core
def dacc(dept, num):
    dacc = 0
    depts = dict()
    checkRow = list()
    for i in range(1, shrows[num]):
        depts[i] = sh[num].cell_value(i, 3)
    for key, value in depts.iteritems():
        if value == dept:
            checkRow.append(key)
    for j in checkRow:
        if sh[num].cell_value(j, 12) == '':
            dacc += 0
        else:
            dacc += sh[num].cell_value(j, 12)
    if len(checkRow) == 0:
        return 0
    else:
        return round((dacc / len(checkRow)), 2)

# Get average core of an instructor in one querying course for each semester
# ic --> Instructor Core
def ic(courseNum, queryName_valid, num):
    c2 = 0
    counter = 0
    for i in query(queryName_valid, num):
        if courseNum[:-2] == str(sh[num].cell_value(i, 1))[:-2]:
            if sh[num].cell_value(i, 13) == '':
                c2 += 0
            else:
                c2 += sh[num].cell_value(i, 13)
                counter += 1
    if counter == 0:
        return 0
    else:
        return round((c2 / counter), 2)

# Get average core of one querying course that taught by querying instructor for each semester
# cc --> Course Core
def cc(courseNum, queryName_valid, num):
    c1 = 0
    counter = 0
    for i in query(queryName_valid, num):
        if courseNum[:-2] == str(sh[num].cell_value(i, 1))[:-2]:
            if sh[num].cell_value(i, 12) == '':
                c1 += 0
            else:
                c1 += sh[num].cell_value(i, 12)
                counter += 1
    if counter == 0:
        return 0
    else:
        return round((c1 / counter), 2)

## Get enroll number for the querying course in the querying semester
#def getEnrollNum(courseNum, semester):
#    semesters = dict()
#    # pair semester (e.g. Fall 2009) and its column number in crn sheet
#    for i in range(3, shcols_crn):
#        semesters[str(sh_crn.cell_value(3, i))] = i
#    count = 0
#    crsrow = [] # a list stores all the rows for querying course number
#    for j in range(5, shrows_crn):
#        if courseNum[:-2] == str(sh_crn.cell_value(j, 2))[:-2]: # if querying course number equals cell value
#            crsrow.append(j)
#    # print crsrow
#    for k in crsrow: # traverse all the rows for this course
#        if sh_crn.cell_value(k, semesters[semester]) != '' and sh_crn.cell_value(k, semesters[semester]) != 0: # find the record 
#            count += int(sh_crn.cell_value(k, semesters[semester])) 
#    return count

# Get section number(s) based on course number and querying name
def getSection(courseNum, queryName_valid):
    section_tmp = list()
    for i in range(1, shrows[asemester[1]]):
        if courseNum == str(sh[asemester[1]].cell_value(i, 1)) and queryName_valid[0] == str(sh[asemester[1]].cell_value(i, 5)) and queryName_valid[1] == str(sh[asemester[1]].cell_value(i, 6)):
            section_tmp.append(sh[asemester[1]].cell_value(i, 2))
    section = list(set(section_tmp))
    section.sort()
    return section

section_list = getSection(courseNum, queryName_valid)
# print section_list

# Get enroll number for the querying course in the querying semester
def getEnrollNum(courseNum, semester):        
    def getRowsinCrn(courseNum, section_list):
        checkRow = list()
        getRow = list()
        for item in crs_all:
            if item == courseNum:
                checkRow = indices(crs_all, courseNum) # locate the rows of querying course in the first sheet of crn workbook
        for i in checkRow:
            if isinstance(sh_crn.cell_value(i, 4), float): # in case 'total for instructional type'
                for item in section_list:
                    if item == sh_crn.cell_value(i, 4): # locate section
                        #checkCrn.append(int(sh_crn.cell_value(i, 5)))
                        getRow.append(i) # get the list of rows
            elif isinstance(sh_crn.cell_value(i, 4), unicode) and len(sh_crn.cell_value(i, 4)) < 4: # in case the section number in letters
                for item in section_list:
                    if item == sh_crn.cell_value(i, 4):
                        #checkCrn.append(int(sh_crn.cell_value(i, 5)))
                        getRow.append(i) # get the list of rows
        return getRow
    rows = getRowsinCrn(courseNum, section_list)
#    print rows
    enrollNum = 0
    for i in range(6, shcols_crn):
        if semester == str(sh_crn.cell_value(0, i)): # locate the semester
            for item in rows:
                if sh_crn.cell_value(item, i) != '' and sh_crn.cell_value(item, i) != 0:
                    enrollNum += int(sh_crn.cell_value(item, i)) # add up the enroll numbers of each course each section
            break
    return enrollNum
enrollNumber = getEnrollNum(courseNum, convertSemester(asemester[0])) # enroll number of the querying course in the querying semester
if enrollNumber == 0: # in case there's no record in crn_rev.xls
    enrollNumber = rspnsNumOneCrs[-1]
    #print str(rspsrate)+"%"
else:
    rspsrate = round((rspnsNumOneCrs[-1] / enrollNumber), 4)
    
print "3. ", enrollNumber, "students has enrolled in", courseNum[:-2], "taught by", queryName_full, "in", convertSemester(asemester[0]), ".\n"
print "   ", rspnsNumOneCrs[-1], "responses in", courseNum[:-2], "taught by", queryName_full, "in", theSemester, "have been received.\n"
print "    The response rate is", str(rspsrate * 100) + "%.\n"

print "\nPlease press ENTER to continue...\n", raw_input()

# Get level of a course from a course list (result: 100-200, 300-400, 500-600)
def checkCrsLvl(courseNumDict):
    lvl_12 = dict()
    lvl_34 = dict()
    lvl_56 = dict()
    exc = dict()
    courseLevel = {'100-200': lvl_12, '300-400': lvl_34, '500-600': lvl_56, 'exception': exc}
    for key, value in courseNumDict.iteritems():
        str_value = str(value)
        if str_value[-5] == '1' or str_value[-5] == '2':
            lvl_12[key] = str_value
        elif str_value[-5] == '3' or str_value[-5] == '4':
            lvl_34[key] = str_value
        elif str_value[-5] == '5' or str_value[-5] == '6':
            lvl_56[key] = str_value
        else:
            exc[key] = str_value # e.g. BCM90000 what a weirdo....
    return courseLevel
# print checkCrsLvl(crstmp)

# Get the response number, enroll number, caic, cacc, daic, dacc for each level of the [dept] department and the college
def checkCrsLvl_all(semester):
    i = 0
    # locate semester (locate a sheet)
    for item in sh:
        if asemester[0] == item.cell_value(1, 0):
            i = sh.index(item) #semester index
            break
    # response number check
    course_rsps_c = dict()
    course_rsps_dept = dict()
    counter_rsps_c = list()
    counter_rsps_dept = list()
    c5_r = d5_r = 0 # c5: total college response number; d5: total departmental response number
    for j in range(1, shrows[i]): # for college
        course_rsps_c[j] = sh[i].cell_value(j, 1) # insert all courses of the college
        c5_r += 1
        if sh[i].cell_value(j, 3) == dept: # for department
            d5_r += 1
            course_rsps_dept[j] = sh[i].cell_value(j, 1) # insert courses only in dept department
    result_rsps_c = checkCrsLvl(course_rsps_c) # check course level
    # get response number of each level in college(response number equals total rows)  
    c1_r = len(result_rsps_c['100-200']) # 100-200 level college response number
    c2_r = len(result_rsps_c['300-400'])
    c3_r = len(result_rsps_c['500-600'])
    c4_r = len(result_rsps_c['exception'])
    counter_rsps_c.append(c1_r)
    counter_rsps_c.append(c2_r)
    counter_rsps_c.append(c3_r)
    counter_rsps_c.append(c5_r)
    counter_rsps_c.append(c4_r)
#    print counter_rsps_c, '\n'     
#    # the following 3 lines exist only because of the exception, BCM90000
#    i1 = counter_rsps_c.index(c4_r)
#    i2 = counter_rsps_c.index(c5_r)
#    counter_rsps_c[i1], counter_rsps_c[i2] = counter_rsps_c[i2], counter_rsps_c[i1] # switch the index of c4 and c5 for later use
    result_rsps_dept = checkCrsLvl(course_rsps_dept) # check course level
    # get response number of each level in department (response number equals total rows)
    d1_r = len(result_rsps_dept['100-200']) # 100-200 level departmental response number
    d2_r = len(result_rsps_dept['300-400'])
    d3_r = len(result_rsps_dept['500-600'])
    d4_r = len(result_rsps_dept['exception'])
    counter_rsps_dept.append(d1_r)
    counter_rsps_dept.append(d2_r)
    counter_rsps_dept.append(d3_r)
    counter_rsps_dept.append(d5_r)
    counter_rsps_dept.append(d4_r)
#    print counter_rsps_dept
#    j1 = counter_rsps_dept.index(d4_r)
#    j2 = counter_rsps_dept.index(d5_r)
#    counter_rsps_dept[j1], counter_rsps_dept[j2] = counter_rsps_dept[j2], counter_rsps_dept[j1]
#    print counter_rsps_dept
    
    # enroll number check
    semesterDict = dict()
    course_enroll_c = dict()
    course_enroll_dept = dict()
    counter_enroll_c = list()
    counter_enroll_dept = list()
    c1_e = c2_e = c3_e = c4_e = c5_e = 0 # c5_e: total college enroll number; c1_e: 100-200 level college enroll number, etc...
    d1_e = d2_e = d3_e = d4_e = 0 # d1_e: 100-200 level department enroll number, etc...  
    for k in range(6, shcols_crn):
        semesterDict[k] = sh_crn.cell_value(0, k)
#    print semesterDict
    tmp = convertSemester(semester)
    
    for key, value in semesterDict.iteritems():
        if tmp == value: # locate semester
#            print semesterDict.get(key)
            i_ = key # get column number
    
    # get total enroll number of college (c5_e), and total enroll number of department (d5_e)
    # get the dictionaries {row: course} for whole college and the department (these rows show the enroll number)        
    for m in range(1, shrows_crn):
        if sh_crn.cell_value(m, i_) != '': 
            # the following condition is to exclude the "total for instructional type" rows            
            if isinstance(sh_crn.cell_value(m, 4), float): # in case 'total for instructional type'
                c5_e += sh_crn.cell_value(m, i_)
                course_enroll_c[m] = sh_crn.cell_value(m, 2)
#                for key, value in deptCrsPair.iteritems():
#                    for item in deptCrsPair.values():
#                        for section in item:
                for item in deptCrsLst:
                    if sh_crn.cell_value(m, 2)[:-3] == item[:-3]:  # and sh_crn.cell_value(m, 4) == section:
#                        d5_e += sh_crn.cell_value(m, i_)
                        course_enroll_dept[m] = sh_crn.cell_value(m, 2)
            elif isinstance(sh_crn.cell_value(m, 4), unicode) and len(sh_crn.cell_value(m, 4)) < 4: # in case the section number in letters
                c5_e += sh_crn.cell_value(m, i_)
                course_enroll_c[m] = sh_crn.cell_value(m, 2)
#                for key, value in deptCrsPair.iteritems():
#                    for item in deptCrsPair.values():
#                        for section in item:
                for item in deptCrsLst:            
                    if sh_crn.cell_value(m, 2)[:-3] == item[:-3]: #and sh_crn.cell_value(m, 4) == section:
#                        d5_e += sh_crn.cell_value(m, i_)
                        course_enroll_dept[m] = sh_crn.cell_value(m, 2)
    
#    for item in deptCrsLst:
#        for n in range(1, shrows_crn):
#            if sh_crn.cell_value(n, i_) != '':
#                if isinstance(sh_crn.cell_value(n, 4), float):
#                    if sh_crn.cell_value(n, 2)[:-3] == item[:-3]:
#                        d5_e += sh_crn.cell_value(n, i_)
#                        course_enroll_dept[n] = sh_crn.cell_value(n, 2)
#                elif isinstance(sh_crn.cell_value(n, 4), unicode) and len(sh_crn.cell_value(n, 4)) < 4:
#                    if sh_crn.cell_value(n, 2)[:-3] == item[:-3]:
#                        d5_e += sh_crn.cell_value(n, i_)
#                        course_enroll_dept[n] = sh_crn.cell_value(n, 2)
   
    # get result of course level devision for college: {course level: {row, course number}} 
    result_enroll_c = checkCrsLvl(course_enroll_c)
    for key, value in result_enroll_c.iteritems(): 
        for key_ in value.keys():
            if key == '100-200':
                c1_e += sh_crn.cell_value(key_, i_)
            elif key == '300-400':
                c2_e += sh_crn.cell_value(key_, i_)
            elif key == '500-600':
                c3_e += sh_crn.cell_value(key_, i_)
            elif key == 'exception':
                c4_e += sh_crn.cell_value(key_, i_)
    counter_enroll_c.append(int(c1_e))
    counter_enroll_c.append(int(c2_e))
    counter_enroll_c.append(int(c3_e))
    counter_enroll_c.append(int(c5_e))
    counter_enroll_c.append(int(c4_e))
    
    # get result of course level devision for the department: {course level: {row, course number}}
    result_enroll_dept = checkCrsLvl(course_enroll_dept)
    for keyd, valued in result_enroll_dept.iteritems():
        for key_d in valued.keys():
            if keyd == '100-200':
                d1_e += sh_crn.cell_value(key_d, i_)
            elif keyd == '300-400':
                d2_e += sh_crn.cell_value(key_d, i_)
            elif keyd == '500-600':
                d3_e += sh_crn.cell_value(key_d, i_)
            elif keyd == 'exception':
                d4_e += sh_crn.cell_value(key_d, i_)
    d5_e = d1_e + d2_e + d3_e + d4_e # d5_e: total department enroll number; 
    counter_enroll_dept.append(int(d1_e))
    counter_enroll_dept.append(int(d2_e))
    counter_enroll_dept.append(int(d3_e))
    counter_enroll_dept.append(int(d5_e))
    counter_enroll_dept.append(int(d4_e))
    
    print "4. ", counter_enroll_c[3], "students has enrolled in the college in", theSemester + ".\n"
    print "    The enrollment numbers of each level courses in the college are:", counter_enroll_c[0:3], ", respectively.\n"
    print "\nPlease press ENTER to continue...\n", raw_input() 
    print "5. At the college level", counter_rsps_c[3], "responses were received.\n"
    print "   The response numbers of each level courses in the college are:", counter_rsps_c[0:3], ", respectively.\n" 
    print "\nPlease press ENTER to continue...\n", raw_input()
    print "6. ", counter_enroll_dept[3], "students has enrolled in the courses offered by", dept, "department.\n"
    print "    The enrollment numbers of each level", dept, "courses are:", counter_enroll_dept[0:3], ", respectively.\n"
    print "\nPlease press ENTER to continue...\n", raw_input()
    print "7. For the courses offered by", dept, "department", counter_rsps_dept[3], "responses were received.\n"
    print "   The response numbers of each level courses in", dept, "department are:", counter_rsps_dept[0:3], ", respectively.\n"
    print "\nPlease press ENTER to continue...\n", raw_input()
   
    # val_caic --> average instructor core of all courses in the college
    # val_cacc --> average course core of all courses in the college
    # val_daic --> average instructor core of all courses in the [dept] department
    # val_dacc --> average course core of all courses in the [dept] department  
    val_caic = caic(asemester[1])
    print "\n8. The college average instructor core is:", str(val_caic) + "."
    val_cacc = cacc(asemester[1])
    print "   The college average course core is:", str(val_cacc) + "."
    val_daic = daic(dept, asemester[1])
    print "   The", str(dept), "average instructor core is:", str(val_daic) + "."
    val_dacc = dacc(dept, asemester[1])
    print "   The", str(dept), "average course core is:", str(val_dacc) + "."
    print "\nPlease press ENTER to continue...\n", raw_input()
    
    # instructor/course average core check
    def aic_level(index, num_list):
        ic_tmp = 0
        for num in num_list:
            if sh[index].cell_value(num, 13) == '':
                ic_tmp += 0
            else:
                ic_tmp += sh[index].cell_value(num, 13)
        if len(num_list) == 0:
            return 0
        else:
            return round(ic_tmp / len(num_list), 2)
    def acc_level(index, num_list):
        cc_tmp = 0
        for num in num_list:
            if sh[index].cell_value(num, 12) == '':
                cc_tmp += 0
            else:
                cc_tmp += sh[index].cell_value(num, 12)
        if len(num_list) == 0:
            return 0
        else:
            return round(cc_tmp / len(num_list), 2)
    
    # val_caic_level: a list stores average instructor core for each level in the college, respectively
    # val_cacc_level: a list stores average course core for each level in the college, respectively
    # val_daic_level: a list stores average instructor core for each level in the [dept] department, respectively
    # val_dacc_level: a list stores average course core for each level in the college, respectively
    print "9. The college average instructor core (caic), college average course core (cacc),\n", dept, "average instructor core (daic), and", dept, "average course core of each level are:\n"
    val_caic_level = list()
    val_caic_level.append(aic_level(i, result_rsps_c['100-200'].keys()))
    val_caic_level.append(aic_level(i, result_rsps_c['300-400'].keys()))
    val_caic_level.append(aic_level(i, result_rsps_c['500-600'].keys()))
    val_caic_level.append(val_caic)
    print '\n   100-200 caic:', val_caic_level[0], ';', '300-400 caic:', val_caic_level[1], ';', '500-600 caic:', val_caic_level[2]
    val_cacc_level = list()
    val_cacc_level.append(acc_level(i, result_rsps_c['100-200'].keys()))
    val_cacc_level.append(acc_level(i, result_rsps_c['300-400'].keys()))
    val_cacc_level.append(acc_level(i, result_rsps_c['500-600'].keys()))
    val_cacc_level.append(val_cacc)
    print '   100-200 cacc:', val_cacc_level[0], ';', '300-400 cacc:', val_cacc_level[1], ';', '500-600 cacc:', val_cacc_level[2]
    val_daic_level = list()
    val_daic_level.append(aic_level(i, result_rsps_dept['100-200'].keys()))
    val_daic_level.append(aic_level(i, result_rsps_dept['300-400'].keys()))
    val_daic_level.append(aic_level(i, result_rsps_dept['500-600'].keys()))
    val_daic_level.append(val_daic)
    print '   100-200 daic:', val_daic_level[0], ';', '300-400 daic:', val_daic_level[1], ';', '500-600 daic:', val_daic_level[2]
    val_dacc_level = list()
    val_dacc_level.append(acc_level(i, result_rsps_dept['100-200'].keys()))
    val_dacc_level.append(acc_level(i, result_rsps_dept['300-400'].keys()))
    val_dacc_level.append(acc_level(i, result_rsps_dept['500-600'].keys()))
    val_dacc_level.append(val_dacc)
    print '   100-200 dacc:', val_dacc_level[0], ';', '300-400 dacc:', val_dacc_level[1], ';', '500-600 dacc:', val_dacc_level[2]   
        
    return counter_enroll_c, counter_enroll_dept, counter_rsps_c, counter_rsps_dept, val_cacc_level, val_caic_level, val_dacc_level, val_daic_level  
summary = checkCrsLvl_all(asemester[0])     

# Output the data to the drawing function for drawing charts
def output(queryName_valid):
    ic_all = []
    cc_all = []
    daic_all = []
    dacc_all = []
    caic_all = []
    cacc_all = []
    semester_all = []
#     global outputData
#     outputData = []
    for num in shnum:
        ic_all.append(ic(courseNum, queryName_valid, num))
        cc_all.append(cc(courseNum, queryName_valid, num))
        daic_all.append(daic(dept, num))
        dacc_all.append(dacc(dept, num))
        caic_all.append(caic(num))
        cacc_all.append(cacc(num))
        semester_all.append(sh[num].cell_value(1, 0))    
#     for num in shnum:
#         outputData.append((str(semester_all[num]), ic_all[num], cc_all[num], daic_all[num], dacc_all[num], caic_all[num], cacc_all[num]))

    return semester_all, ic_all, cc_all, daic_all, dacc_all, caic_all, cacc_all  

output_for_charts = output(queryName_valid)

#print "\n"
#for item in output_for_charts:
#    print item
#
#print "\n"
#for item in outputData:
#    print item

# Write out the data into an Excel file output.xls
def main():
    book = Workbook(encoding='utf-8')
    sheet1 = book.add_sheet('data_output')
    sheet1.write(0, 0, theSemester)
    sheet1.write(1, 0, courseNum[:-2])
    sheet1.write(2, 0, queryName_full)
    sheet1.write(3, 0, enrollNumber)
    sheet1.write(4, 0, rspnsNumOneCrs[-1])
    sheet1.write(5, 0, rspsrate)
    sheet1.write(6, 0, cout)
    sheet1.write(7, 0, iout)
    sheet1.write(8, 0, dept)
    
    for i in range(4):
        sheet1.write(i, 1, summary[1][i])
        sheet1.write(i, 2, summary[3][i])
        if summary[1][i] == 0:
            sheet1.write(i, 3, 0)
        else:
            sheet1.write(i, 3, round(summary[3][i] / summary[1][i], 4))
        sheet1.write(i, 4, summary[6][i])
        sheet1.write(i, 5, summary[7][i])
        
        sheet1.write(i + 4, 1, summary[0][i])
        sheet1.write(i + 4, 2, summary[2][i])
        if summary[0][i] == 0:
            sheet1.write(i + 4, 3, 0)
        else:
            sheet1.write(i + 4, 3, round(summary[2][i] / summary[0][i], 4))
        sheet1.write(i + 4, 4, summary[4][i])
        sheet1.write(i + 4, 5, summary[5][i])
    
    sheet1.write(0, 6, 'Semester')
    sheet1.write(0, 7, 'Course')
    sheet1.write(0, 8, 'Instructor')
    sheet1.write(0, 9, 'Dept Ave.')
    sheet1.write(0, 10, 'College Ave.')
    
    for j in range(len(output_for_charts[0])):
        sheet1.write(j + 1, 6, output_for_charts[0][j])
        sheet1.write(j + 1, 7, output_for_charts[2][j])
        sheet1.write(j + 1, 8, output_for_charts[1][j])
        sheet1.write(j + 1, 9, output_for_charts[3][j])
        sheet1.write(j + 1, 10, output_for_charts[5][j])
    
    book.save('output.xls')
    
    print "\nPlease press ENTER to continue...\n"
    print raw_input()
    print "\nThe report for", courseNum[:-2], "taught by", queryName_full, "in", theSemester, "has been generated.\n" 
    print "Please review the report in Report.xlsm.\n"
    print "Press ENTER to exit", raw_input()

if __name__ == '__main__':
    main()