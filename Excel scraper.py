#Scrapes Allocate excel file for details needed for calendar.py
import openpyxl
import re
import Calendar

wb = openpyxl.load_workbook('timetable.xlsx')
sheet = wb.active
minimum = sheet.min_row+1
maximum = sheet.max_row+1
maxEvents = 0
summary = []        #Done
description = []    #Done
location = []       #Done
start_date = []     #Done
end_date = []       #Done
endRec_date = []    #Done
type_class = []     #Done
reClass = re.compile(r'Laboratory|Lecture|Tutorial|Support-Class')
reDate = re.compile(r'(\d?\d)/(\d?\d)')
reTime = re.compile(r'(\d\d):(\d\d)')
reDuration = re.compile(r'(\d).?(\d)?')

for i in range(minimum, maximum):
    #Summary
    summary.append(sheet['A'+str(i)].value[0:7])
    #Description
    description.append(sheet['B'+str(i)].value)
    #Location
    location.append(sheet['H'+str(i)].value)

    #Regex class type
    reClass2 = reClass.search(sheet['C'+str(i)].value)
    #Class Type
    if reClass2.group() == 'Laboratory':
        type_class.append('304vevnmo631b8nu4cg855lmss@group.calendar.google.com')
    elif reClass2.group() == 'Lecture':
        type_class.append('1m3m8oev40j7eor3q0c3d9jnv4@group.calendar.google.com')
    elif reClass2.group() == 'Support-Class':
        type_class.append('v00gb7jtnnqnt972k6mshuodvs@group.calendar.google.com')
    elif reClass2.group() == 'Tutorial':
        type_class.append('86v4l5efhgi1v6uv9orgjfcdp0@group.calendar.google.com')

    #Regex starting dates
    dateResult = reDate.findall(sheet['K'+str(i)].value)
    day = dateResult[0][0]
    month = dateResult[0][1]
    #Making sure day and month are 2 digits
    if len(day) == 1:
        day = '0' + day
    if len(month) == 1:
        month = '0' + month
    #Start Date
    start_date.append('2018-'+month+'-'+day+'T'+sheet['F'+str(i)].value+':00')

    #Regex start time and duration of classes
    timeResult = reTime.findall(sheet['F'+str(i)].value)
    timeDuration = reDuration.findall(sheet['J'+str(i)].value)
    finishHour = int(timeResult[0][0])+int(timeDuration[0][0])
    #Adding them up
    if timeDuration[0][1] != '':
        if (int(timeResult[0][1])+int(timeDuration[0][1])*6) == 60:
            finishMin = '00'
            finishHour += 1
        else:
            finishMin = str(int(timeResult[0][1])+int(timeDuration[0][1])*6)
    else:
        finishMin = timeResult[0][1]
    #End Date
    end_date.append('2018-'+month+'-'+day+'T'+str(finishHour)+':'+finishMin+':00')
    
    #Recurring Classes
    try:
        day2 = dateResult[1][0]
        month2 = dateResult[1][1]
        if len(day2) == 1:
            day2 = '0' + day2
        if len(month2) == 1:
            month2 = '0' + month2
        endRec_date.append('2018'+month2+day2)
    except:
        endRec_date.append('2018'+month+day)

    #Increment maxEvents
    maxEvents += 1

    #Second session for dates with gaps
    try:
        day = dateResult[2][0]
        month = dateResult[2][1]
        if len(day) == 1:
            day = '0' + day
        if len(month) == 1:
            month = '0' + month
        day2 = dateResult[3][0]
        month2 = dateResult[3][1]
        if len(day2) == 1:
            day2 = '0' + day2
        if len(month2) == 1:
            month2 = '0' + month2
        
        summary.append(sheet['A'+str(i)].value[0:7])
        description.append(sheet['B'+str(i)].value)
        location.append(sheet['H'+str(i)].value)
        start_date.append('2018-'+month+'-'+day+'T'+sheet['F'+str(i)].value+':00')
        end_date.append('2018-'+month+'-'+day+'T'+str(finishHour)+':'+finishMin+':00')
        endRec_date.append('2018'+month2+day2)
        if reClass2.group() == 'Laboratory':
            type_class.append('304vevnmo631b8nu4cg855lmss@group.calendar.google.com')
        elif reClass2.group() == 'Lecture':
            type_class.append('1m3m8oev40j7eor3q0c3d9jnv4@group.calendar.google.com')
        elif reClass2.group() == 'Support-Class':
            type_class.append('v00gb7jtnnqnt972k6mshuodvs@group.calendar.google.com')
        elif reClass2.group() == 'Tutorial':
            type_class.append('86v4l5efhgi1v6uv9orgjfcdp0@group.calendar.google.com')

        maxEvents += 1
    except:
        pass

    #Third session for dates with gaps
    try:
        day = dateResult[4][0]
        month = dateResult[4][1]
        if len(day) == 1:
            day = '0' + day
        if len(month) == 1:
            month = '0' + month
        day2 = dateResult[5][0]
        month2 = dateResult[5][1]
        if len(day2) == 1:
            day2 = '0' + day2
        if len(month2) == 1:
            month2 = '0' + month2
        
        summary.append(sheet['A'+str(i)].value[0:7])
        description.append(sheet['B'+str(i)].value)
        location.append(sheet['H'+str(i)].value)
        start_date.append('2018-'+month+'-'+day+'T'+sheet['F'+str(i)].value+':00')
        end_date.append('2018-'+month+'-'+day+'T'+str(finishHour)+':'+finishMin+':00')
        endRec_date.append('2018'+month2+day2)
        if reClass2.group() == 'Laboratory':
            type_class.append('304vevnmo631b8nu4cg855lmss@group.calendar.google.com')
        elif reClass2.group() == 'Lecture':
            type_class.append('1m3m8oev40j7eor3q0c3d9jnv4@group.calendar.google.com')
        elif reClass2.group() == 'Support-Class':
            type_class.append('v00gb7jtnnqnt972k6mshuodvs@group.calendar.google.com')
        elif reClass2.group() == 'Tutorial':
            type_class.append('86v4l5efhgi1v6uv9orgjfcdp0@group.calendar.google.com')
        maxEvents += 1
    except:
        pass
    
  

for i in range(0, maxEvents):
    Calendar.main(summary[i], description[i], location[i], start_date[i], end_date[i], endRec_date[i], type_class[i])

    
