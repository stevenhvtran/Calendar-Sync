#Improved Version of Allocate+ -> Google Calendar
import re, openpyxl, Calendar
from datetime import datetime, timedelta
wb = openpyxl.load_workbook('timetable.xlsx')
sheet = wb.active
    
#PARSES EXCEL DATA INTO USEABLE DATA
def format_subjectCode(sjCode):
    return sjCode[0:7]

def format_description(desc):
    #Translates short-form to full descriptions
    desc_dict = {'INTRO CSN&SEC' : 'Introduction to computer systems, networks and security',
                 'ALG PROG PYTHON ADV' : 'Algorithms and programming in python (advanced)',
                 'DISC MATH COMPSC' : 'Discrete mathematics for computer science',
                 'TECHNQS.FOR MODLNG.' : 'Techniques for modelling'
                 }
    return desc_dict[desc]

def format_group(group):
    #Translates group to corresponding Google Calendar ID
    group_url = {'Labo' : 'gvquuiks4so096cqf03ff09oug@group.calendar.google.com',
                 'Lect' : 'ljspq0s3acqdfd60moq3stl1po@group.calendar.google.com',
                 'Supp' : 'dm26abfjlu02377it5b8i2h610@group.calendar.google.com',
                 'Tuto' : 'ki66knb2lffb0nccap4ltedrls@group.calendar.google.com'
                 }
    return group_url[group[0:4]]

def format_time(time):
    time = time.split(':')  #Puts time into a list [hh,mm]
    return timedelta(hours = int(time[0]), minutes = int(time[1]))

def format_location(loc):
    #Translates street code to full street name
    location_dict = {'Rnf' : 'Rainforest Walk',
                     'Exh' : 'Exhibition Walk',
                     'All' : 'Alliance Lane',
                     'Inn' : 'Innovation Walk',
                     'Anc' : 'Ancora Imparo Way'
                     }
    #Regex formatting of location to *number* *street name* *building name*
    reLoc = re.compile(r'CL_(\d?\d)(\w\w\w)/(\S?\S?\S?\S?)')
    reLoc = reLoc.findall(loc)
    return '{0} {1} ({2})'.format(reLoc[0][0], location_dict[reLoc[0][1]], reLoc[0][2])
    
def format_duration(dur):
    #Returns a datetime value with class duration using regex
    reDur = re.compile(r'\d.?\d?')
    return timedelta(hours = float(reDur.findall(dur)[0]))

def format_dates(date):
    #Creates nested strings in which each class period and start-end dates are split up
    date = date.split(', ')
    for i in range(0,len(date)):
        date[i] = date[i].split('-')
        for j in range(0, len(date[i])):
            date[i][j] = datetime.strptime(date[i][j]+'/2018', '%d/%m/%Y')
    return date

def time_format(time):
    return str(time.isoformat())

def recTime_format(time):
    return str(time.strftime('%Y%m%d'))

#Outputs n row data in a dictionary to be further processed
def excel_row(row_number):
    r = str(row_number)    
    #Gives types of data and cell-coordinates
    data_types = {'subjectCode' : 'A'+r,
                  'description' : 'B'+r,
                  'group' : 'C'+r,
                  'time' : 'F'+r,
                  'location' : 'H'+r,
                  'duration' : 'J'+r,
                  'dates' : 'K'+r
                  }
    #Creates new dictionary with co-ord replaced with formatted data
    for data_name, cell_coord in data_types.items():
        string_combined = "format_{0}('{1}')".format(data_name, sheet[cell_coord].value)
        data_types[data_name] = eval(string_combined)  
    return data_types

def main():
    for data in range(sheet.min_row+1, sheet.max_row+1):
        #Replaces the integer with useful data
        data = excel_row(data)
        for class_period in range(0,len(data['dates'])):
            #Changes formatting of times to what they need to be
            start_datetime = data['dates'][class_period][0] + data['time'] #Remember to use time_format later
            end_datetime = start_datetime + data['duration']
            try:
                end_recurrence = data['dates'][class_period][1]
            except:
                end_recurrence = data['dates'][class_period][0]
            #Initialises everything
            Calendar.main(data['subjectCode'],
                          data['description'],
                          data['location'],
                          time_format(start_datetime),
                          time_format(end_datetime),
                          recTime_format(end_recurrence),
                          data['group']
                          )
