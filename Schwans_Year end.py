##Change Made##
from openpyxl import Workbook
import openpyxl

from datetime import datetime
from datetime import timedelta

from dateutil.relativedelta import *
import uuid


wb_write=Workbook() 
write1=wb_write.active  #Activating Workbook


Excel__Line_counter=1 # Counter for traversing to new lines in Excel sheet 
Unique_ID_list=[]

print("\nEnter the below information for Monthly Periods\n")

Program_Year=int(input('Enter the Program Year : ')) #Program Year Provided in the Work Item

Monthly_key=int(input('Enter the start of Key in Monthly periods(As per provided in the Work Item : ')) # Key for Monthly Periods as Provided in the Work Item

Year_monthStart=datetime(Program_Year,1,1) #Date for the Monthly Periods

monthtrack=Year_monthStart # Monthly Traversal Tracker

for i in range(1,13): # Loop for inputting data for Monthly Periods
	Excel__Line_counter=Excel__Line_counter+1
	uqid=uuid.uuid1()
	UniqueId=str(uqid)
	Unique_ID_list.append(UniqueId)
	write1['A{}'.format(Excel__Line_counter)]=UniqueId
	write1['B{}'.format(Excel__Line_counter)]=monthtrack.strftime("%B") #Displaying Month from date
	# monthtrack=monthtrack+1
	write1['C{}'.format(Excel__Line_counter)]=Monthly_key
	Monthly_key=Monthly_key+1
	write1['D{}'.format(Excel__Line_counter)]=monthtrack
	write1['E{}'.format(Excel__Line_counter)]=monthtrack+relativedelta(months=+1)-timedelta(days=1)
	write1['F{}'.format(Excel__Line_counter)]=Program_Year
	write1['G{}'.format(Excel__Line_counter)]='9ECEF9E1-7589-4015-A516-E34AC7301924'
	write1['H{}'.format(Excel__Line_counter)]='NULL'
	monthtrack=monthtrack+relativedelta(months=+1)
	


print('\n***Information for Weekly Peroiods Starts below***\n')


periods=int(input('Enter the number of Periods for weekly periods(Should be around 52-53. Please increase/Decrease the periods as per requirement : '))


date = str(input('Enter the Date when the Sunday Period Begins (YYYY-MM-DD)')) #date when the sunday period starts
day=date.split('-')


Sunday_Period=datetime(int(day[0]),int(day[1]),int(day[2])) #Starting Weekly period


for i in range (1,periods+1): #Loop for inputting Data for Weekly Periods(Sunday and Monday Periods)
	
	Sunday_Period_end=Sunday_Period+timedelta(days=7) 

	Monday_Period=Sunday_Period+timedelta(days=1)
	Monday_Period_End=Monday_Period+timedelta(days=7)


	Excel__Line_counter=Excel__Line_counter+1
	
	write1['A{}'.format(Excel__Line_counter)]=str(uuid.uuid1())
	write1['B{}'.format(Excel__Line_counter)]='NULL'
	# monthtrack=monthtrack+1
	write1['C{}'.format(Excel__Line_counter)]=i
	Monthly_key=Monthly_key+1
	write1['D{}'.format(Excel__Line_counter)]=Sunday_Period
	write1['E{}'.format(Excel__Line_counter)]=Sunday_Period_end
	write1['F{}'.format(Excel__Line_counter)]=Program_Year
	write1['G{}'.format(Excel__Line_counter)]='CD20DAFB-D97D-4DF0-859E-8CCAB9FFFD38'
	write1['H{}'.format(Excel__Line_counter)]=Unique_ID_list[Sunday_Period.month-1]
	
	Excel__Line_counter=Excel__Line_counter+1
	
	write1['A{}'.format(Excel__Line_counter)]=str(uuid.uuid1())
	write1['B{}'.format(Excel__Line_counter)]='NULL'
	# monthtrack=monthtrack+1
	write1['C{}'.format(Excel__Line_counter)]=i
	Monthly_key=Monthly_key+1
	write1['D{}'.format(Excel__Line_counter)]=Monday_Period
	write1['E{}'.format(Excel__Line_counter)]=Monday_Period_End
	write1['F{}'.format(Excel__Line_counter)]=Program_Year
	write1['G{}'.format(Excel__Line_counter)]='CD20DAFB-D97D-4DF0-859E-8CCAB9FFFD38'
	write1['H{}'.format(Excel__Line_counter)]=Unique_ID_list[Monday_Period.month-1]

	Sunday_Period=Sunday_Period_end
	Monday_Period=Monday_Period_End


wb_write.save('Yearend.xlsx')
print('Flie Ready!!!')











