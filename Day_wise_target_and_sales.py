import os
import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from math import log, floor
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import pytz
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyodbc as db
import xlrd
from PIL import Image
from matplotlib.patches import Patch
import datetime as dd
import sys

from PIL import Image, ImageDraw, ImageFont
import pytz
import datetime as dd
from PIL import Image
from datetime import datetime

# ----------------------------new code-----------------------------
# ------------------------------------------------------------------
def hazar(number):
    k = 1000.0
    final_number = number / 1000
    return '%.1f %s ' % (final_number, 'K')


# ------------------------------------------------------------------
def joker(number):
    number = number / 1000
    number = int(number)
    number = format(number, ',')
    number = number + 'K'
    return number


# ---------------------------------------------------------------------
def for_bar(number):
    number = round(number, 1)
    number = format(number, ',')
    number = number + 'K'
    return number

def get_value(value):
    if (len(value) > 6):
        return str(value[0:len(value) - 6] + "," + value[len(value) - 6:len(value) - 3] + ","
                   + value[len(value) - 3:len(value)])
    elif (len(value) > 3):
        return str(value[0:len(value) - 3] + "," + value[len(value) - 3:len(value)])
    elif (len(value) > 0):
        return value
    else:
        return "-"

dirpath = os.path.dirname(os.path.realpath(__file__))

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=137.116.139.217;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp@123')


cursor = connection.cursor()

daily_sales_df = pd.read_sql_query("""select Right(transdate,2) as [day], Sum(EXTINVMISC)/1000 as  EverydaySales from OESalesDetails  
            where LEFT(TRANSDATE,6) =convert(varchar(6), GETDATE(),112)  and AUDTORG='RNGSKF'
			group by transdate
			order by transdate""", connection)


EveryD_Target_df = pd.read_sql_query("""Declare @CurrentMonth NVARCHAR(MAX);
            Declare @DaysInMonth NVARCHAR(MAX);
            SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
            SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
            select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
            where YEARMONTH = @CurrentMonth and AUDTORG= 'RNGSKF'""", connection)
totarget = EveryD_Target_df.values
target_for_target=int(totarget[0, 0])
# sys.exit()


Every_day = daily_sales_df['day'].tolist()
print(Every_day)
print(len(Every_day))
y_pos = np.arange(len(Every_day))
print(y_pos)
every_day_sale = daily_sales_df['EverydaySales'].tolist()
print(every_day_sale)
n=1
Target=[]
labell=[]
for z in y_pos:
    labell.append(n)
    Target.append(int(target_for_target/1000))
    n=n+1
print(Target)
print(labell)
#sys.exit()
#aging_total=sum(performance)
fig, ax = plt.subplots(figsize=(12.81, 4.8))
labels = labell
#Targets = [1000,2000,3000,4000,5000,6000,7000,8000,9000,10000,11000,12000]
#Sales = [int(Ld_toto[0, 0] / 1000), int(MTD_momo[0, 0] / 1000), int(YTD_yoyo[0, 0] / 1000)]

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

#fig, ax = plt.subplots()
#rects1 = ax.bar(x - width / 2, Target, width, label='Target')
rects2 = ax.bar(y_pos, every_day_sale, width, label='Sales')

# Add some text for labels, title and custom x-axis tick labels, etc.
line=ax.plot(Target, color='orange', label='Target')
ax.set_ylabel('Amount', fontsize='14', color='black', fontweight='bold')
ax.set_xlabel('Day', fontsize='14', color='black', fontweight='bold')
ax.set_title('12. MTD Target vs Sales', fontsize=16, fontweight='bold', color='#3e0a75')

ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.legend()


def autolabel(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .995 * height,
                get_value(str(height)) + "K",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


#autolabel(rects1)
autolabel(rects2)

fig.tight_layout()

#plt.show()
plt.savefig("Day_Wise_Target_vs_Sales.png")
print('9. Day Wise Target vs Sales')