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

# LD_MTD_YTD_df = pd.read_sql_query(""" """, connection)
LD_Target_df = pd.read_sql_query("""Declare @CurrentMonth NVARCHAR(MAX);
            Declare @DaysInMonth NVARCHAR(MAX);
            SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
            SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
            select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
            where YEARMONTH = @CurrentMonth and AUDTORG= 'RNGSKF'""", connection)
toto = LD_Target_df.values
print(int(toto[0, 0]))
ld_target=int(toto[0, 0])
# sys.exit()


MTD_Target_df = pd.read_sql_query("""Declare @CurrentMonth NVARCHAR(MAX);

            SET @CurrentMonth = convert(varchar(6), GETDATE(),112)

            select ISNULL((Sum(TARGET)), 0) as  MTDTarget from TDCL_BranchTarget  
            where YEARMONTH = @CurrentMonth and AUDTORG= 'RNGSKF' """, connection)

momo = MTD_Target_df.values
Mtd_target=int(momo[0, 0])
print(Mtd_target)
#sys.exit()
from datetime import date

given_date = datetime.today().date()

toto=datetime.today().date()
first = given_date.replace(day=1)
day1=date(toto.year,toto.month,toto.day)
day2=date(first.year,first.month,first.day)
no_of_days=(day1-day2).days
print(no_of_days)

import calendar
import datetime
now = datetime.datetime.now()
total_days=calendar.monthrange(now.year, now.month)[1]
print(total_days)

final_mtd_target = int((Mtd_target/total_days)*no_of_days)
print(final_mtd_target)
#sys.exit()

YTD_Target_df = pd.read_sql_query(""" select ISNULL((Sum(TARGET)), 0) as  YTDTarget from TDCL_BranchTarget  
            where convert(varchar(4), YEARMONTH,112) = convert(varchar(4), GETDATE(),112)
              and AUDTORG= 'RNGSKF'""",
                                  connection)

yoyo = YTD_Target_df.values
print(int(yoyo[0, 0]))


Ytd_target=int(yoyo[0, 0])
print(Ytd_target)

final_ytd_target=int(Ytd_target-((Mtd_target/total_days)*(total_days-no_of_days)))
print(final_ytd_target)
#sys.exit()
LD_Sales_df = pd.read_sql_query("""Declare @Currentday NVARCHAR(MAX);

            SET @Currentday = convert(varchar(8), GETDATE()-1,112);

            select  Sum(EXTINVMISC) as  YesterdaySales from OESalesDetails  
            where LEFT(TRANSDATE,8) = @Currentday and AUDTORG='RNGSKF' """, connection)

Ld_toto = LD_Sales_df.values
print(int(Ld_toto[0, 0]))
# sys.exit()

MTD_Sales_df = pd.read_sql_query(""" Declare @Currentmonth NVARCHAR(MAX);

            SET @Currentmonth = convert(varchar(6), GETDATE(),112);

            select  Sum(EXTINVMISC) as  MTDSales from OESalesDetails  
            where LEFT(TRANSDATE,6) = @Currentmonth and AUDTORG='RNGSKF'""", connection)

MTD_momo = MTD_Sales_df.values
print(int(MTD_momo[0, 0]))
# sys.exit()

YTD_Sales_df = pd.read_sql_query("""Declare @Currentyear NVARCHAR(MAX);

            SET @Currentyear = convert(varchar(4), GETDATE(),112);

            select  Sum(EXTINVMISC) as  YTDSales from OESalesDetails  
            where LEFT(TRANSDATE,4) = @Currentyear and AUDTORG='RNGSKF' """, connection)

YTD_yoyo = YTD_Sales_df.values
print(int(YTD_yoyo[0, 0]))
#sys.exit()

labels = ['LD', 'MTD', 'YTD']
aa =int(final_mtd_target/1000)
print('a',aa)
bb= int(final_ytd_target/1000)
print('b',bb)
cc=int(ld_target / 1000)
Targets = [cc, aa,bb]
print(Targets)
#Targets = [int(toto[0,0]/1000), int(momo[0,0]/1000), int(yoyo[0,0]/1000)]
Sales = [int(Ld_toto[0, 0] / 1000), int(MTD_momo[0, 0] / 1000), int(YTD_yoyo[0, 0] / 1000)]

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

fig, ax = plt.subplots()
rects1 = ax.bar(x - width / 2, Targets, width, label='Target')
rects2 = ax.bar(x + width / 2, Sales, width, label='Sales')

# Add some text for labels, title and custom x-axis tick labels, etc.
ax.set_ylabel('Amount', fontsize='14', color='black', fontweight='bold')
ax.set_xlabel('Group Wise Sales', fontsize='14', color='black', fontweight='bold')
ax.set_title('11. Target vs Sales', fontsize=16, fontweight='bold', color='#3e0a75')

ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.legend()


def autolabel(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .995 * height,
                get_value(str(height)) + "K",
                ha='center', va='bottom', fontsize=12, rotation=5, fontweight='bold')


autolabel(rects1)
autolabel(rects2)

fig.tight_layout()

#plt.show()
plt.savefig('LD_MTD_YTD_TARGET_vs_sales.png')
print('11. Target vs Sales ')