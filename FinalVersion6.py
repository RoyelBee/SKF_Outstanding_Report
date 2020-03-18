import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from math import log, floor
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyodbc as db
import xlrd
from matplotlib.patches import Patch
from PIL import Image, ImageDraw, ImageFont
import pytz
import datetime as dd
from PIL import Image
from datetime import datetime

import last_day_mtd_ytd_sales_target
import Day_wise_target_and_sales

date = datetime.today()
day = str(date.day) + '/' + str(date.month) + '/' + str(date.year)
tz_NY = pytz.timezone('Asia/Dhaka')
datetime_BD = datetime.now(tz_NY)
time = datetime_BD.strftime("%I:%M %p")
date = datetime.today()
x = dd.datetime.now()
day = str(date.day) + '-' + str(x.strftime("%b")) + '-' + str(date.year)
print(date)
print(day)
tz_NY = pytz.timezone('Asia/Dhaka')
datetime_BD = datetime.now(tz_NY)
time = datetime_BD.strftime("%I:%M %p")
print(datetime_BD)
print(time)
img = Image.open("new_ai.png")
title = ImageDraw.Draw(img)
timestore = ImageDraw.Draw(img)
tag = ImageDraw.Draw(img)
branch = ImageDraw.Draw(img)
font = ImageFont.truetype("Stencil_Regular.ttf", 60, encoding="unic")
font1 = ImageFont.truetype("ROCK.ttf", 50, encoding="unic")
font2 = ImageFont.truetype("ROCK.ttf", 35, encoding="unic")
branch_name = 'Rangpur'
#
tag.text((25, 8), 'SK+F', (255, 255, 255), font=font)
branch.text((25, 270), branch_name + " Branch", (255, 209, 0), font=font1)
timestore.text((25, 380), time + "\n" + day, (255, 255, 255), font=font2)
img.save('banner_ai.png')

img.show()


# sys.exit()
# ----------------------------new code-----------------------------
# -----------convert the number into thousand----------------------
def hazar(number):
    k = 1000.0
    final_number = number / 1000
    return '%.1f %s ' % (final_number, 'K')

# ----------convert the number into thousand and give comma -------
def joker(number):
    number = number / 1000
    number = int(number)
    number = format(number, ',')
    number = number + 'K'
    return number

# --------------give comma in a number and add k-------------------
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


# ------------------------Add comma---------------------------------------

# -------------------------------------------------------------------
dirpath = os.path.dirname(os.path.realpath(__file__))

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=137.116.139.217;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp@123')


def human_format(number):
    units = ['', 'K', 'M', 'B', 'T', 'P']
    k = 1000.0
    magnitude = int(floor(log(number, k)))
    return '%.1f %s ' % (number / k ** magnitude, units[magnitude])


cursor = connection.cursor()

outstanding_df = pd.read_sql_query(""" select
            SUM(CASE WHEN TERMS='CASH' THEN OUT_NET END) AS TotalOutStandingOnCash,
            SUM(CASE WHEN TERMS not like '%CASH%' THEN OUT_NET END) AS TotalOutStandingOnCredit

            from  [ARCOUT].dbo.[CUST_OUT]
            where AUDTORG = 'RNGSKF' AND [INVDATE] <= convert(varchar(8),DATEADD(D,0,GETDATE()),112)
                                    """, connection)

cash = int(outstanding_df['TotalOutStandingOnCash'])
credit = int(outstanding_df['TotalOutStandingOnCredit'])

data = [cash, credit]
total = cash + credit
# total = format((cash + credit), ',')
print(total)
# total = format(total, ',')
# total=hazar(total)
# total = total/1000
# total=round(total,1)
# total=format(total, ',')
total = 'Total \n' + joker(total)
# colors
colors = ['#f9ff00', '#ff8600']

legend_element = [Patch(facecolor='#f9ff00', label='Cash'),
                  Patch(facecolor='#ff8600', label='Credit')]

# ca = format(cash, ',')
# cre = format(credit, ',')

# -------------------new code--------------------------


ca = joker(cash)
cre = joker(credit)

DataLabel = [ca, cre]
# -----------------------------------------------------

fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(data, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=90, pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, total, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#00daff')

# draw circle
centre_circle = plt.Circle((0, 0), 0.50, fc='white')

fig = plt.gcf()

fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('1. Total Outstanding', fontsize=16, fontweight='bold', color='#3e0a75')

ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left',
           fontsize=11)
plt.savefig('terms_wise_outstanding.png')

# plt.show()
print('1. Terms wise Outstanding Generated')

# sys.exit()

# #--------- Category wise Credit ( Matured , Not Mature) -------------
credit_category_df = pd.read_sql_query(""" Select case when Days_Diff>0 then 'Matured Credit' else 'Regular Credit' End as  'Category',Sum(OUT_NET) as Amount from
                    (select INVNUMBER,INVDATE,
                    CUSTOMER,TERMS,MAINCUSTYPE,
                    CustomerInformation.CREDIT_LIMIT_DAYS,
                    datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                    OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                    join ARCHIVESKF.dbo.CustomerInformation
                    on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                    where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash') as TblCredit
                    group by case when Days_Diff>0 then 'Matured Credit' else 'Regular Credit' end
                                                        """, connection)

matured = int(credit_category_df.Amount.iloc[0])
not_mature = int(credit_category_df.Amount.iloc[1])
#
# print('Matured= ', matured)
# print("Regular = ", regular)

values = [matured, not_mature]
# colors
colors = ['#ffb667', '#b35e00']

legend_element = [Patch(facecolor='#ffb667', label='Matured'),
                  Patch(facecolor='#b35e00', label='Not Mature')]

total_credit = matured + not_mature
# total_credit = format(total_credit, ',')


total_credit = 'Total \n' + joker(total_credit)
# matured = format(matured, ',')
# not_mature = format(not_mature, ',')

# ------------------new code--------------------
matured = joker(matured)
not_mature = joker(not_mature)

# matured=hazar(matured)
# not_mature=hazar(not_mature)
DataLabel = [matured, not_mature]
# -----------------------------------------------

fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(values, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=120,
                                 pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, total_credit, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#00daff')

# draw circle
centre_circle = plt.Circle((0, 0), 0.50, fc='white')

fig = plt.gcf()

fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('2. Credit Outstanding', fontsize=16, fontweight='bold', color='#3e0a75')

ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left', fontsize=11)
plt.savefig('category_wise_credit.png')

# plt.show()
print('2. Category wise Credit Generated ')

# sys.exit()


# #------- Sector Wise Credit not_mature -------------------
sector_credit_df = pd.read_sql_query(""" Select case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end  as CustType,Sum(OUT_NET) as Amount from
                (select INVNUMBER,INVDATE,
                CUSTOMER,TERMS,MAINCUSTYPE,
                CustomerInformation.CREDIT_LIMIT_DAYS,
                datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                join ARCHIVESKF.dbo.CustomerInformation
                on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash') as TblCredit
                where Days_Diff<=0
                group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end""", connection)

Institution = int(sector_credit_df.Amount.iloc[0])
retail = int(sector_credit_df.Amount.iloc[1])
#
# print('retail= ', retail)
# print("Institution = ", Institution)

values = [Institution, retail]
# colors
colors = ['#e1052f', '#00cd6b']

legend_element = [Patch(facecolor='#e1052f', label='Institution'),
                  Patch(facecolor='#00cd6b', label='Retail')]

sector_total_credit = Institution + retail
sector_total_credit = "Total \n" + joker(sector_total_credit)

Institution = joker(Institution)
retail = joker(retail)
DataLabel = [Institution, retail]
fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(values, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=120,
                                 pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, sector_total_credit, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#00daff')

# draw circle
centre_circle = plt.Circle((0, 0), 0.50, fc='white')

fig = plt.gcf()

fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('5. Non-Matured Credit', fontsize=16, fontweight='bold', color='#3e0a75')

ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left', fontsize=11)
plt.savefig('regular_credit.png')

# plt.show()
print('5. Sector wise Non Mature Credit- Regular Generated')

# ------- Sector Wise Credit Mature  -------------------
mature_credit_df = pd.read_sql_query("""
                Select case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end  as CustType,Sum(OUT_NET) as Amount from
                (select INVNUMBER,INVDATE,
                CUSTOMER,TERMS,MAINCUSTYPE,
                CustomerInformation.CREDIT_LIMIT_DAYS,
                datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                join ARCHIVESKF.dbo.CustomerInformation
                on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash') as TblCredit
                where Days_Diff>0
                group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end
                """, connection)

Institution = int(mature_credit_df.Amount.iloc[0])
retail = int(mature_credit_df.Amount.iloc[1])

# print('retail= ', retail)
# print("Institution = ", Institution)

values = [Institution, retail]
# colors
colors = ['#f213e5', '#bcf303']

legend_element = [Patch(facecolor='#f213e5', label='Institution'),
                  Patch(facecolor='#bcf303', label='Retail')]

Sector_total_matured = Institution + retail

Sector_total_matured = 'Total \n' + joker(Sector_total_matured)

Institution = joker(Institution)
retail = joker(retail)
DataLabel = [Institution, retail]
fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(values, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=130,
                                 pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, Sector_total_matured, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#00daff')

# draw circle
centre_circle = plt.Circle((0, 0), 0.50, fc='white')

fig = plt.gcf()

fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('3. Matured Credit', fontsize=16, fontweight='bold', color='#3e0a75')

ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left', fontsize=11)
plt.savefig('matured_credit.png')

# # plt.show()
print('3. Sector wise Credit- Matured Generated')

# #-- Closed to matured credit --------------------
CloseTo_mature_df = pd.read_sql_query("""
                SELECT  AgingDays, sum(CreditAmount)/1000 as Amount FROM
                (Select CUSTOMER, CUSTNAME, INVDATE,   Sum(OUT_NET) as CreditAmount ,
                case
                 when TblCredit.Days_Diff between '-3' and '0'  THEN '0 - 3 days'
                when TblCredit.Days_Diff between '-10' and '-4'  THEN '4 - 10 days'
                when TblCredit.Days_Diff between '-15' and '-11'  THEN '11 - 15 days'
                else '16+ Days' end  as AgingDays,

				CASE
				when TblCredit.Days_Diff between '-3' and '0' then 1
				when TblCredit.Days_Diff between '-10' and '-4' then 2
				when TblCredit.Days_Diff between '-15' and '-11'  then 3
				ELSE  4
				END AS SERIAL
                from
                    (select CUSTNAME, INVNUMBER,INVDATE,
            CUSTOMER,TERMS,MAINCUSTYPE,
            CustomerInformation.CREDIT_LIMIT_DAYS,
            (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) as Days_Diff,
            OUT_NET from [ARCOUT].dbo.[CUST_OUT]
            join ARCHIVESKF.dbo.CustomerInformation
            on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
            where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash' and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)<=0
			) as TblCredit
           -- where Days_Diff<=0

                    group by CUSTOMER, CUSTNAME, INVDATE,
                        case
                 when TblCredit.Days_Diff between '-3' and '0'  THEN '0 - 3 days'
                when TblCredit.Days_Diff between '-10' and '-4'  THEN '4 - 10 days'
                when TblCredit.Days_Diff between '-15' and '-11'  THEN '11 - 15 days'
                             else '16+ Days' end

				,CASE
				when TblCredit.Days_Diff between '-3' and '0' then 1
				when TblCredit.Days_Diff between '-10' and '-4' then 2
				when TblCredit.Days_Diff between '-15' and '-11'  then 3
				ELSE  4
				END ) AS T1
				group by T1.AgingDays, SERIAL
				order by SERIAL

                                """, connection)

width = 0.75
AgingDays = CloseTo_mature_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = CloseTo_mature_df['Amount']
tovalue = sum(performance)
print(tovalue)

# sys.exit()
fig, ax = plt.subplots()
bars = plt.bar(y_pos, performance, width, align='center', alpha=1)

bars[0].set_color('#f00228')
bars[1].set_color('#ff6500')
bars[2].set_color('#deff00')
bars[3].set_color('#2c8e14')


def autolabel(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .995 * height,
                for_bar(height),
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel(bars)


def autolabel2(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .5 * height,
                str(round(((height / tovalue) * 100), 1)) + "%",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel2(bars)

plt.xticks(y_pos, AgingDays, fontsize=12)
plt.yticks(fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('6. Non-Matured Credit Age', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('closed_to_matured_credit.png')
# # plt.show()
print('6. Closed to mature Credit Generated')

# sys.exit()
# # --- Aging Matured Credit ---------------------------------
aging_mature_df = pd.read_sql_query("""
                 SELECT  AgingDays, sum(Amount)/1000 as Amount FROM
                (Select
                case
                 when TblCredit.Days_Diff between '1' and '3'  THEN '1 - 3 days'
                when TblCredit.Days_Diff between '4' and '10'  THEN '4 - 10 days'
                when TblCredit.Days_Diff between '11' and '15'  THEN '11 - 15 days'
                else '16+ Days' end  as AgingDays,
                --OUT_NET
                Sum(OUT_NET) as Amount,

				CASE
				when TblCredit.Days_Diff between '1' and '3'  THEN 1
                when TblCredit.Days_Diff between '4' and '10'  THEN 2
                when TblCredit.Days_Diff between '11' and '15'  THEN 3
				ELSE  4
				END AS SERIAL
                from
                    (select INVNUMBER,INVDATE,
            CUSTOMER,TERMS,MAINCUSTYPE,
            CustomerInformation.CREDIT_LIMIT_DAYS,
            (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) as Days_Diff,
            OUT_NET from [ARCOUT].dbo.[CUST_OUT]
            join ARCHIVESKF.dbo.CustomerInformation
            on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
            where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash' ) as TblCredit
            where Days_Diff>0

                    group by
                        case
                 when TblCredit.Days_Diff between '1' and '3'  THEN '1 - 3 days'
                when TblCredit.Days_Diff between '4' and '10'  THEN '4 - 10 days'
                when TblCredit.Days_Diff between '11' and '15'  THEN '11 - 15 days'
                             else '16+ Days' end,
						CASE
				when TblCredit.Days_Diff between '1' and '3'  THEN 1
                when TblCredit.Days_Diff between '4' and '10'  THEN 2
                when TblCredit.Days_Diff between '11' and '15'  THEN 3
				ELSE  4 END ) AS T1
				group by T1.AgingDays, SERIAL
				order by SERIAL
                                    """, connection)

width = 0.75
AgingDays = aging_mature_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = aging_mature_df['Amount']
aging_total = sum(performance)
fig, ax = plt.subplots()
bars = plt.bar(y_pos, performance, width, align='center', alpha=1, color=colors)


def autolabel(bars):
    # attach some text labels
    for rect in bars:
        height = int(rect.get_height())
        ax.text(rect.get_x() + rect.get_width() / 2., .995 * height,
                for_bar(height), ha='center', va='bottom', fontsize=12, fontweight='bold')


bars[0].set_color('#2c8e14')
bars[1].set_color('#deff00')
bars[2].set_color('#ff6500')
bars[3].set_color('#f00228')

autolabel(bars)


def autolabel3(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .5 * height,
                str(round(((height / aging_total) * 100), 1)) + "%",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel3(bars)

plt.xticks(y_pos, AgingDays, fontsize=12)
plt.yticks(fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('4. Matured Credit Age', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('aging_matured_credit.png')
# plt.close()
# plt.show()
print('4. Aging Mature Credit Generated')

# --------------------------- Sector wise cash --------------------------
sector_wise_cash_df = pd.read_sql_query(""" Select case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end  as CustType,Sum(OUT_NET) as Amount from
                (select INVNUMBER,INVDATE,
                CUSTOMER,TERMS,MAINCUSTYPE,
                CustomerInformation.CREDIT_LIMIT_DAYS,
                datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                join ARCHIVESKF.dbo.CustomerInformation
                on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS='Cash') as TblCredit

                group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end """, connection)
Institution = int(sector_wise_cash_df.Amount.iloc[0])
retail = int(sector_wise_cash_df.Amount.iloc[1])
#
# print('retail= ', retail)
# print("Institution = ", Institution)
values = [Institution, retail]
# colors
colors = ['#a9e11a', '#1798ca']
legend_element = [Patch(facecolor='#a9e11a', label='Institution'),
                  Patch(facecolor='#1798ca', label='Retail')]
sector_total_credit = Institution + retail
sector_total_credit = "Total \n" + joker(sector_total_credit)

Institution = joker(Institution)
retail = joker(retail)
DataLabel = [Institution, retail]

fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(values, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=120,
                                 pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, sector_total_credit, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#b1bec1')
# draw circle
centre_circle = plt.Circle((0, 0), 0.50, fc='white')
fig = plt.gcf()
fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('7. Cash Outstanding', fontsize=16, fontweight='bold', color='#3e0a75')
ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left', fontsize=11)
plt.savefig('Category_wise_cash.png')
# plt.show()
print('7. Category wise Cash Generated')

# ----------------------- Cash Drop -----------------------------------

cash_drop_df = pd.read_sql_query(""" SELECT  AgingDays, sum(Amount)/1000 as Amount FROM (  Select
    case
		when Days_Diff between 0 and 3  THEN '0 - 3 days'
		when Days_Diff between 4 and 10  THEN '4 - 10 days'
		when Days_Diff between 11 and 15  THEN '11 - 15 days'
		else '16+ Days' end  as AgingDays,
    --OUT_NET
    Sum(OUT_NET) as Amount,

	CASE
		when TblCredit.Days_Diff between '0' and '3'  THEN 1
        when TblCredit.Days_Diff between '4' and '10'  THEN 2
        when TblCredit.Days_Diff between '11' and '15'  THEN 3
		ELSE  4
		END AS SERIAL
    from
        (select INVNUMBER,INVDATE,
        CUSTOMER,TERMS,MAINCUSTYPE,
        datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 as Days_Diff,
        OUT_NET from [ARCOUT].dbo.[CUST_OUT]
        where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS='Cash') as TblCredit
        group by

        case
            when Days_Diff between 0 and 3 THEN '0 - 3 days'
            when Days_Diff between 4 and 10  THEN '4 - 10 days'
            when Days_Diff between 11 and 15  THEN '11 - 15 days'
                else '16+ Days' end
		,CASE
		when TblCredit.Days_Diff between '0' and '3'  THEN 1
        when TblCredit.Days_Diff between '4' and '10'  THEN 2
        when TblCredit.Days_Diff between '11' and '15'  THEN 3
		ELSE  4 end ) as T1
		group by T1.AgingDays, SERIAL
	order by SERIAL """, connection)

width = 0.75
AgingDays = cash_drop_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = cash_drop_df['Amount']
new_totall = sum(performance)
fig, ax = plt.subplots()
bars = plt.bar(y_pos, performance, width, align='center', alpha=0.9, color='#fd00ff')

bars[0].set_color('#f00228')
bars[1].set_color('#ff6500')
bars[2].set_color('#deff00')
bars[3].set_color('#2c8e14')


def autolabel(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .995 * height,
                for_bar(height),
                ha='center', va='bottom', fontweight='bold', fontsize=12)


autolabel(bars)


def autolabel4(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .5 * height,
                str(round(((height / new_totall) * 100), 1)) + "%",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel4(bars)

plt.xticks(y_pos, AgingDays, fontsize=12)
plt.yticks(fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('8. Aging of Cash Drop', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('aging_cash_drop.png')

# # plt.show()
print('8. Aging Days - Cash Drop Generated')

# ------------- Join credit  figures ---------------------------

imp1 = Image.open(dirpath + "/terms_wise_outstanding.png")
widthx, heightx = imp1.size
imp2 = Image.open(dirpath + "/category_wise_credit.png")
imageSize = Image.new('RGB', (1283, 481))
imageSize.paste(imp1, (1, 0))
imageSize.paste(imp2, (widthx + 2, 0))
imageSize.save(dirpath + "/all_credit.png")

imp3 = Image.open(dirpath + "/matured_credit.png")
widthx, heightx = imp1.size
imp4 = Image.open(dirpath + "/aging_matured_credit.png")

imageSize = Image.new('RGB', (1283, 481))
imageSize.paste(imp3, (1, 0))
imageSize.paste(imp4, (widthx + 2, 0))
imageSize.save(dirpath + "/all_Matured_credit.png")

# ---------------------------------------------------------

imp10 = Image.open(dirpath + "/Category_wise_cash.png")
widthx, heightx = imp1.size
imp11 = Image.open(dirpath + "/aging_cash_drop.png")

imageSize = Image.new('RGB', (1283, 481))
imageSize.paste(imp10, (1, 0))
imageSize.paste(imp11, (widthx + 2, 0))
imageSize.save(dirpath + "/all_cash.png")

# ------------- Closed to Matured, Matured Credit  ---------------------------

imp5 = Image.open(dirpath + "/regular_credit.png")
widthx, heightx = imp1.size
imp6 = Image.open(dirpath + "/closed_to_matured_credit.png")

imageSize = Image.new('RGB', (1283, 481))
imageSize.paste(imp5, (1, 0))
imageSize.paste(imp6, (widthx + 2, 0))
imageSize.save(dirpath + "/all_regular_credit.png")

# # --------------------return kpi----------------------------
# # ----------------------------------------------------------
# # ----------------------------------------------------------
#
left, width = 0.0, .19
bottom, height = .5, 1
right = left + width
top = 1
fig = plt.figure(figsize=(16, 4))
ax = fig.add_axes([0, 0, 1, 1])
# ---------- Remove border from the figures  ------------------
for item in [fig, ax]:
    item.patch.set_visible(False)
fig.patch.set_visible(False)
ax.axis('off')
# -------------------------------------------------------------
p = patches.Rectangle(
    (left, bottom), width, height,
    color='#fcea17'
)
ax.add_patch(p)


def human_format(number):
    units = ['', 'K', 'M', 'B', 'T', 'P']
    k = 1000.0
    magnitude = int(floor(log(number, k)))
    return '%.1f %s ' % (number / k ** magnitude, units[magnitude])


# -------- Last Day Return  ----------------------------------------------
lastDaySales = pd.read_sql_query("""
                select  Sum(EXTINVMISC) as  LastDaySales from OESalesDetails
                where TRANSDATE = convert(varchar(8),DATEADD(D,0,GETDATE()-1),112)
                and AUDTORG='RNGSKF'
                                """, connection)
LastDayReturn = pd.read_sql_query("""
                select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
                where AUDTORG= 'RNGSKF' and transtype<>1 and PRICELIST <> 0 and
                (TRANSDATE = convert(varchar(8),DATEADD(D,0,GETDATE()-1),112))
                 """, connection)
ld_sales = int(lastDaySales['LastDaySales'])
ld_return = float(LastDayReturn['ReturnAmount'])
l_retn = abs(ld_return)
ret1 = float((l_retn / ld_sales) * 100)
return_p = '%.2f' % (ret1)
return_p = str(return_p) + '%'
kpi_label = 'LD' + "\n"
ax.text(.5 * (left + right), .5 * (bottom + top), kpi_label,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=24, color='black',
        transform=ax.transAxes)
ax.text(.5 * (left + right), .4 * (bottom + top), return_p,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=34, color='red',
        transform=ax.transAxes)
print('Last Day Return Added')
# --------------------  Monthly  Return  Box --------------------------------
left, width = .20, .19
bottom, height = .5, .5
right = left + width
top = 1
p = patches.Rectangle(
    (left, bottom), width, height,
    color='#fcea17'
)
ax.add_patch(p)
monthly_sales = pd.read_sql_query(""" Declare @monthStartDay NVARCHAR(MAX);
                Declare @monthCurrentDay NVARCHAR(MAX);
                SET @monthStartDay = convert(varchar(8),DATEADD(month, DATEDIFF(month, 0,  GETDATE()), 0),112)
                set @monthCurrentDay = convert(varchar(8),DATEADD(D,0,GETDATE()),112);
                select  Sum(EXTINVMISC) as  MTDSales from OESalesDetails
                where TRANSDATE between  @monthStartDay and @monthCurrentDay
                and AUDTORG='RNGSKF'
                """, connection)
monthly_return_df = pd.read_sql_query("""select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
        where AUDTORG= 'RNGSKF' and transtype<>1 and PRICELIST <> 0 and
        (TRANSDATE between
        (convert(varchar(8),DATEADD(mm, DATEDIFF(mm, 0, GETDATE()), 0),112))
        and (convert(varchar(8),DATEADD(D,0,GETDATE()),112)))
                    """, connection)
monthly_return = float(monthly_return_df['ReturnAmount'])
m_sales = int(monthly_sales['MTDSales'])
retn = abs(monthly_return)
ret1 = float((retn / m_sales) * 100)
return_p = '%.2f' % (ret1)
return_p = str(return_p) + '%'
kpi_label = 'MTD' + "\n"
ax.text(.5 * (left + right), .5 * (bottom + top), kpi_label,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=24, color='black',
        transform=ax.transAxes)
ax.text(.5 * (left + right), .4 * (bottom + top), return_p,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=34, color='red',
        transform=ax.transAxes)
print('MTD Return Added')
# # ---------- Yearly return Box ------------------------
yearly_sales = pd.read_sql_query("""
            select  Sum(EXTINVMISC) as  YTDSales from OESalesDetails
            where AUDTORG='RNGSKF' AND TRANSDATE>= (convert(varchar(8),DATEADD(yy, DATEDIFF(yy, 0, GETDATE()), 0),112))
            """, connection)
yearly_return = pd.read_sql_query("""
                  select sum(EXTINVMISC) as ReturnAmount from OESalesDetails where
                AUDTORG = 'RNGSKF' and
                transtype<>1 and PRICELIST <> 0 and
                (TRANSDATE between (convert(varchar(8),DATEADD(yy, DATEDIFF(yy, 0, GETDATE()), 0),112))
                and (convert(varchar(8),DATEADD(D,0,GETDATE()),112)))
 """, connection)
left, width = .40, .19
bottom, height = .5, .5
right = left + width
top = 1
y_sales = int(yearly_sales['YTDSales'])
yearly_return_amount = int(yearly_return['ReturnAmount'])
return_amount = abs(yearly_return_amount)
return_p = float((return_amount / y_sales) * 100)
return_p = '%.2f' % (return_p)
return_p = str(return_p) + '%'
p = patches.Rectangle(
    (left, bottom), width, height,
    color='#fcea17'
)
ax.add_patch(p)
kpi_label = 'YTD' + "\n"
ax.text(.5 * (left + right), .5 * (bottom + top), kpi_label,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=24, color='black',
        transform=ax.transAxes)
ax.text(.5 * (left + right), .4 * (bottom + top), return_p,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=34, color='red',
        transform=ax.transAxes)
print('YTD Return Added')

# # ---------- YAGO MTD  Return Box ------------------------
yago_monthly_sales = pd.read_sql_query("""Declare @YagoMonthStartDay NVARCHAR(MAX);
                Declare @YagomonthCurrentDay NVARCHAR(MAX);
                SET @YagoMonthStartDay = convert(varchar(6), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112)
                set @YagomonthCurrentDay = convert(varchar(8), DATEADD(year, -1, GETDATE()), 112)
                select  Sum(EXTINVMISC) as  MTDSales from OESalesDetails
                where TRANSDATE between  @YagoMonthStartDay and @YagomonthCurrentDay
                and AUDTORG='RNGSKF'
                 """, connection)
yago_monthly_return_df = pd.read_sql_query("""select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
where AUDTORG= 'RNGSKF' and transtype<>1 and PRICELIST <> 0 and
transdate between (convert(varchar(6), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112))
                    """, connection)
monthly_return = float(yago_monthly_return_df['ReturnAmount'])
m_sales = int(yago_monthly_sales['MTDSales'])
left, width = .60, .19
bottom, height = .5, .5
right = left + width
top = 1
p = patches.Rectangle(
    (left, bottom), width, height,
    color='#fc8b17'
)
ax.add_patch(p)
retn = abs(monthly_return)
ret1 = float((retn / m_sales) * 100)
return_p = '%.2f' % (ret1)
return_p = str(return_p) + '%'
kpi_label = 'YAGO MTD' + "\n"
ax.text(.5 * (left + right), .5 * (bottom + top), kpi_label,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=24, color='black',
        transform=ax.transAxes)
ax.text(.5 * (left + right), .4 * (bottom + top), return_p,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=34, color='black',
        transform=ax.transAxes)
print('YAGO MTD Return Added')
# # # ---------- YAGO YTD Return  Box ------------------------
yago_yearly_sales = pd.read_sql_query("""
                select  Sum(EXTINVMISC) as  YTDSales from OESalesDetails
                where AUDTORG='RNGSKF' AND TRANSDATE between (convert(varchar(8), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
                and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112))
                """, connection)
yago_yearly_return = pd.read_sql_query("""
                select sum(EXTINVMISC) as ReturnAmount from OESalesDetails where
                AUDTORG = 'RNGSKF' and
                transtype<>1 and PRICELIST <> 0 and
                (TRANSDATE between (convert(varchar(8), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
                and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112)))
                """, connection)
left, width = .80, .20
bottom, height = .5, .5
right = left + width
top = 1
y_sales = int(yago_yearly_sales['YTDSales'])
yearly_return_amount = int(yago_yearly_return['ReturnAmount'])
return_amount = abs(yearly_return_amount)
return_p = float((return_amount / y_sales) * 100)
return_p = '%.2f' % (return_p)
return_p = str(return_p) + '%'
p = patches.Rectangle(
    (left, bottom), width, height,
    color='#fc8b17'
)
ax.add_patch(p)
kpi_label = 'YAGO YTD' + "\n"
ax.text(.5 * (left + right), .5 * (bottom + top), kpi_label,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=24, color='black',
        transform=ax.transAxes)
ax.text(.5 * (left + right), .4 * (bottom + top), return_p,
        horizontalalignment='center',
        verticalalignment='center',
        fontsize=34, color='black',
        transform=ax.transAxes)
# plt.tight_layout()
plt.savefig('./return.png')
# # plt.show()
dirpath = os.path.dirname(os.path.realpath(__file__))
im = Image.open(dirpath + '/return.png')
# im.show()
left = 0
top = 0
right = 1600
bottom = 200
im1 = im.crop((left, top, right, bottom))
im1.save('return.png')
# im1.show()
print('Return Generated ')

# # Join Title and Results
title_img = Image.open(dirpath + "/return_text.png")
return_img = Image.open(dirpath + "/return.png")
dst = Image.new('RGB', (1602, 301))
dst.paste(title_img, (0, 0))
dst.paste(return_img, (1, title_img.height))
dst.save('main_return.png')

# ------------------------Cause wise return--------------------------

cause_wise_return_df = pd.read_sql_query("""select case
                when Cause_Of_Return_ID = '000'  THEN 'Not Mentioned'
                when Cause_Of_Return_ID = '005'  THEN 'Product Short'
				when Cause_Of_Return_ID = '010'  THEN 'Shop Closed'
                when Cause_Of_Return_ID = '015'  THEN 'Canceled/Cash Short'
				when Cause_Of_Return_ID = '020'  THEN 'Computer Mistake'
                when Cause_Of_Return_ID = '025'  THEN 'Next Day Delivery'
				when Cause_Of_Return_ID = '030'  THEN 'Part Sale'
                when Cause_Of_Return_ID = '035'  THEN 'Not Ordered'
				when Cause_Of_Return_ID = '040'  THEN 'Not Delivered'
                when Cause_Of_Return_ID = '050'  THEN 'Approved Return'
				when Cause_Of_Return_ID = '065'  THEN 'MSO Mistake'
				End AS Cause
				, ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
				where AUDTORG= 'RNGSKF' and transtype<>1 and PRICELIST <> 0 and
				TRANSDATE between
				(convert(varchar(8),DATEADD(mm, DATEDIFF(mm, 0, GETDATE()), 0),112))
				and (convert(varchar(8),DATEADD(D,0,GETDATE()),112))
				group by Cause_Of_Return_ID""", connection)

Cause_name = cause_wise_return_df['Cause']
y_pos = np.arange(len(Cause_name))
Return_amount = abs(cause_wise_return_df['ReturnAmount'])
Return_amount = Return_amount.values.tolist()

total = sum(Return_amount)

for_looping = 0
new_return_amount = []
for some_value in Return_amount:
    changed_values = (Return_amount[for_looping] / total) * 100
    new_return_amount.insert(for_looping, changed_values)
    for_looping = for_looping + 1
print('Updated list :', new_return_amount)

color = ['#1ff015', '#418af2', '#f037d9', '#ecc13f', '#e2360a', '#1ff015', '#418af2', '#f037d9', '#ecc13f', '#e2360a']
width = 0.75

plt.figure(figsize=(2, 1))
fig, ax = plt.subplots()
rects1 = plt.bar(y_pos, new_return_amount, width, align='center', alpha=0.9, color=color)


def autolabel(bars):
    loop = 0
    for bar in bars:
        show = float((Return_amount[loop] / total) * 100)
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2., 1 * height,
                '%.2f' % (show) + '%',
                ha='center', va='bottom', fontsize=12, fontweight='bold')
        loop = loop + 1


autolabel(rects1)
# plt.yticks(y_pos, name)
plt.xticks(y_pos, Cause_name, fontsize='12', color='black')
plt.yticks(np.arange(0, 101, 10), color='black', fontsize='12')
plt.xlabel('Return Cause', fontsize='14', color='black', fontweight='bold')
plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
plt.title('10. Cause wise Return', color='#3e0a75', fontsize='16', fontweight='bold')
plt.tight_layout()
# plt.gca().invert_yaxis()
# fig.set_size_inches(12.81, 4.8)
plt.savefig('Cause_with_return.png')
# plt.show()
print('10. cause wise return generated')

# ----------------------Delivery man wise return---------------------

delivery_man_wise_return_df = pd.read_sql_query("""
    select TWO.ShortName as DPNAME, Sales.ReturnAmount from
    (select  LEFT(DPID,5) AS DPID, AUDTORG,ISNULL(sum(case when TRANSTYPE<>1 then INVNETH *-1 end), 0) /ISNULL(sum(case when TRANSTYPE=1 then INVNETH end), 0)*100 as ReturnAmount from OESalesSummery
            where AUDTORG= 'RNGSKF' and 
            left(TRANSDATE,6)=convert(varchar(6),getdate(),112)
            group by LEFT(DPID,5),AUDTORG) as Sales
    left join
    (select   distinct AUDTORG,ShortName, LEFT(DPID,5) as DPID from DP_ShortName where AUDTORG= 'RNGSKF') as TWO
    on Sales.DPID = TWO.DPID
    and Sales.AUDTORG=TWO.AUDTORG
    order by ReturnAmount DESC""", connection)

DPNAME = delivery_man_wise_return_df['DPNAME']
y_pos = np.arange(len(DPNAME))
ReturnAmount = abs(delivery_man_wise_return_df['ReturnAmount'])
ReturnAmount = ReturnAmount.values.tolist()

# total = sum(Return_amount)  # ------it is taken from cause wise return--------
# totalof5 = sum(ReturnAmount)
#
# for_looping = 0
# new_return_amount = []
# for some_value in ReturnAmount:
#     changed_values = (ReturnAmount[for_looping] / total) * 100
#     new_return_amount.insert(for_looping, changed_values)
#     for_looping = for_looping + 1
# print('Updated list :', new_return_amount)

ran=max(ReturnAmount)
print(round(ran)+2)
#sys.exit()
color = '#418af2'
# width = 0.75

# plt.figure(figsize=(2, 1))
fig, ax = plt.subplots(figsize=(12.81, 4.8))
rects1 = plt.bar(y_pos, ReturnAmount, align='center', alpha=0.9, color=color)


def autolabel(bars):
    loop = 0
    for bar in bars:
        # show = float((ReturnAmount[loop] / total) * 100)
        show = ReturnAmount[loop]
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, (1.05 * height),
                '%.2f' % (show) + '%', ha='center', va='bottom', fontsize=12, rotation=70, fontweight='bold')
        loop = loop + 1


autolabel(rects1)

# plt.yticks(y_pos, name)
plt.xticks(y_pos, DPNAME, rotation='vertical', fontsize='12')
plt.yticks(np.arange(0, round(ran)+(.6*round(ran))), fontsize='12')
plt.xlabel('Delivery Person', fontsize='14', color='black', fontweight='bold')
plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
plt.title("9. Delivery Person's Return %", color='#3e0a75', fontsize='16', fontweight='bold')
plt.tight_layout()
# plt.gca().invert_yaxis()
# fig.set_size_inches(12.81, 4.8)
plt.savefig('Delivery_man_wise_return.png')
# plt.show()
print('12. Delivery man wise return generated')

# ------------adding cause wise return and delivery man wise return----------

imp20 = Image.open(dirpath + "/Cause_with_return.png")
widthx, heightx = imp20.size
imp21 = Image.open(dirpath + "/LD_MTD_YTD_TARGET_vs_sales.png")
imageSize = Image.new('RGB', (1283, 482))
imageSize.paste(imp20, (1, 1))
imageSize.paste(imp21, (widthx + 2, 1))
imageSize.save(dirpath + "/Cause_wise_delivery_man_wise_return.png")

# --------------- Close TO Mature Credit Redords -------------------


# ------------adding cause wise return and delivery man wise return----------

imp22 = Image.open(dirpath + "/Delivery_man_wise_return.png")
widthx, heightx = imp22.size
imageSize = Image.new('RGB', (1283, 482))
imageSize.paste(imp22, (1, 1))
imageSize.save(dirpath + "/new_total_delivery_man_wise_return.png")

# ------------adding cause wise return and delivery man wise return----------

imp23 = Image.open(dirpath + "/Day_Wise_Target_vs_Sales.png")
widthx, heightx = imp23.size
imageSize = Image.new('RGB', (1283, 482))
imageSize.paste(imp23, (1, 1))
imageSize.save(dirpath + "/Day_wise_target_sales.png")

# --------------- Close TO Mature Credit Redords -------------------


ClosedToMaturedcredit_df = pd.read_sql_query("""
               select CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',
CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',INVNUMBER as 'Inv Number',
convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date', CustomerInformation.CREDIT_LIMIT_DAYS as 'Days Limit',
(datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)*-1 as 'Matured In Days', OUT_NET as 'Credit Amount'
from [ARCOUT].dbo.[CUST_OUT]
join ARCHIVESKF.dbo.CustomerInformation
on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

where  [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash'
	and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)  between -3 and 0
	order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)desc
	, OUT_NET desc

 """, connection)
writer = pd.ExcelWriter('ClosedToMatured.xlsx', engine='xlsxwriter')
ClosedToMaturedcredit_df.index = np.arange(1, len(ClosedToMaturedcredit_df) + 1)
ClosedToMaturedcredit_df.to_excel(writer, sheet_name='Sheet1', index=True)
workbook = writer.book

worksheet = writer.sheets['Sheet1']
worksheet.set_column('A:A', 4)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 17)
worksheet.set_column('G:G', 14)
worksheet.set_column('H:H', 18)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 20)

writer.save()
print('Closed to Matured : Excel Created')

ClosedToMaturedcredittable = pd.read_sql_query("""
  select top 20 CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',
CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',INVNUMBER as 'Inv Number',
convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date', CustomerInformation.CREDIT_LIMIT_DAYS as 'Days Limit',
(datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)*-1 as 'Matured In Days', OUT_NET as 'Credit Amount'
from [ARCOUT].dbo.[CUST_OUT]
join ARCHIVESKF.dbo.CustomerInformation
on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

where  [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash'
	and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)  between -3 and 0
	order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)desc
	, OUT_NET desc
               """, connection)
writer = pd.ExcelWriter('ClosedToMaturedTable.xlsx', engine='xlsxwriter')
ClosedToMaturedcredittable.index = np.arange(1, len(ClosedToMaturedcredittable) + 1)
ClosedToMaturedcredittable.to_excel(writer, sheet_name='Sheet1', index=True)

workbook = writer.book
worksheet = writer.sheets['Sheet1']
worksheet.set_column('A:A', 4)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 17)
worksheet.set_column('G:G', 14)
worksheet.set_column('H:H', 18)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 20)

writer.save()
print('Closed to Matured Mail Table  : Excel Created')


# --------------Functions ---------------
def get_html_table():
    xl = xlrd.open_workbook('ClosedToMaturedTable.xlsx')
    sh = xl.sheet_by_name('Sheet1')

    th = ""
    td = ""
    for i in range(0, 1):
        th = th + "<tr>\n"
        th = th + "<th class=\"unit\">ID</th>"
        for j in range(1, sh.ncols):
            th = th + "<th class=\"unit\" >" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    for i in range(1, sh.nrows):
        td = td + "<tr>\n"
        td = td + "<td class=\"idcol\">" + str(i) + "</td>"
        for j in range(1, 2):
            td = td + "<td class=\"idcol\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(2, 7):
            td = td + "<td class=\"unit\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(7, sh.ncols):
            td = td + "<td class=\"idcol\">" + get_value(str(int(sh.cell_value(i, j)))) + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    return html


AgeingMaturedcredit_df = pd.read_sql_query("""
               select CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',
                INVNUMBER as 'Inv Number', convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date',
                 CustomerInformation.CREDIT_LIMIT_DAYS as 'Days Limit',
                datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as 'Days Passed',OUT_NET as 'Credit Amount'

                from [ARCOUT].dbo.[CUST_OUT]
                join ARCHIVESKF.dbo.CustomerInformation
                on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

                where  [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash' and OUT_NET>1
                and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) >= 1
                order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) desc
                                , OUT_NET desc
                 """, connection)

writer = pd.ExcelWriter('AgingMatured.xlsx', engine='xlsxwriter')
AgeingMaturedcredit_df.index = np.arange(1, len(AgeingMaturedcredit_df) + 1)
AgeingMaturedcredit_df.to_excel(writer, sheet_name='Sheet1', index=True)

workbook = writer.book
worksheet = writer.sheets['Sheet1']
worksheet.set_column('A:A', 4)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 17)
worksheet.set_column('G:G', 14)
worksheet.set_column('H:H', 18)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 20)
writer.save()
print('Aging Matured : Excel Created')

AgeingMaturedcredittable = pd.read_sql_query("""
               select top 20 CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',
                 INVNUMBER as 'Inv Number', convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date',
                 CustomerInformation.CREDIT_LIMIT_DAYS as 'Days Limit',
                datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as 'Days Passed'
                ,OUT_NET as 'Credit Amount'
                from [ARCOUT].dbo.[CUST_OUT]
                join ARCHIVESKF.dbo.CustomerInformation
                on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

                where  [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS<>'Cash' and OUT_NET>1
                and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) >= 1
                order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) desc
                 , OUT_NET desc
                """, connection)

writer = pd.ExcelWriter('AgingMaturedTable.xlsx', engine='xlsxwriter')
AgeingMaturedcredittable.index = np.arange(1, len(AgeingMaturedcredittable) + 1)
AgeingMaturedcredittable.to_excel(writer, sheet_name='Sheet1', index=True)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
writer.save()


# --------------Functions ---------------
def get_html_table1():
    xl = xlrd.open_workbook('AgingMaturedTable.xlsx')
    sh = xl.sheet_by_name('Sheet1')

    th = ""
    td = ""
    for i in range(0, 1):
        th = th + "<tr>\n"
        th = th + "<th class=\"unit\">ID</th>"
        for j in range(1, sh.ncols):
            th = th + "<th class=\"unit\" >" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    for i in range(1, sh.nrows):
        td = td + "<tr>\n"
        td = td + "<td class=\"idcol\">" + str(i) + "</td>"
        for j in range(1, 2):
            td = td + "<td class=\"idcol\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(2, 7):
            td = td + "<td class=\"unit\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(7, sh.ncols):
            td = td + "<td class=\"idcol\">" + get_value(str(int(sh.cell_value(i, j)))) + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    return html


CashDrop_df = pd.read_sql_query("""
             Select CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',
             INVNUMBER as 'Inv Number',convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date',
            datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 as 'Days Over', OUT_NET as 'Credit Amount'
            from [ARCOUT].dbo.[CUST_OUT]
            join ARCHIVESKF.dbo.CustomerInformation
            on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

            where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS='Cash'  and OUT_NET>1
            and (datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1) >=4
            order by datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 desc
              , OUT_NET desc
             """, connection)
writer = pd.ExcelWriter('CashDrop.xlsx', engine='xlsxwriter')
CashDrop_df.index = np.arange(1, len(CashDrop_df) + 1)
CashDrop_df.to_excel(writer, sheet_name='Sheet1', index=True)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
worksheet.set_column('A:A', 4)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 17)
worksheet.set_column('G:G', 14)
worksheet.set_column('H:H', 18)
worksheet.set_column('I:I', 20)
worksheet.set_column('J:J', 20)
writer.save()
print('Cash Drop : Excel Created')

CashDroptable_df = pd.read_sql_query("""
        Select top 20 CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',
         INVNUMBER as 'Inv Number',convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date',
        datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 as 'Days Over', OUT_NET as 'Credit Amount'
        from [ARCOUT].dbo.[CUST_OUT]
        join ARCHIVESKF.dbo.CustomerInformation
        on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

        where [ARCOUT].dbo.[CUST_OUT].AUDTORG = 'RNGSKF' and TERMS='Cash'  and OUT_NET>1
        and (datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1) >=4
        order by datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 desc
         , OUT_NET desc
        """, connection)

writer = pd.ExcelWriter('CashDropTable.xlsx', engine='xlsxwriter')
CashDroptable_df.index = np.arange(1, len(CashDroptable_df) + 1)
CashDroptable_df.to_excel(writer, sheet_name='Sheet1', index=True)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
writer.save()


# --------------Functions ---------------
def get_html_table2():
    xl = xlrd.open_workbook('CashDropTable.xlsx')
    sh = xl.sheet_by_name('Sheet1')
    th = ""
    td = ""
    for i in range(0, 1):
        th = th + "<tr>\n"
        th = th + "<th class=\"unit\">ID</th>"
        for j in range(1, sh.ncols):
            th = th + "<th class=\"unit\" >" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    for i in range(1, sh.nrows):
        td = td + "<tr>\n"
        td = td + "<td class=\"idcol\">" + str(i) + "</td>"
        for j in range(1, 2):
            td = td + "<td class=\"idcol\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(2, 7):
            td = td + "<td class=\"unit\">" + str(sh.cell_value(i, j)) + "</td>\n"
        for j in range(7, sh.ncols):
            td = td + "<td class=\"idcol\">" + get_value(str(int(sh.cell_value(i, j)))) + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    return html


all_table = """ <!DOCTYPE html>
                        <html lang="en">
                        <head>
                            <meta charset="UTF-8">
                            <style>
                                    table
                                    {   
                                        width: 796px;
                                        border-collapse: collapse;
                                        border: 1px solid gray;
                                        padding: 5px;
                                    }
                                    td{
                                        padding-top: 5px;
                                        border-bottom: 1px solid #ddd;
                                        text-align: left;
                                        white-space: nowrap;
                                        border: 1px solid gray;
                                        text-align: justify;
                                        text-justify: inter-word;

                                    }
                                    th{
                                        padding: 2px;
                                        border: 1px solid gray;
                                        font-size: 10px !important;
                                        text-align:left;
                                        white-space: nowrap;
                                        text-align: justify;
                                        text-justify: inter-word;

                                    }
                                    th.table30tr{
                                        font-size:15px !important;
                                        height: 24px !important;
                                        background-color: brown;
                                        color: white;
                                        text-align: left;
                                        white-space: nowrap;
                                    }

                                    th.unit{

                                        background-color: #5ef28d;
                                        width:22px;
                                        font-size: 10px;
                                        white-space: nowrap;
                                        text-align: justify;
                                        text-justify: inter-word;

                                    }
                                    td.idcol{
                                        text-align: right;
                                        background-color: #efedf2;
                                        white-space: nowrap;
                                        font-size: 9px;
                                        text-justify: inter-word;
                                    }
                                    td.unit{
                                        background-color: #efedf2;
                                        white-space: nowrap;
                                        font-size: 9px;
                                        text-align: justify;
                                        text-justify: inter-word;

                                    }
                                    tr:hover {
                                        background-color:#f5f5f5;
                                    }
                                </style>
                            <title>Mail</title>
                        </head>
                        <body>
                        <h3 style='text-align:left'> Top 20 Closed to Mature Credit</h3>
                        <table>
                            <p style="text-align:left"> """ + get_html_table() + """</p>
                        </table>

                        <h3 style='text-align:left'> Top 20 Aging Mature Credit</h3>
                        <table>
                            <p style="text-align:left"> """ + get_html_table1() + """</p>
                        </table>

                         <h3 style='text-align:left'> Top 20 Cash Drop Credit</h3>
                        <table>
                            <p style="text-align:left"> """ + get_html_table2() + """</p>
                        </table>
                        </body>
                        </html>
                        """

# -----------------------------------------------------------------
# ------------ Email Section --------------------------------------
# -----------------------------------------------------------------
# msgRoot = MIMEMultipart('related')
# me = 'erp-bi.service@transcombd.com'
# to = 'fazle.rabby@transcombd.com'
# recipient = to

# ------------ Group email ---------------------------------------
msgRoot = MIMEMultipart('related')
me = 'erp-bi.service@transcombd.com'
# to = ['biswas@transcombd.com', 'yakub@transcombd.com']
# cc = ['zubair.transcom@gmail.com', 'aftab.uddin@transcombd.com']
# bcc = ['rejaul.islam@transcombd.com', 'fazle.rabby@transcombd.com']
to = ['fazle.rabby@transcombd.com']
#cc = ['rejaul.islam@transcombd.com']
#bcc = ['shamiul.islam@transcombd.com']

#recipient = to+cc+bcc
recipient = to

date = datetime.today()
today = str(date.day) + '/' + str(date.month) + '/' + str(date.year) + ' ' + date.strftime("%I:%M %p")
today1 = str(date.day) + ' ' + str(date.strftime("%B")) + ' ' + str(date.year) + ' at ' + date.strftime("%I:%M %p")

subject = "SK+F Formulation Reports - " + today

email_server_host = 'mail.transcombd.com'
port = 25

# msgRoot['to'] = recipient
# msgRoot['from'] = me
# msgRoot['subject'] = subject
msgRoot['From'] = me
msgRoot['To'] = ', '.join(to)
#msgRoot['Cc'] = ', '.join(cc)
#msgRoot['Bcc'] = ', '.join(bcc)
msgRoot['Subject'] = subject

msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

# We reference the image in the IMG SRC attribute by the ID we give it below

# html = """
#                     <img src="cid:banner" height='260' width='796'> <br>
#                    <img src="cid:all_credit"  width='796'><br>
#                    <img src="cid:all_Matured_credit" width='796'> <br>
#                    <img src="cid:all_regular_credit" width='796'> <br>
#                    <img src="cid:all_cash"  width='796'> <br>
#                    <img src="cid:main_return" width='796'> <br>
#                     <img src="cid:return_with_cause" width='796'>
#                     <div style="text-align:left" width='796px'> """ + all_table + """</div>
#
#                    If there is any inconvenience, you are requested to communicate with the ERP BI Service:
#                    <br><b>(Mobile: 01713-389972, 01713-380499)</b><br><br>
#                    Regards<br><b>ERP BI Service</b><br>Information System Automation (ISA)<br><br>
#                    <i><font color="blue">****This is a system generated stock report ****</i></font>"""
msgText = MIMEText("""
                    <img src="cid:banner" height='200' width='796'> <br>
                   <img src="cid:all_credit"  width='796'><br>
                   <img src="cid:all_Matured_credit"  width='796'> <br>
                   <img src="cid:all_regular_credit"  width='796'> <br>
                   <img src="cid:all_cash"  width='796'> <br>
                   <img src="cid:main_return" width='796'> <br>

                   <img src="cid:delivery_mans_return_with_cause"  width='796'>
                    <img src="cid:return_with_cause"  width='796'>
                   <img src="cid:day_wise_sales_target"  width='796'>
                     """ + all_table + """

                    <br>
                   If there is any inconvenience, you are requested to communicate with the ERP BI Service:
                   <br><b>(Mobile: 01713-389972, 01713-380499)</b><br><br>
                   Regards<br><b>ERP BI Service</b><br>Information System Automation (ISA)<br><br>
                   <i><font color="blue">****This is a system generated stock report ****</i></font>""", 'html')

msgAlternative.attach(msgText)

# --------- Set Credit image in mail   -----------------------
fp = open(dirpath + '/banner_ai.png', 'rb')
banner = MIMEImage(fp.read())
fp.close()

banner.add_header('Content-ID', '<banner>')
msgRoot.attach(banner)

# --------- Set Credit image in mail   -------------------------
fp = open(dirpath + '/all_credit.png', 'rb')
credit = MIMEImage(fp.read())
fp.close()

credit.add_header('Content-ID', '<all_credit>')
msgRoot.attach(credit)

# --------- Set Matured Credit image in mail   -------------------------
fp = open(dirpath + '/all_Matured_credit.png', 'rb')
all_Matured_credit = MIMEImage(fp.read())
fp.close()

all_Matured_credit.add_header('Content-ID', '<all_Matured_credit>')
msgRoot.attach(all_Matured_credit)

# --------- Set Matured Credit image in mail   -------------------------
fp = open(dirpath + '/all_regular_credit.png', 'rb')
all_regular_credit = MIMEImage(fp.read())
fp.close()

all_regular_credit.add_header('Content-ID', '<all_regular_credit>')
msgRoot.attach(all_regular_credit)

# --------- Set Cash Drop image in mail   -------------------------
fp = open(dirpath + '/all_cash.png', 'rb')
all_cash = MIMEImage(fp.read())
fp.close()

all_cash.add_header('Content-ID', '<all_cash>')
msgRoot.attach(all_cash)

# --------- Set Return image in mail   -------------------------
fp = open(dirpath + '/main_return.png', 'rb')
main_return = MIMEImage(fp.read())
fp.close()

main_return.add_header('Content-ID', '<main_return>')
msgRoot.attach(main_return)

# --------set Cause_wise_delivery_man_wise_return in mail-----------

fp = open(dirpath + '/Cause_wise_delivery_man_wise_return.png', 'rb')
return_with_cause = MIMEImage(fp.read())
fp.close()

return_with_cause.add_header('Content-ID', '<return_with_cause>')
msgRoot.attach(return_with_cause)

# ---------------------------------------------------------------------

fp = open(dirpath + '/new_total_delivery_man_wise_return.png', 'rb')
return_with_cause = MIMEImage(fp.read())
fp.close()

return_with_cause.add_header('Content-ID', '<delivery_mans_return_with_cause>')
msgRoot.attach(return_with_cause)

# ---------------------------------------------------------------------

fp = open(dirpath + '/Day_wise_target_sales.png', 'rb')
return_with_cause = MIMEImage(fp.read())
fp.close()

return_with_cause.add_header('Content-ID', '<day_wise_sales_target>')
msgRoot.attach(return_with_cause)

# Closed to matured Excel attachment
part = MIMEBase('application', "octet-stream")
file_location = dirpath + '/ClosedToMatured.xlsx'

# Create the attachment file (only do it once)
filename = os.path.basename(file_location)
attachment = open(file_location, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msgRoot.attach(part)

# Aging matured Excel attachment--------------
part = MIMEBase('application', "octet-stream")
file_location = dirpath + '/AgingMatured.xlsx'
# Create the attachment file (only do it once)
filename = os.path.basename(file_location)
attachment = open(file_location, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msgRoot.attach(part)

# Cash Drop  Excel attachment--------------
part = MIMEBase('application', "octet-stream")
file_location = dirpath + '/CashDrop.xlsx'
# Create the attachment file (only do it once)
filename = os.path.basename(file_location)
attachment = open(file_location, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msgRoot.attach(part)

# -------- Decleration group mail -----------------------------
# msgRoot['From'] = me
# msgRoot['To'] = ', '.join(to)
# msgRoot['Cc'] = ', '.join(cc)
# msgRoot['Bcc'] = ', '.join(bcc)
# msgRoot['Subject'] = subject


# ----------- Finally send mail and close server connection ---
server = smtplib.SMTP(email_server_host, port)
server.ehlo()
print('\n-----------------')
print('Sending Mail')
server.sendmail(me, recipient, msgRoot.as_string())
print('Mail Send')
print('-------------------')
server.close()

# Html_file = open("Test.html", "w")
# Html_file.write(html)
# Html_file.close()

from datetime import datetime
import pytz

tz_NY = pytz.timezone('Asia/Dhaka')
datetime_BD = datetime.now(tz_NY)
print("Execution time:", datetime_BD.strftime("%I:%M %p"))
import winsound

winsound.Beep(1000, 500)
