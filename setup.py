import os

dirpath = os.path.dirname(os.path.realpath(__file__))

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
import sys

Six_branches_list = ['BSLSKF', 'SYLSKF', 'CTGSKF', 'HZJSKF', 'FRDSKF', 'MHKSKF']  # set the branch names here

# for each_branch in Six_branches_list:
#     branc_names = each_branch
#     print(branc_names)

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

branc_names = 'BSLSKF'

if branc_names == 'BSLSKF':
    branch_name = 'Barisal'

if branc_names == 'SYLSKF':
    branch_name = 'Shylet'

if branc_names == 'CTGSKF':
    branch_name = 'Chittagong-South'

if branc_names == 'HZJSKF':
    branch_name = 'Chandpur'

if branc_names == 'FRDSKF':
    branch_name = 'Faridpur'

if branc_names == 'MHKSKF':
    branch_name = 'Mohakhali'

tag.text((25, 8), 'SK+F', (255, 255, 255), font=font)
branch.text((25, 270), branch_name + " Branch", (255, 209, 0), font=font1)
timestore.text((25, 435), time + "\n" + day, (255, 255, 255), font=font2)
img.save('banner_ai.png')


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


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=137.116.139.217;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp@123')

cursor = connection.cursor()


def human_format(number):
    units = ['', 'K', 'M', 'B', 'T', 'P']
    k = 1000.0
    magnitude = int(floor(log(number, k)))
    return '%.1f %s ' % (number / k ** magnitude, units[magnitude])


outstanding_df = pd.read_sql_query(""" select
                SUM(CASE WHEN TERMS='CASH' THEN OUT_NET END) AS TotalOutStandingOnCash,
                SUM(CASE WHEN TERMS not like '%CASH%' THEN OUT_NET END) AS TotalOutStandingOnCredit

                from  [ARCOUT].dbo.[CUST_OUT]
                where AUDTORG like ? AND [INVDATE] <= convert(varchar(8),DATEADD(D,0,GETDATE()),112)
                                        """, connection, params={branc_names})

cash = int(outstanding_df['TotalOutStandingOnCash'])
credit = int(outstanding_df['TotalOutStandingOnCredit'])

data = [cash, credit]
total = cash + credit
total = 'Total \n' + joker(total)

colors = ['#f9ff00', '#ff8600']

legend_element = [Patch(facecolor='#f9ff00', label='Cash'),
                  Patch(facecolor='#ff8600', label='Credit')]

# -------------------new code--------------------------

ca = joker(cash)
cre = joker(credit)

DataLabel = [ca, cre]
# -----------------------------------------------------

fig1, ax = plt.subplots()
wedges, labels, autopct = ax.pie(data, colors=colors, labels=DataLabel, autopct='%.1f%%', startangle=90,
                                 pctdistance=.7)
plt.setp(autopct, fontsize=14, color='black', fontweight='bold')
plt.setp(labels, fontsize=14, fontweight='bold')
ax.text(0, -.1, total, ha='center', fontsize=14, fontweight='bold', backgroundcolor='#00daff')

centre_circle = plt.Circle((0, 0), 0.50, fc='white')

fig = plt.gcf()

fig.gca().add_artist(centre_circle)
# Equal aspect ratio ensures that pie is drawn as a circle
plt.title('1. Total Outstanding', fontsize=16, fontweight='bold', color='#3e0a75')

ax.axis('equal')
plt.legend(handles=legend_element, loc='lower left',
           fontsize=11)
plt.tight_layout()
plt.savefig('terms_wise_outstanding.png')

print('1. Terms wise Outstanding Generated')
plt.close()

# #--------- Category wise Credit ( Matured , Not Mature) -------------
credit_category_df = pd.read_sql_query(""" Select case when Days_Diff>0 then 'Matured Credit' else 'Regular Credit' End as  'Category',Sum(OUT_NET) as Amount from
                        (select INVNUMBER,INVDATE,
                        CUSTOMER,TERMS,MAINCUSTYPE,
                        CustomerInformation.CREDIT_LIMIT_DAYS,
                        datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                        OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                        join ARCHIVESKF.dbo.CustomerInformation
                        on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                        where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash') as TblCredit
                        group by case when Days_Diff>0 then 'Matured Credit' else 'Regular Credit' end
                                                            """, connection, params={branc_names})

matured = int(credit_category_df.Amount.iloc[0])
not_mature = int(credit_category_df.Amount.iloc[1])

values = [matured, not_mature]

colors = ['#ffb667', '#b35e00']

legend_element = [Patch(facecolor='#ffb667', label='Matured'),
                  Patch(facecolor='#b35e00', label='Not Mature')]

total_credit = matured + not_mature

total_credit = 'Total \n' + joker(total_credit)

# ------------------new code--------------------
matured = joker(matured)
not_mature = joker(not_mature)

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
plt.tight_layout()
plt.savefig('category_wise_credit.png')

print('2. Category wise Credit Generated ')
plt.close()

# #------- Sector Wise Credit not_mature -------------------
sector_credit_df = pd.read_sql_query(""" Select case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end  as CustType,Sum(OUT_NET) as Amount from
                    (select INVNUMBER,INVDATE,
                    CUSTOMER,TERMS,MAINCUSTYPE,
                    CustomerInformation.CREDIT_LIMIT_DAYS,
                    datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS as Days_Diff,
                    OUT_NET from [ARCOUT].dbo.[CUST_OUT]
                    join ARCHIVESKF.dbo.CustomerInformation
                    on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST
                    where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash') as TblCredit
                    where Days_Diff<=0
                    group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end """, connection,
                                     params={branc_names})

Institution = int(sector_credit_df.Amount.iloc[0])
retail = int(sector_credit_df.Amount.iloc[1])

values = [Institution, retail]

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
plt.tight_layout()
plt.savefig('regular_credit.png')
plt.close()
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
                    where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash') as TblCredit
                    where Days_Diff>0
                    group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end
                    """, connection, params={branc_names})

Institution = int(mature_credit_df.Amount.iloc[0])
retail = int(mature_credit_df.Amount.iloc[1])

values = [Institution, retail]

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
plt.tight_layout()
plt.savefig('matured_credit.png')
plt.close()
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
                where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash' and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)<=0
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

                                    """, connection, params={branc_names})

width = 0.75
AgingDays = CloseTo_mature_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = CloseTo_mature_df['Amount']
tovalue = sum(performance)
maf_kor = max(performance)
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
                ha='center', va='bottom', fontsize=12, rotation=45, fontweight='bold')


autolabel(bars)


def autolabel2(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .5 * height,
                str(round(((height / tovalue) * 100), 1)) + "%",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel2(bars)

plt.xticks(y_pos, AgingDays, fontsize=12)
plt.yticks(np.arange(0, maf_kor + (.6 * maf_kor), maf_kor / 5), fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
# plt.yticks(np.arange(0, round(ran) + (.6 * round(ran))), fontsize='12')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('6. Non-Matured Credit Age', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('closed_to_matured_credit.png')
plt.close()
print('6. Closed to mature Credit Generated')

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
                where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash' ) as TblCredit
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
                                        """, connection, params={branc_names})

width = 0.75
AgingDays = aging_mature_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = aging_mature_df['Amount']
aging_total = sum(performance)
fig, ax = plt.subplots()
bars = plt.bar(y_pos, performance, width, align='center', alpha=1, color=colors)
maf_kor2 = max(performance)


def autolabel(bars):
    # attach some text labels
    for rect in bars:
        height = int(rect.get_height())
        ax.text(rect.get_x() + rect.get_width() / 2., .995 * height,
                for_bar(height), ha='center', va='bottom', fontsize=12, rotation=45, fontweight='bold')


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
# plt.yticks(fontsize=12)
plt.yticks(np.arange(0, maf_kor2 + (.6 * maf_kor2), maf_kor2 / 5), fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('4. Matured Credit Age', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('aging_matured_credit.png')
plt.close()
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
                    where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS='Cash') as TblCredit

                    group by case when MAINCUSTYPE='RETAIL' then 'Retail' else 'Institute' end """, connection,
                                        params={branc_names})
Institution = int(sector_wise_cash_df.Amount.iloc[0])
retail = int(sector_wise_cash_df.Amount.iloc[1])

values = [Institution, retail]
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
plt.tight_layout()
plt.savefig('Category_wise_cash.png')
plt.close()
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
            where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS='Cash') as TblCredit
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
        order by SERIAL """, connection, params={branc_names})

width = 0.75
AgingDays = cash_drop_df['AgingDays']
y_pos = np.arange(len(AgingDays))
performance = cash_drop_df['Amount']
new_totall = sum(performance)
fig, ax = plt.subplots()
bars = plt.bar(y_pos, performance, width, align='center', alpha=0.9, color='#fd00ff')
maf_kor3 = max(performance)
bars[0].set_color('#f00228')
bars[1].set_color('#ff6500')
bars[2].set_color('#deff00')
bars[3].set_color('#2c8e14')


def autolabel(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .995 * height,
                for_bar(height),
                ha='center', va='bottom', fontweight='bold', rotation=45, fontsize=12)


autolabel(bars)


def autolabel4(bars):
    for bar in bars:
        height = int(bar.get_height())
        ax.text(bar.get_x() + bar.get_width() / 2., .5 * height,
                str(round(((height / new_totall) * 100), 1)) + "%",
                ha='center', va='bottom', fontsize=12, fontweight='bold')


autolabel4(bars)

plt.xticks(y_pos, AgingDays, fontsize=12)
# plt.yticks(fontsize=12)
plt.yticks(np.arange(0, maf_kor3 + (.6 * maf_kor3), maf_kor3 / 5), fontsize=12)
plt.xlabel('Aging Days', color='black', fontsize=14, fontweight='bold')
plt.ylabel('Amount', color='black', fontsize=14, fontweight='bold')
plt.title('8. Aging of Cash Drop', color='#3e0a75', fontweight='bold', fontsize=16)
plt.tight_layout()
plt.savefig('aging_cash_drop.png')
plt.close()
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
                    and AUDTORG like ?
                                    """, connection, params={branc_names})
LastDayReturn = pd.read_sql_query("""
                    select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
                    where AUDTORG like ? and transtype<>1 and PRICELIST <> 0 and
                    (TRANSDATE = convert(varchar(8),DATEADD(D,0,GETDATE()-1),112))
                     """, connection, params={branc_names})
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
                    and AUDTORG like ?
                    """, connection, params={branc_names})

monthly_return_df = pd.read_sql_query("""select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
            where AUDTORG like ? and transtype<>1 and PRICELIST <> 0 and
            (TRANSDATE between
            (convert(varchar(8),DATEADD(mm, DATEDIFF(mm, 0, GETDATE()), 0),112))
            and (convert(varchar(8),DATEADD(D,0,GETDATE()),112)))
                        """, connection, params={branc_names})

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
                where AUDTORG like ? AND TRANSDATE>= (convert(varchar(8),DATEADD(yy, DATEDIFF(yy, 0, GETDATE()), 0),112))
                """, connection, params={branc_names})
yearly_return = pd.read_sql_query("""
                      select sum(EXTINVMISC) as ReturnAmount from OESalesDetails where
                    AUDTORG like ? and
                    transtype<>1 and PRICELIST <> 0 and
                    (TRANSDATE between (convert(varchar(8),DATEADD(yy, DATEDIFF(yy, 0, GETDATE()), 0),112))
                    and (convert(varchar(8),DATEADD(D,0,GETDATE()),112)))
     """, connection, params={branc_names})
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
yago_monthly_sales = pd.read_sql_query(""" Declare @YagoMonthStartDay NVARCHAR(MAX);
                    Declare @YagomonthCurrentDay NVARCHAR(MAX);
                    SET @YagoMonthStartDay = convert(varchar(6), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112)
                    set @YagomonthCurrentDay = convert(varchar(8), DATEADD(year, -1, GETDATE()), 112)
                    select  Sum(EXTINVMISC) as  MTDSales from OESalesDetails
                    where TRANSDATE between  @YagoMonthStartDay and @YagomonthCurrentDay
                    and AUDTORG like ?
                     """, connection, params={branc_names})

yago_monthly_return_df = pd.read_sql_query("""select ISNULL(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails
    where AUDTORG like ? and transtype<>1 and PRICELIST <> 0 and
    transdate between (convert(varchar(6), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
    and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112))
                        """, connection, params={branc_names})
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
                    where AUDTORG like ? AND TRANSDATE between (convert(varchar(8), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
                    and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112))
                    """, connection, params={branc_names})
yago_yearly_return = pd.read_sql_query("""
                    select isnull(sum(EXTINVMISC), 0) as ReturnAmount from OESalesDetails where
                    AUDTORG like ? and
                    transtype<>1 and PRICELIST <> 0 and
                    (TRANSDATE between (convert(varchar(8), DATEFROMPARTS ( DATEPART(yyyy, GETDATE()) - 1, 1, 1 ), 112))
                    and (convert(varchar(8), DATEADD(year, -1, GETDATE()), 112)))
                    """, connection, params={branc_names})
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
dirpath = os.path.dirname(os.path.realpath(__file__))
im = Image.open(dirpath + '/return.png')
left = 0
top = 0
right = 1600
bottom = 200
im1 = im.crop((left, top, right, bottom))
im1.save('return.png')
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
                    where AUDTORG like ? and transtype<>1 and PRICELIST <> 0 and
                    TRANSDATE between
                    (convert(varchar(8),DATEADD(mm, DATEDIFF(mm, 0, GETDATE()), 0),112))
                    and (convert(varchar(8),DATEADD(D,0,GETDATE()),112))
                    group by Cause_Of_Return_ID """, connection, params={branc_names})

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

color = ['#1ff015', '#418af2', '#f037d9', '#ecc13f', '#e2360a', '#1ff015', '#418af2', '#f037d9', '#ecc13f',
         '#e2360a']
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
                ha='center', va='bottom', fontsize=12, rotation=45, fontweight='bold')
        loop = loop + 1


autolabel(rects1)
plt.xticks(y_pos, Cause_name, fontsize='12', color='black')
plt.yticks(np.arange(0, 101, 10), color='black', fontsize='12')
plt.xlabel('Return Cause', fontsize='14', color='black', fontweight='bold')
plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
plt.title('10. Cause wise Return', color='#3e0a75', fontsize='16', fontweight='bold')
plt.tight_layout()
plt.savefig('Cause_with_return.png')
plt.close()
print('10. cause wise return generated')

# ----------------------Delivery man wise return---------------------

delivery_man_wise_return_df = pd.read_sql_query("""
    select TWO.ShortName as DPNAME, ISNULL(SUM(Sales.ReturnAmount), 0) as ReturnAmount from
    (select  DPID, AUDTORG,
    ISNULL(sum(case when TRANSTYPE<>1 then INVNETH *-1 end), 0) /sum(case when TRANSTYPE=1 then INVNETH end)*100 as ReturnAmount
    from OESalesSummery
    where AUDTORG like ? and 
    left(TRANSDATE,6)=convert(varchar(6),getdate(),112)
    group by DPID,AUDTORG) as Sales
    left join
    (select distinct AUDTORG,ShortName,DPID from DP_ShortName where AUDTORG like ?) as TWO
    on Sales.DPID = TWO.DPID
    and Sales.AUDTORG=TWO.AUDTORG
    where TWO.ShortName is not null
    and Sales.ReturnAmount>0
    group by TWO.ShortName
    order by ReturnAmount DESC  """, connection, params=(branc_names, branc_names))

if delivery_man_wise_return_df.empty == True:
    fig, ax = plt.subplots(figsize=(12.81, 4.8))
    plt.bar('No Record', 0.00, align='center', alpha=0.9)

    plt.xlabel('Delivery Person', fontsize='14', color='black', fontweight='bold')
    plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
    plt.title("9. Delivery Person's Return %", color='#3e0a75', fontsize='16', fontweight='bold')
    plt.tight_layout()
    plt.savefig('Delivery_man_wise_return.png')

else:
    DPNAME = delivery_man_wise_return_df['DPNAME']
    y_pos = np.arange(len(DPNAME))
    ReturnAmount = abs(delivery_man_wise_return_df['ReturnAmount'])
    ReturnAmount = ReturnAmount.values.tolist()

    ran = max(ReturnAmount)
    color = '#418af2'
    fig, ax = plt.subplots(figsize=(12.81, 4.8))
    rects1 = plt.bar(y_pos, ReturnAmount, align='center', alpha=0.9, color=color)


    def autolabel(bars):
        loop = 0
        for bar in bars:
            show = ReturnAmount[loop]
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, (1.05 * height),
                    '%.2f' % (show) + '%', ha='center', va='bottom', fontsize=12, rotation=70, fontweight='bold')
            loop = loop + 1


    autolabel(rects1)

    plt.xticks(y_pos, DPNAME, rotation='vertical', fontsize='12')
    plt.yticks(np.arange(0, round(ran) + (.6 * round(ran))), fontsize='12')
    plt.xlabel('Delivery Person', fontsize='14', color='black', fontweight='bold')
    plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
    plt.title("9. Delivery Person's Return %", color='#3e0a75', fontsize='16', fontweight='bold')
    plt.tight_layout()

    plt.savefig('Delivery_man_wise_return.png')

    plt.close()
    print('9. Delivery man wise return generated')

# ----------------------------------------------------------------------------

daily_sales_df = pd.read_sql_query("""select Right(transdate,2) as [day], Sum(EXTINVMISC)/1000 as  EverydaySales from OESalesDetails  
                where LEFT(TRANSDATE,6) =convert(varchar(6), GETDATE(),112)  and AUDTORG like ?
                group by transdate
                order by transdate """, connection, params={branc_names})

EveryD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                Declare @DaysInMonth NVARCHAR(MAX);
                SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                where YEARMONTH = @CurrentMonth and AUDTORG like ?""", connection, params={branc_names})

totarget = int(EveryD_Target_df.YesterdayTarget)
if totarget == 0:
    EveryD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                     Declare @DaysInMonth NVARCHAR(MAX);
                     SET @CurrentMonth = convert(varchar(6), GETDATE()-1,112)
                     SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                     select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                     where YEARMONTH = @CurrentMonth and AUDTORG like ?""", connection, params={branc_names})

    target = int(EveryD_Target_df.YesterdayTarget)
    print(target)

Every_day = daily_sales_df['day'].tolist()

y_pos = np.arange(len(Every_day))

every_day_sale = daily_sales_df['EverydaySales'].tolist()

n = 1
Target = []
labell = []
for z in y_pos:
    labell.append(n)
    Target.append(int(target / 1000))
    n = n + 1

fig, ax = plt.subplots(figsize=(12.81, 4.8))
labels = labell

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

rects2 = ax.bar(y_pos, every_day_sale, width, label='Sales')
maf_kor4 = max(Target)
# Add some text for labels, title and custom x-axis tick labels, etc.
line = ax.plot(Target, color='orange', label='Target')
plt.yticks(np.arange(0, maf_kor4 + (.8 * maf_kor4), maf_kor4 / 5), fontsize=12)
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
                ha='center', va='bottom', fontsize=12, rotation=45, fontweight='bold')


autolabel(rects2)

fig.tight_layout()

plt.savefig("Day_Wise_Target_vs_Sales.png")
print('12. Day Wise Target vs Sales')
plt.close()
# --------------------------------------------------------------------------

LD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                Declare @DaysInMonth NVARCHAR(MAX);
                SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                where YEARMONTH = @CurrentMonth and AUDTORG like ? """, connection, params={branc_names})
toto = LD_Target_df.values
ld_target = int(toto[0, 0])

MTD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                select ISNULL((Sum(TARGET)), 0) as  MTDTarget from TDCL_BranchTarget  
                where YEARMONTH = @CurrentMonth and AUDTORG like ? """, connection, params={branc_names})

momo = MTD_Target_df.values
Mtd_target = int(momo[0, 0])

from datetime import date

given_date = datetime.today().date()

toto = datetime.today().date()
first = given_date.replace(day=1)
day1 = date(toto.year, toto.month, toto.day)
day2 = date(first.year, first.month, first.day)
no_of_days = (day1 - day2).days

import calendar
import datetime

now = datetime.datetime.now()
total_days = calendar.monthrange(now.year, now.month)[1]

final_mtd_target = int((Mtd_target / total_days) * no_of_days)

YTD_Target_df = pd.read_sql_query(""" select ISNULL((Sum(TARGET)), 0) as  YTDTarget from TDCL_BranchTarget  
                where convert(varchar(4), YEARMONTH,112) = convert(varchar(4), GETDATE(),112)
                  and AUDTORG like ? """,
                                  connection, params={branc_names})

yoyo = YTD_Target_df.values
Ytd_target = int(yoyo[0, 0])
final_ytd_target = int(Ytd_target - ((Mtd_target / total_days) * (total_days - no_of_days)))

LD_Sales_df = pd.read_sql_query(""" Declare @Currentday NVARCHAR(MAX);
                SET @Currentday = convert(varchar(8), GETDATE()-1,112);
                select  Sum(EXTINVMISC) as  YesterdaySales from OESalesDetails  
                where LEFT(TRANSDATE,8) = @Currentday and AUDTORG like ? """, connection, params={branc_names})

Ld_toto = LD_Sales_df.values
MTD_Sales_df = pd.read_sql_query(""" Declare @Currentmonth NVARCHAR(MAX);
                SET @Currentmonth = convert(varchar(6), GETDATE(),112);
                select  Sum(EXTINVMISC) as  MTDSales from OESalesDetails  
                where LEFT(TRANSDATE,6) = @Currentmonth and AUDTORG like ? """, connection, params={branc_names})

MTD_momo = MTD_Sales_df.values
YTD_Sales_df = pd.read_sql_query(""" Declare @Currentyear NVARCHAR(MAX);
                SET @Currentyear = convert(varchar(4), GETDATE(),112);
                select  Sum(EXTINVMISC) as  YTDSales from OESalesDetails  
                where LEFT(TRANSDATE,4) = @Currentyear and AUDTORG like ? """, connection, params={branc_names})

YTD_yoyo = YTD_Sales_df.values

labels = ['LD', 'MTD', 'YTD']
aa = int(final_mtd_target / 1000)
bb = int(final_ytd_target / 1000)
cc = int(ld_target / 1000)
Targets = [cc, aa, bb]

Sales = [int(Ld_toto[0, 0] / 1000), int(MTD_momo[0, 0] / 1000), int(YTD_yoyo[0, 0] / 1000)]

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

fig, ax = plt.subplots()
rects1 = ax.bar(x - width / 2, Targets, width, label='Target')
rects2 = ax.bar(x + width / 2, Sales, width, label='Sales')

maf_kor5 = max(Targets)
# Add some text for labels, title and custom x-axis tick labels, etc.
plt.yticks(np.arange(0, maf_kor5 + (.8 * maf_kor5), maf_kor5 / 5), fontsize=12)
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
                ha='center', va='bottom', fontsize=12, rotation=45, fontweight='bold')


autolabel(rects1)
autolabel(rects2)

fig.tight_layout()

# plt.show()
plt.savefig('LD_MTD_YTD_TARGET_vs_sales.png')
print('11. Target vs Sales ')

plt.close()

# -------------------------------------cumulative sales vs target-----------------


daily_sales2_df = pd.read_sql_query("""select Right(transdate,2) as [day], Sum(EXTINVMISC)/1000 as  EverydaySales from OESalesDetails  
                where LEFT(TRANSDATE,6) =convert(varchar(6), GETDATE(),112)  and AUDTORG like ?
                group by transdate
                order by transdate""", connection, params={branc_names})

EveryD_Target2_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                Declare @DaysInMonth NVARCHAR(MAX);
                SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                select ISNULL(((Sum(TARGET)/@DaysInMonth)/1000), 0) as  YesterdayTarget from TDCL_BranchTarget  
                where YEARMONTH = @CurrentMonth and AUDTORG like ? """, connection, params={branc_names})
if totarget == 0:
    EveryD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                    Declare @DaysInMonth NVARCHAR(MAX);
                    SET @CurrentMonth = convert(varchar(6), GETDATE()-1,112)
                    SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                    select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                    where YEARMONTH = @CurrentMonth and AUDTORG like ?""", connection, params={branc_names})

    target = int(EveryD_Target_df.YesterdayTarget)

target = int(EveryD_Target2_df.YesterdayTarget)
Every_day = daily_sales2_df['day'].tolist()
y_pos = np.arange(len(Every_day))

every_day_sale = daily_sales2_df['EverydaySales'].tolist()
n = 1
Target = []
labell = []
for z in y_pos:
    labell.append(n)
    Target.append(int(target / 1000))
    n = n + 1
labell.append(20)

# ----------------code for cumulitive sales------------
import calendar
import datetime

now = datetime.datetime.now()
total_days = calendar.monthrange(now.year, now.month)[1]
# print(total_days)

new_target = target
# print(new_target)

z = len(labell)
# print(len(labell))
fin_target = 0
ttt = []
for t_value in range(0, total_days + 1):
    # print(t_value)
    fin_target = new_target * t_value
    # print(fin_target)
    ttt.append(fin_target)
    fin_target = 0
# print(ttt) #-------------------target data

values = every_day_sale
length = len(values)

new = [0]
final = 0
for val in values:
    # print(val)
    get_in = values.index(val)
    # print(get_in)
    if get_in == 0:
        new.append(val)
    else:
        for i in range(0, get_in + 1):
            final = final + values[i]
        new.append(final)
        final = 0                                

# print(every_day_sale)
# print(new)#--------------------------sales data


x = range(len(ttt))
xx = range(len(new))

list_index_for_target = len(ttt) - 1
# print(list_index_for_target)

list_index_for_sale = len(new) - 1
# print(list_index_for_sale)
# Change the color and its transparency
fig, ax = plt.subplots(figsize=(12.81, 4.8))
plt.fill_between(x, ttt, color="skyblue", alpha=1)
plt.plot(xx, new, color="red", linewidth=5, linestyle="-")
plt.xlabel('Days', fontsize='14', color='black', fontweight='bold')
plt.ylabel('Amount', fontsize='14', color='black', fontweight='bold')
# ax.set_ylabel('Amount', fontsize='14', color='black', fontweight='bold')
# ax.set_xlabel('Day', fontsize='14', color='black', fontweight='bold')
plt.title('13. MTD Target vs Sales - Cumulative', fontsize=16, fontweight='bold', color='#3e0a75')
plt.xticks(np.arange(1, total_days + 1, 1))
plt.legend(['sales', 'target'], loc='upper left', fontsize='14')
plt.text(list_index_for_sale, ttt[list_index_for_sale], get_value(str(int(ttt[list_index_for_sale]))) + 'K',
         color='black', fontsize=12, fontweight='bold')
plt.text(list_index_for_sale, new[list_index_for_sale], get_value(str(int(new[list_index_for_sale]))) + 'K',
         color='black', fontsize=12, fontweight='bold')
plt.text(list_index_for_target, ttt[list_index_for_target], get_value(str(int(ttt[list_index_for_target]))) + 'K',
         color='black', fontsize=12, fontweight='bold')

plt.grid()
plt.savefig("Cumulative_Day_Wise_Target_vs_Sales.png")
plt.close()
print('13. Cumulative day wise target sales')
# plt.show()


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

imp24 = Image.open(dirpath + "/Cumulative_Day_Wise_Target_vs_Sales.png")
widthx, heightx = imp24.size
imageSize = Image.new('RGB', (1283, 482))
imageSize.paste(imp24, (1, 1))
imageSize.save(dirpath + "/Cumulative_Day_wise_target_sales.png")

# --------------- Close TO Mature Credit Redords -------------------

ClosedToMaturedcredit_df = pd.read_sql_query("""
                   select CUSTOMER as 'Cust ID', CUSTNAME as 'Cust Name',
    CustomerInformation.TEXTSTRE1 as 'Address', CustomerInformation.MSOTR as 'Territory',INVNUMBER as 'Inv Number',
    convert(varchar,convert(datetime,(convert(varchar(8),INVDATE,112))),106)  as 'Inv Date', CustomerInformation.CREDIT_LIMIT_DAYS as 'Days Limit',
    (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)*-1 as 'Matured In Days', OUT_NET as 'Credit Amount'
    from [ARCOUT].dbo.[CUST_OUT]
    join ARCHIVESKF.dbo.CustomerInformation
    on [CUST_OUT].CUSTOMER = CustomerInformation.IDCUST

    where  [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash'
        and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)  between -3 and 0
        order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)desc
        , OUT_NET desc

     """, connection, params={branc_names})
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

    where  [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash'
        and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)  between -3 and 0
        order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS)desc
        , OUT_NET desc
                   """, connection, params={branc_names})
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

                    where  [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash' and OUT_NET>1
                    and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) >= 1
                    order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) desc
                                    , OUT_NET desc
                     """, connection, params={branc_names})

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

                    where  [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS<>'Cash' and OUT_NET>1
                    and (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) >= 1
                    order by (datediff([dd] , CONVERT (DATETIME , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1-CREDIT_LIMIT_DAYS) desc
                     , OUT_NET desc
                    """, connection, params={branc_names})

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

                where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS='Cash'  and OUT_NET>1
                and (datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1) >=4
                order by datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 desc
                  , OUT_NET desc
                 """, connection, params={branc_names})
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

            where [ARCOUT].dbo.[CUST_OUT].AUDTORG like ? and TERMS='Cash'  and OUT_NET>1
            and (datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1) >=4
            order by datediff([dd] , CONVERT (DATE , LTRIM(cust_out.INVDATE) , 102) , GETDATE())+1 desc
             , OUT_NET desc
            """, connection, params={branc_names})

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

# ------------ Group email ----------------------------------------
msgRoot = MIMEMultipart('related')
me = 'erp-bi.service@transcombd.com'
to = ['rejaul.islam@transcombd.com', '']
cc = ['', '']
bcc = ['', '']

# to = ['tdclndm@tdcl.transcombd.com',''] cc = ['hislam@skf.transcombd.com','muhammad.mainuddin@tdcl.transcombd.com',
# 'nurul.amin@tdcl.transcombd.com'] bcc = ['biswascma@yahoo.com', 'bayezid@transcombd.com',
# 'zubair.transcom@gmail.com', 'yakub@transcombd.com', 'tawhid@transcombd.com', 'rejaul.islam@transcombd.com',
# 'fazle.rabby@transcombd.com','aftab.uddin@transcombd.com','roseline@transcombd.com']

recipient = to + cc + bcc

subject = "SK+F Formulation Reports - " + branch_name + " Branch"

email_server_host = 'mail.transcombd.com'
port = 25

msgRoot['From'] = me
# msgRoot['to'] = recipient
msgRoot['To'] = ', '.join(to)
msgRoot['Cc'] = ', '.join(cc)
msgRoot['Bcc'] = ', '.join(bcc)
msgRoot['Subject'] = subject

msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

msgText = MIMEText("""
                        <img src="cid:banner" height='200' width='796'> <br>
                       <img src="cid:all_credit"  width='796'><br>
                       <img src="cid:all_Matured_credit"  width='796'> <br>
                       <img src="cid:all_regular_credit"  width='796'> <br>
                       <img src="cid:all_cash"  width='796'> <br>
                       <img src="cid:main_return" width='796'> <br>

                       <img src="cid:delivery_mans_return_with_cause"  width='796'><br>
                        <img src="cid:return_with_cause"  width='796'><br>
                       <img src="cid:day_wise_sales_target"  width='796'><br>
                       <img src="cid:cumulative_day_wise_sales_target"  width='796'><br>

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

# ---------------------------------------------------------------------------
fp = open(dirpath + '/Cumulative_Day_wise_target_sales.png', 'rb')
return_with_cause = MIMEImage(fp.read())
fp.close()

return_with_cause.add_header('Content-ID', '<cumulative_day_wise_sales_target>')
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

# ----------- Finally send mail and close server connection ---
server = smtplib.SMTP(email_server_host, port)
server.ehlo()
print('\n-----------------')
print('Sending Mail')
server.sendmail(me, recipient, msgRoot.as_string())
print('Mail Send')
print('-------------------')
server.close()
