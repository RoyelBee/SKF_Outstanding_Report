import pandas as pd
import pyodbc as db
import matplotlib.pyplot as plt

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=137.116.139.217;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp@123')

branc_names = 'BSLSKF'
# delivery_man_wise_return_df = pd.read_sql_query("""
#     select TWO.ShortName as DPNAME, ISNULL(SUM(Sales.ReturnAmount), 0) as ReturnAmount from
#     (select  DPID, AUDTORG,
#     ISNULL(sum(case when TRANSTYPE<>1 then INVNETH *-1 end), 0) /sum(case when TRANSTYPE=1 then INVNETH end)*100 as ReturnAmount
#     from OESalesSummery
#     where AUDTORG like ? and
#     left(TRANSDATE,6)=convert(varchar(6),getdate(),112)
#     group by DPID,AUDTORG) as Sales
#     left join
#     (select distinct AUDTORG,ShortName,DPID from DP_ShortName where AUDTORG like ?) as TWO
#     on Sales.DPID = TWO.DPID
#     and Sales.AUDTORG=TWO.AUDTORG
#     where TWO.ShortName is not null
#     and Sales.ReturnAmount>0
#     group by TWO.ShortName
#     order by ReturnAmount DESC  """, connection, params=(branc_names, branc_names))
#
# print(delivery_man_wise_return_df)
#
# if delivery_man_wise_return_df.empty==True:
#     fig, ax = plt.subplots(figsize=(12.81, 4.8))
#     rects1 = plt.bar('No Record', 0.00, align='center', alpha=0.9)
#
#     plt.xlabel('Delivery Person', fontsize='14', color='black', fontweight='bold')
#     plt.ylabel('Return Percentage', fontsize='14', color='black', fontweight='bold')
#     plt.title("9. Delivery Person's Return %", color='#3e0a75', fontsize='16', fontweight='bold')
#     plt.tight_layout()
#
#     plt.savefig('Delivery_man_wise_return.png')
#
# else:
#     print('Hello ')


EveryD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                Declare @DaysInMonth NVARCHAR(MAX);
                SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                where YEARMONTH = @CurrentMonth and AUDTORG like ?""", connection, params={branc_names})


totarget = EveryD_Target_df.values
if totarget == 0:
    EveryD_Target_df = pd.read_sql_query(""" Declare @CurrentMonth NVARCHAR(MAX);
                    Declare @DaysInMonth NVARCHAR(MAX);
                    SET @CurrentMonth = convert(varchar(6), GETDATE()-1,112)
                    SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                    select ISNULL((Sum(TARGET)/@DaysInMonth), 0) as  YesterdayTarget from TDCL_BranchTarget  
                    where YEARMONTH = @CurrentMonth and AUDTORG like ?""", connection, params={branc_names})

    target = int(EveryD_Target_df.YesterdayTarget)
    print(target)
else:
    print('Current Target ')
