import pyodbc
import ctypes, sys


server = 'APEX-DB01.APEX.LOCAL'
database = 'APEX'
username = r'APEX\e2user'
password = 'Apex2019'



conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                      'SERVER='+server,trusted_connection='yes')
    # ;UID='+'APEX\e2user'+';PWD='+ password

    # conn= pyodbc.connect(r'DRIVER={SQL server};SERVER=APEX-DB01.APEX.LOCAL;DATABASE=APEX;UID=APEX\e2uer;PWD=Apex2019',)

    # conn = pyodbc.connect(DirStrings.Apex_E2_DB_Server)


    # cur = conn.cursor()


    # ur.execute("SELECT "
    #                         "OrderDet.JobNo, "
    #                         "OrderDet.Priority, "
    #                         "Orders.CustCode, "
    #                         "Orders.PONum, "
    #                         "OrderDet.PartNo, "
    #                         "OrderDet.PartDesc, "
    #                         "OrderDet.QtyOrdered, "
    #                         "OrderDet.DueDate, "
    #                         "OrderRouting.WorkCntr, "
    #                         "OrderDet.User_Date2 "
    #                     "FROM "
    #                         "OrderDet "
    #                     "INNER JOIN "
    #                         "Orders "
    #                     "ON "
    #                         "OrderDet.OrderNo = Orders.OrderNo "
    #                     "INNER JOIN "
    #                         "OrderRouting "
    #                     "ON "
    #                         "OrderDet.JobNo = OrderRouting.JobNo "
    #                     "INNER JOIN "
    #                         "Releases "
    #                     "ON "
    #                         "OrderDet.JobNo = Releases.JobNo "
    #                     "WHERE "
    #                         "OrderDet.DueDate BETWEEN ? AND ? "
    #                     "AND "
    #                         "Releases.DelType = '0' "
    #                     "AND "
    #                         "OrderDet.Status = 'Open' "
    #                     "AND "
    #                         "(OrderRouting.Status = 'Current' OR (OrderRouting.Status = 'Finished' AND (OrderRouting.WorkCntr = 'Stageship' OR OrderRouting.WorkCntr = 'Shp To Stk'))) "
    #                     "AND NOT "
    #                         "(OrderDet.PartNo = 'Expedite FEE' "
    #                         "OR "
    #                         "OrderDet.PartNo = 'Fair Fee' "
    #                         "OR "
    #                         "OrderDet.PartNo = 'RPM EXPEDITE' "
    #                         "OR "
    #                         "OrderDet.PartNo = 'NRE' "
    #                         "OR "
    #                         "OrderDet.JobNo LIKE ? "
    #                         "OR "
    #                         "OrderRouting.JobNo = 'Null' "
    #                         "OR "
    #                         "OrderRouting.JobNo = '') "
    #                     "ORDER BY "
    #                         "OrderDet.DueDate ASC, "
    #                         "OrderDet.JobNo ASC",
    #                         str(StartDate),
    #                         str(NewEndDate),
    #                         "%C%")

            # self.row = cur.fetchall()
    # Getting part number information
    # cur.execute(
    #     'let'
    #         'Source = Sql.Databases("apex-DB01"),'
    #         'TEST_81620 = Source{[Name="TEST_81620"]}[Data],'
    #         'dbo_Part_Number = TEST_81620{[Schema="dbo",Item="Part_Number"]}[Data]'
    #     'in'
    #      'dbo_Part_Number')

    # row = cursor.fetchone()
    # for x in range(5):
    #     print(row)
    #     row = cursor.fetchone()
    # # All the PN info you need should be in “Part_Number” and “Revision_Level”

# ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1) 3