# Mobile County Program
# Equipment Form
# Pulls information out of each equipment form file and puts the data into the SQL database
# then moves the file from the main folder into the "Updated Files" folder

# Intialize with imports
import os
import win32com.client as win32
import sqlite3
import pandas as pd
import datetime


def create_connection(db_file):
    """ Create a database connection to the SQLite database
        specified by db_file
    :param db_file: database file
    :return:    Connection object or None
    """

    conn = sqlite3.connect(db_file)
    return conn


def last_value(cur, table, field):
    """ Finds the last record of the specified field from the specified table
    :param cur:  cursor object
    :param table:   database table from a sqlite database
    :param field:   field from within the specified table
    :return:    Value in the last field from the specified table
    """
    # lid: int
    # query = ''' SELECT field
    #             FROM table
    #             LAST_VALUE(field) OVER(
    #             ORDER BY field);
    #             VALUES(?,?)'''
    # entry = (table, field)
    query = """ SELECT * FROM Parts"""
    # value = cur.execute(query, entry)
    cur.execute(query)
    print(cur.lastrowid)
    return cur.lastrowid


def file_exist_delete(check_file):
    """ Determines if the file exist or not.  If it exists, it will be deleted.
    :param check_file:    file to be checked - includes path and filename
    :return:    None
    """
    if os.path.exists(check_file):
        os.remove(check_file)
    return


def convert_excel(orig_file, update_file):
    """

    :param orig_file:   .xlsm file to be converted to .xlsx file
    :param update_file: .xlsx file converted from .xlsm file
    :return:    None
    """

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(orig_file)
    excel.DisplayAlerts = False
    wb.DoNotPromptForConvert = True
    wb.CheckCompatibility = False
    # xlFileFormat enumeration - Source https://docs.microsoft.com/en-us/office/vba/api/Excel.XlFileFormat
    xlOpenXMLWorkbook = 51
    wb.SaveAs(update_file, FileFormat=xlOpenXMLWorkbook, ConflictResolution=2)
    wb.Close()
    # wb = None
    excel.Application.Quit()
    return


# def check_tables(cur, table_name):
#     """
#
#     :param: cur:        cursor object
#     :param table_name:   name of table to check for existence
#     :return:    Boolean
#     """
#     check_query = '''SELECT name
#                     FROM sqlite_master
#                     WHERE type="table" AND name=table_name'''
#
#     if not cur.execute(check_query):
#         return False
#
#     return True
#


def main():
    # Setup paths and file names for opening and saving excel files.
    main_path = "\\\\storage\\Departments\\Programs\\Equipment\\PartsForms"
    update_path = (
        "\\\\storage\\Departments\\Programs\\Equipment\\PartsForms\\UpdatedFiles"
    )
    prefilename = "f1234-2214-03012020"
    postfilename1 = ".xlsm"
    postfilename2 = ".xlsx"
    ofilename = [prefilename, postfilename1]
    openfilename = "".join(ofilename)
    cfilename = [prefilename, postfilename2]
    closefilename = "".join(cfilename)

    # check to see if a file with the same name already exist
    # if so delete the file
    original_file = os.path.join(main_path, openfilename)
    check_file = os.path.join(update_path, closefilename)
    file_exist_delete(check_file)

    # Open the .xlsm file and save it as a .xlsx file
    convert_excel(original_file, check_file)

    # Open the excel file and create pandas dataframes for each sheet
    excel_file = os.path.join(update_path, closefilename)
    print("Reading in the file and creating dataframes...")
    fleet_work_orders = pd.read_excel(
        excel_file, header=None, sheet_name="FWOTemp", skiprows=1
    )
    item_info = pd.read_excel(
        excel_file, header=None, sheet_name="ItemsTemp", skiprows=1
    )
    po_info = pd.read_excel(excel_file, header=None, sheet_name="POTemp", skiprows=1)

    # label the columns for each dataframe
    fleet_work_orders.columns = [
        "FleetWO",
        "WODate",
        "EquipNo",
        "Mechanic",
        "MechDate",
        "Clerk",
        "ClerkDate",
        "Buyer",
        "BuyerDate",
        "PONo",
    ]
    item_info.columns = [
        "FleetWO",
        "PONo",
        "Quantity",
        "PartNo",
        "Description",
        "PartFilled",
        "PartOrdered",
        "Price",
        "COPartNo",
    ]
    po_info.columns = ["FleetWO", "PONo", "PODate", "ReqNo", "Vendor"]

    # Set working directory to the database directory
    os.chdir(os.path.join(main_path, "Database"))

    # open a connection to the database
    conn = create_connection("Equpment.db")
    with conn:
        cursor = conn.cursor()
        # create the tables if they do not exist
        # create_tables(cursor, 'FleetWO', 'Parts', 'PurchaseOrders')

        cursor.execute(
            """CREATE TABLE IF NOT EXISTS FleetWO(
            FleetWO INT,
            WODate TEXT,
            EquipNo INT,
            Mechanic TEXT,
            MechDate TEXT,
            Clerk TEXT,
            ClerkDate TEXT,
            Buyer TEXT,
            BuyerDate TEXT,
            PONo INT
            );"""
        )
        cursor.execute(
            """CREATE TABLE IF NOT EXISTS Parts(
            PartID DOUBLE,
            FleetWO INT,
            PONo INT,
            Qty INT,
            PartNo TEXT,
            PartDesc TEXT,
            PartFilled INT,
            PartOrdered INT,
            Price REAL,
            COPartNo TEXT
            );"""
        )
        cursor.execute(
            """CREATE TABLE IF NOT EXISTS PurchaseOrders(
            FleetWO INT,
            PONo INT,
            PODate TEXT,
            ReqNo TEXT,
            Vendor TEXT
            );"""
        )

        conn.commit()
        print("Entering data into the 'FleetWO' table of the Equipment Database...")
        row: int
        for row in range(0, len(fleet_work_orders["FleetWO"])):
            fleet_wo = fleet_work_orders["FleetWO"][row]
            wo_date = str(fleet_work_orders["WODate"][row])[:10]
            equip_num = fleet_work_orders["EquipNo"][row]
            mechanic = fleet_work_orders["Mechanic"][row]
            mechanic_date = str(fleet_work_orders["MechDate"][row])[:10]
            clerk = fleet_work_orders["Clerk"][row]
            clerk_date = str(fleet_work_orders["ClerkDate"][row])[:10]
            buyer = fleet_work_orders["Buyer"][row]
            buyer_date = str(fleet_work_orders["BuyerDate"][row])[:10]
            po_num = fleet_work_orders["PONo"][row]
            sql = """ INSERT INTO FleetWO(FleetWO, WODate, EquipNo, Mechanic, MechDate, Clerk, ClerkDate,
                       Buyer, BuyerDate, PONo) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
            fwo_values = (
                fleet_wo,
                wo_date,
                equip_num,
                mechanic,
                mechanic_date,
                clerk,
                clerk_date,
                buyer,
                buyer_date,
                po_num,
            )
            print(type(fwo_values))
            print(fwo_values)
            cursor.execute(sql, fwo_values)
            conn.commit()
            cursor.execute("""SELECT * FROM FleetWO""")
            for row in range(cursor.rowcount):
                print(row)

        print("Entering data into the 'Parts' table of the Equipment Database...")
        last_id: int
        last_id = 0
        for row in range(0, len(item_info["PartNo"])):
            fleet_wo = item_info["FleetWO"][row]
            po_num = item_info["PONo"][row]
            quantity = item_info["Quantity"][row]
            part_num = item_info["PartNo"][row]
            part_description = item_info["Description"][row]
            part_filled = item_info["PartFilled"][row]
            part_ordered = item_info["PartOrdered"][row]
            unit_price = item_info["Price"][row]
            county_part_num = item_info["COPartNo"][row]
            last_id = last_value(cursor, "Parts", "PartID")
            part_id = last_id + 1
            print(part_id)
            sql1 = """ INSERT INTO Parts(PartID, FleetWO, PONo, Qty, PartNo, PartDesc, PartFilled, PartOrdered,
                       Price, COPartNo) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?,?)"""
            item_values = (
                part_id,
                fleet_wo,
                po_num,
                quantity,
                part_num,
                part_description,
                part_filled,
                part_ordered,
                unit_price,
                county_part_num,
            )
            print(item_values)
            cursor.execute(sql1, item_values)
            conn.commit()

        print(
            "Entering data into the 'PurchaseOrders' table of the Equipment Database..."
        )
        for row in range(0, len(po_info["PONo"])):
            fleet_wo = po_info["FleetWO"][row]
            po_num = po_info["PONo"][row]
            po_date = po_info["PODate"][row]
            requisition_num = po_info["ReqNo"][row]
            vendor = po_info["Vendor"][row]
            sql2 = """ INSERT INTO PurchaseOrders(FleetWO, PONo, PODate, ReqNo, Vendor)
                    VALUES(?, ?, ?, ?, ?)"""
            po_values = (fleet_wo, po_num, po_date, requisition_num, vendor)
            print(po_values)
            cursor.execute(sql2, po_values)
            conn.commit()

    print(
        "All values have been entered into the Equipment Database and the files have been moved to the UpdatedFiles directory"
    )


if __name__ == "__main__":
    main()
