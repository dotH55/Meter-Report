from datetime import date, timedelta
import os, csv, sys
import pyodbc

#Date = (date.today() - timedelta(days = 2)).strftime("%Y-%m-%d")
Date = date.today().strftime("%Y-%m-%d")

def main():

    CSV_File = "US_BILLING_" + Date.replace('-', '') + "_01.csv"
    #at this point the program has retrieved the file and extracted it
    # in the current directory 
    #at this point the program has retrieved the file and extracted it
    # in the current directory 
    
    server = ''
    database = ''
    username = ''
    password = ''
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()



    file = open(CSV_File)
    reader = csv.reader(file)
    rows = []
    for row in reader:
        rows.append(row)

    print("\n")
    rowcount = 0
    for r in rows: #each r is a row with an array of elements each being a column of the csv
        if rowcount > 2: #rows below the csv file info. I.E. Machines
            print("\n " + str(rowcount))
            for i in range(0,3): #prints the first three cols, Date Time and Machine Model
                print(r[i])

            print("Serial: " + r[3]) #col of 
            
            var109 = "0"
            var112 = "0"
            var120 = "0"
            var124 = "0"

            print("\nMeter Info")
            count = 0
            for i in r:
                if str(i) == str(109) and count % 2 == 0: # Total Black 2
                    print("Meter Count ID Number: " + i + "   Value: " + r[count + 1])
                    var109 = r[count + 1]
                if  str(i) == str(112) and count % 2 == 0: #Total 1
                    print("Meter Count ID Number: " + i + "   Value: " + r[count + 1])
                    var112 = r[count+ 1]
                if str(i) == str(120) and count % 2 == 0: #Total Full Color/Large
                    print("Meter Count ID Number: " + i + "   Value: " + r[count + 1])
                    var120 = r[count + 1]
                if str(i) == str(124) and count % 2 == 0: #Total Full Color + Single Color 2)
                    print("Meter Count ID Number: " + i + "   Value: " + r[count + 1])
                    var124 = r[count + 1]
                if count == len(r) - 1:
                    sqlDateStr = str(r[1])
                    SQLSTR = "INSERT INTO DSI.dbo.Meter (SerialNumber, MeterDate, BWMeter, ColorMeter, Source, InvoiceNumber, ObtainedBy, rowguid, BW_Large, Color_Large) VALUES ('" + r[3] + "', '" + sqlDateStr + "'," + var109 + " ,  " + var124 + " , 'Canon', ' " " ' , 'Canon', NEWID(), " + var112 + "  ,  " + var120 + " )" 
                    cursor.execute(SQLSTR) 
                    cnxn.commit()
                    SQLSTR = "UPDATE DSI.dbo.Serialized SET  Last_Remote_Checkin = '" + sqlDateStr + "', Last_Remote_Source = 'Canon' WHERE SerialNumber = '" + r[3] +"'"
                    cursor.execute(SQLSTR) 
                    cnxn.commit()
                count += 1
        elif rowcount <= 2:
            print(r) # the first 3 rows of the csv. Not machine info. COUNTRY CODE CORPORATE CODE CREATEDATE 
        rowcount += 1
    
    file.close()
    os.remove(CSV_File)

if __name__ == "__main__":
    try:
        main()
    except Exception:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)







