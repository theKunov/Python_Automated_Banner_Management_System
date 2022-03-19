from operator import indexOf
from sqlite3 import connect
import pyodbc, sys, time 

from setup_functions import start, close_chrome
from setup_main import run_prog

# Add server information here
connect_str = 'DRIVER={ODBC Driver 17 for SQL Server}; \
                SERVER=; \
                DATABASE= ; \
                UID= ; \
                PWD=;'

get_data = "SELECT DISTINCT TOP 999 \
            CONCAT(RIGHT(YEAR(B.DateFrom), 2), 'CW', DATEPART(ISO_WEEK, B.DateFrom), '-', REPLACE(B.SK_Booking,' ','-'), '-', REPLACE(V.Name,' ','-'), '-', REPLACE(S.SlotName,' ','-'), '-', P.PlacementCode) Name \
            ,S.SK_BannerSlot \
            ,DATEPART(ISO_WEEK, B.DateFrom) Week \
            ,DATEADD(hour, -14, B.DateFrom) DateStart \
            ,DATEADD(hour, 18, B.DateTo) DateEnd \
            FROM WEBMANAGER.dbo.Tbl_Booking B WITH (NOLOCK) \
            INNER JOIN WEBMANAGER.dbo.Tbl_BannerSlot S WITH (NOLOCK) \
            ON B.SK_BannerSlot = S.SK_BannerSlot \
            INNER JOIN WEBMANAGER.dbo.Tbl_BannerGroup G WITH (NOLOCK) \
            ON S.SK_BannerGroup = G.SK_BannerGroup \
            INNER JOIN WEBMANAGER.dbo.Tbl_PlacementBannerGroup PG WITH (NOLOCK) \
            ON S.SK_BannerGroup = PG.SK_BannerGroup \
            INNER JOIN WEBMANAGER.dbo.Tbl_Placement P WITH (NOLOCK) \
            ON PG.SK_Placement = P.SK_Placement \
            INNER JOIN WEBMANAGER.dbo.Tbl_Banner I WITH (NOLOCK) \
            ON B.SK_Banner = I.SK_Banner \
            INNER JOIN MARCOM.dbo.ad_hersteller V WITH (NOLOCK) \
            ON B.SK_Manufacturer = V.id \
            INNER JOIN WEBMANAGER.dbo.Tbl_BannerDetail D WITH (NOLOCK) \
            ON D.SK_Banner = B.SK_Banner \
            WHERE B.SK_Status = 5 \
            AND G.SK_BannerGroup NOT IN (8, 17, 27, 35, 36, 37, 38, 39, 49, 50, 51, 52, 55, 57) \
            AND BannerName NOT LIKE '%Intern%' \
            AND D.Value IS NOT NULL \
            AND DATEPART(ISO_WEEK, B.DateFrom) = DATEPART(ISO_WEEK, DATEADD(WEEK, 1, GETDATE()))"

push_data = "UPDATE WEBMANAGER.dbo.Tbl_Booking \
                        SET SK_Status = 6 \
                        WHERE SK_Booking = (?)"

with pyodbc.connect(connect_str) as cnxn:
    cursor = cnxn.cursor()
    cursor.execute(get_data)
    data = cursor.fetchall()


start()

success_banners = []
 
for item in data:

    x = data.index(item)

    # Banner image name
    imgName = str(item[0])

    # Slot of banner
    slot = int(item[1])

   # start date
    s_date = str(item[3]).split("-")
    s_date2 = str(s_date[2]).split(" ", 1) 
    strDate = s_date2[0] + "." + s_date[1] + "." + s_date[0] + " " + s_date2[1]

    # end date
    e_date = str(item[4]).split("-")
    e_date2 = str(e_date[2]).split(" ", 1) 
    endDate = e_date2[0] + "." + e_date[1] + "." + e_date[0] + " " + e_date2[1]

    # Call main banner setup function
    banner_set = run_prog(imgName, strDate, endDate, slot, x)

    if banner_set != "":
        checker =  banner_set in success_banners
        
        if checker == False :
            success_banners.append(banner_set)

    print(success_banners)

    # print(x, imgName, strDate, endDate, slot)


# Stop connection with Chrome browser
time.sleep(2)
close_chrome()
time.sleep(3)


# Update Banners Status by pushing new data to DB

# cursor.execute(push_data)
with pyodbc.connect(connect_str) as conxn:
    cursor2 = conxn.cursor()   

    for banner in success_banners:
        cursor2.execute("UPDATE WEBMANAGER.dbo.Tbl_Booking \
                            SET SK_Status = 6 \
                            WHERE SK_Booking = (?)", banner)    
        conxn.commit()


time.sleep(3)
print("Data Pushed")

# Stop connection with DB server
cnxn.close()
print("Connection to DB stopped")
time.sleep(3)

# Stop Python code
sys.exit()
