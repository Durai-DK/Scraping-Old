import datetime

date = datetime.datetime.now().strftime("%d-%m-%Y")

acc_excel_path = r"D:\Durai\Scraping\Accessories\Save Data's\Final Files\Accessories Price List " + date + ".xlsx"
lap_excel_path = r"D:\Durai\Scraping\Laptop\Save Data's\Final Files\Laptop Price Lists " + date + ".xlsx"
mob_excel_path = r"D:\Durai\Scraping\Mobile\Save Data\Final Files\Mobiles_Price_List " + date + ".xlsx"
tv_excel_path = r"D:\Durai\Scraping\Tv\Save Data\Final Files\Tv Price List " + date + ".xlsx"
kit_app_excel_path = r"D:\Durai\Scraping\Kitchen_appliances\Save Data\Final Files\Kitchen Appliance Price List " + date + ".xlsx"

acc_save_path = r"D:\Durai\Scraping\Least_Price\Save Data\Least Price Accessories Price List " + date + ".xlsx"
lap_save_path = r"D:\Durai\Scraping\Least_Price\Save Data\Least Price Laptop Price Lists " + date + ".xlsx"
mob_save_path = r"D:\Durai\Scraping\Least_Price\Save Data\Least Price Mobiles_Price_List " + date + ".xlsx"
tv_save_path = r"D:\Durai\Scraping\Least_Price\Save Data\Least Price Tv Price List " + date + ".xlsx"
kit_app_save_path = r"D:\Durai\Scraping\Least_Price\Save Data\Least Price Kitchen Appliance " + date + ".xlsx"


l = [acc_excel_path,lap_excel_path,mob_excel_path,tv_excel_path,kit_app_excel_path]

for r in range(1,6):
    # print(type(l[r]))
    if r == 1:

        excel_path = acc_excel_path
        save_path = acc_save_path
        print(" Done Accessories")

    elif r == 2:
        excel_path = lap_excel_path
        save_path = lap_save_path
        print(" Done Laptop")

    elif r == 3:
        excel_path = mob_excel_path
        save_path = mob_save_path
        print(" Done Mobile")

    elif r == 4:
        excel_path = tv_excel_path
        save_path = tv_save_path
        print(" Done Tv")

    elif r == 5:
        excel_path = kit_app_excel_path
        save_path = kit_app_save_path
        print(" Done Kitchen Appliance")
