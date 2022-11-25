from Scraping.Accessories.Py.Form.Acc_Poorvika import Model


def scraping_laptop():

    heading = "Laptops"
    url = "https://www.poorvika.com/s?catagories=categories.lvl1%3A%5B%60Computers+%26+Tablets+%3E+Laptops%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    lap_range = 1
    lap = Model(heading=heading, url=url, row_num=lap_range)
    lap.product_page_list(head=heading)