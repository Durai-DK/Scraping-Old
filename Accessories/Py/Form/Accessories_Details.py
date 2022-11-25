from Scraping.Accessories.Py.Form.Acc_Poorvika import Model


def scraping_accessories():

    heading = "Mobile & Accessories"
    url = "https://www.poorvika.com/s?catagories=categories%3A%5B%60Mobiles+%26+Accessories%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    ma_range = 1
    ma = Model(heading=heading, url=url, row_num=ma_range)
    ma.product_page_list(head=heading)

########################################################################################################################

    heading = "Computer & Laptop Accessories"
    url = "https://www.poorvika.com/s?catagories=categories.lvl1%3A%5B%60Computers+%26+Tablets+%3E+Computer+%26+Laptop+Accessories%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    cl_range = ma.product_num()

    cl = Model(heading=heading, url=url, row_num=cl_range)
    cl.product_page_list(head=heading)

########################################################################################################################

    heading = "Tab & Ipad Accessories"
    url = "https://www.poorvika.com/s?catagories=categories.lvl1%3A%5B%60Computers+%26+Tablets+%3E+Tabs+%26+IPad+Accessories%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    tl_range = cl.product_num()

    tl = Model(heading=heading, url=url, row_num=tl_range)
    tl.product_page_list(head=heading)

########################################################################################################################

    heading = "TV & Audio Accessories"
    url = "https://www.poorvika.com/s?catagories=categories%3A%5B%60TV+%26+Audio%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    ta_range = tl.product_num()

    ta = Model(heading=heading, url=url, row_num=ta_range)
    ta.product_page_list(head=heading)

########################################################################################################################

    heading = "Smart Technology"
    url = "https://www.poorvika.com/s?catagories=categories%3A%5B%60Smart+Technology%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D&page="
    st_range = ta.product_num()

    st = Model(heading=heading, url=url, row_num=st_range)
    st.product_page_list(head=heading)

########################################################################################################################

    heading = "Personal & Health Care"
    url = "https://www.poorvika.com/s?catagories=categories%3A%5B%60Personal+%26+Health+Care%60%5D&stock_status=stock_status%3A%3D%5B%60In+Stock%60%5D"
    ph_range = st.product_num()

    ph = Model(heading=heading, url=url, row_num=ph_range)
    ph.product_page_list(head=heading)

########################################################################################################################