from Scraping.Accessories.Py.Form.Form_All import RunCompression

# Range 175 to 350

Rc = RunCompression()
Rc.run__all(head="Accessories",start=201,end=501,path=2)
