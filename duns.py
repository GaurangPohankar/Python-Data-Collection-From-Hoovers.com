import xlrd
import time
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import re
import csv

#NA.xlsx
name = input("Please Enter the File Name:")
#print(name)
offset = input("Please Enter the Time Offset:")
data_offset = input("Enter 1 For Company Name OR 3 For Dun's No.: ")
#print(offset)

workbook = xlrd.open_workbook(name)
sheet = workbook.sheet_by_index(0)

fields = ['BP_ID', 'BP_NAME1', 'BP_COUNTRY','LOCATION_DUNS_NUMBER','Company_Name','Location','Revenue','street_address','phone','website','address','type_role','employee','incorporated','credit','parent_com','link','contact_name1','contact_title1','contact_name2','contact_title2','contact_name3','contact_title3','category','category_info']
out_file = open('data.csv','w')
csvwriter = csv.DictWriter(out_file, delimiter=',', fieldnames=fields)
dict_service = {}

dict_service['BP_ID'] = 'BP_ID'
dict_service['BP_NAME1'] = 'BP_NAME1'
dict_service['BP_COUNTRY'] = 'BP_COUNTRY'
dict_service['LOCATION_DUNS_NUMBER'] = 'LOCATION DUNS NUMBER'
dict_service['Company_Name'] = 'Company Name'
dict_service['Location'] = 'Location'
dict_service['Revenue'] = 'Revenue'
dict_service['street_address'] = 'Street Address'
dict_service['phone'] = 'Phone'
dict_service['website'] = 'Website'
dict_service['address'] = 'Address'
dict_service['type_role'] = 'Role'
dict_service['employee'] = 'Employee'
dict_service['incorporated'] = 'Incorporated'
dict_service['credit'] = 'Credit'
dict_service['parent_com'] = 'Parent Com.Info'
dict_service['link'] = 'Link'
dict_service['contact_name1'] = 'Contact_name1'
dict_service['contact_title1'] = 'Contact_title1'
dict_service['contact_name2'] = 'Contact_name2'
dict_service['contact_title2'] = 'Contact_title2'
dict_service['contact_name3'] = 'Contact_name3'
dict_service['contact_title3'] = 'Contact_title3'
dict_service['category'] = 'Category'
dict_service['category_info'] = 'Category_info'
with open('data.csv', 'a') as csvfile:
     filewriter = csv.DictWriter(csvfile, delimiter=',', fieldnames=fields)
     filewriter.writerow(dict_service)
     csvfile.close()
     #Write row to CSV
     csvwriter.writerow(dict_service)

data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

#print(len(data))
#print(data[0])

driver = webdriver.Firefox(executable_path='./geckodriver')
driver.set_page_load_timeout(50)    
driver.maximize_window()
     
for i in range(len(data)):
     try:
          print("---------------------------------------------------------------------------")
          print("Dun No :"+data[i+1][int(data_offset)])
          log_url = "http://www.hoovers.com/company-information/cs.html?term="+data[i+1][int(data_offset)]
          driver.get(log_url)
          time.sleep(int(offset))
          whatsapp_page = driver.page_source
          whatsapp_soup = BeautifulSoup(whatsapp_page, "html.parser")
          table = whatsapp_soup.find_all("table")
          for mytable in table:
              table_body = mytable.find('tbody')
              try:
                  rows = table_body.find_all('tr')
                  for tr in rows:
                      cols = tr.find_all('td')
                      company = cols[0].text
                      location = cols[1].text
                      revenue = cols[2].text

                      dict_service['BP_ID'] = data[i+1][0]
                      dict_service['BP_NAME1'] = data[i+1][1]
                      dict_service['BP_COUNTRY'] = data[i+1][2]
                      dict_service['LOCATION_DUNS_NUMBER'] = data[i+1][3]
                      
                      dict_service['Company_Name'] = company
                      dict_service['Location'] = location
                      dict_service['Revenue'] = revenue
                      
                      print("Name :"+company)
                      print("Location :"+location)
                      print("Revenue :"+revenue)

                      try:
                           links_with_text = [a['href'] for a in cols[3].find_all('a', href=True) if a.text]
                           links_with_text = 'http://www.hoovers.com'+links_with_text[0]
                           driver.get(links_with_text)
                           time.sleep(int(offset))
                           com_info = driver.page_source
                           com_info = BeautifulSoup(com_info, "html.parser")
                               
                           street_address_1 = com_info.find_all("div", { "class" : "street_address_1" })
                           company_city = com_info.find_all("span", { "class" : "company_city" })
                           company_region = com_info.find_all("span", { "class" : "company_region" })
                           company_postal = com_info.find_all("span", { "class" : "company_postal" })
                           company_country = com_info.find_all("span", { "class" : "company_country" })
                           com_value = com_info.find_all("span", { "class" : "value"})
                           web = com_info.find_all("div", { "class" : "web" })
                           phone = com_info.find_all("div", { "class" : "phone" })
                           type_role = com_info.find_all("span", { "class" : "role" })
                           
                           try:
                                relations_details = com_info.find_all("div", { "class" : "relations_details" })
                                #print(relations_details)
                                company_relations_details = relations_details[0].find_all("div", { "class" : "company_name" })
                                #print(company_relations_details)
                           except:
                                print("Relations Not Found")

                           try:
                                contact_com = com_info.find_all("div", { "class" : "module_body contact-body" })
                                contact_info = contact_com[0].find_all("div", { "class" : "name" })
                                contact_title = contact_com[0].find_all("div", { "class" : "position sub" })
                           except:
                                print("Contacts Not Found")

                           try:
                                industry_name = com_info.find_all("div", { "class" : "industry_name" })
                                industry_description = com_info.find_all("div", { "class" : "industry_description" })
                           except:
                                print("Contacts Not Found")

                           try:
                                category = industry_name[0].text
                                category_info = industry_description[0].text
                           except:
                                category = "Not Found"
                                category_info = "Not Found"     

                           try:
                                ultimate_parent = company_relations_details[0].text
                           except:
                                ultimate_parent = " "

                           try:
                                sub_parent = company_relations_details[1].text
                           except:
                                sub_parent = " "

                           try:
                                branch_parent = company_relations_details[2].text
                           except:
                                branch_parent = " "
                                
                           try:
                                parent_com = ultimate_parent +","+ sub_parent + "," + branch_parent + ","
                           except:
                                try:
                                     parent_com = ultimate_parent +","+ sub_parent + ","
                                except:
                                     try:
                                          parent_com = ultimate_parent
                                     except:
                                          parent_com ="Not Found"
                                          
                                
                           try:
                                contact_name1 = contact_info[0].text
                                contact_title1 = contact_title[0].text
                           except:
                                contact_name1 = "Not Found"
                                contact_title1 = "Not Found"

                           try:
                                contact_name2 = contact_info[1].text
                                contact_title2 = contact_title[1].text
                           except:
                                contact_name2 = "Not Found"
                                contact_title2 = "Not Found"

                           try:
                                contact_name3 = contact_info[2].text
                                contact_title3 = contact_title[2].text
                           except:
                                contact_name3 = "Not Found"
                                contact_title3 = "Not Found"

                           try:
                                type_role = type_role[0].text
                           except:
                                type_role = "Not Found"

                           try:
                                employee = com_value[0].text
                           except:
                                employee = "Not Found"

                           try:
                                incorporated = com_value[2].text
                           except:
                                incorporated = "Not Found"

                           try:
                                credit = com_value[3].text
                           except:
                                credit = "Not Found"

                           try:
                                company_city = re.sub(' +', ' ', company_city[0].text) 
                           except:
                                company_city = "Not Found"

                           try:
                                company_region = re.sub(' +', ' ', company_region[0].text)
                           except:
                                company_region = "Not Found"
                                
                           try:
                                company_postal = re.sub(' +', ' ', company_postal[0].text)
                           except:
                                company_postal = "Not Found"
                           
                           try:
                                company_country = re.sub(' +', ' ', company_country[0].text)
                           except:
                                company_country = "Not Found"

                           try:
                                website = web[0].text
                           except:
                                website = "Not Found"

                           try:
                                phone = phone[0].text
                           except:
                                phone = "Not Found"

                           try:
                                street_address = street_address_1[0].text
                           except:
                                street_address = "Not Found"
                                
                           try:
                                address = company_region+","+company_city+","+company_postal+","+company_country
                           except:
                                address = "Not Found"


                           print("Street Address :"+ street_address)
                           dict_service['street_address'] = street_address
                           
                           print("Phone :"+phone)
                           dict_service['phone'] = phone
                           
                           print("Website :"+website)
                           dict_service['website'] = website
                           
                           print("Address :"+address)
                           dict_service['address'] = address
                           
                           print("Type Role :"+type_role)
                           dict_service['type_role'] = type_role
                           
                           print("Employee :"+employee)
                           dict_service['employee'] = employee
                           
                           print("Incorporated :"+incorporated)
                           dict_service['incorporated'] = incorporated
                           
                           print("Credit :"+credit)
                           dict_service['credit'] = credit
                           
                           print("Parent Com. And Info:"+parent_com)
                           dict_service['parent_com'] = parent_com
                           
                           print("Searching Link:"+links_with_text)
                           dict_service['link'] = links_with_text
                           
                           print("Person1 :"+contact_name1)
                           dict_service['contact_name1'] = contact_name1
                           
                           print("Title1 :"+contact_title1)
                           dict_service['contact_title1'] = contact_title1
                           
                           print("Person2 :"+contact_name2)
                           dict_service['contact_name2'] = contact_name2
                           
                           print("Title2 :"+contact_title2)
                           dict_service['contact_title2'] = contact_title2
                           
                           print("Person3 :"+contact_name3)
                           dict_service['contact_name3'] = contact_name3
                           
                           print("Title3 :"+contact_title3)
                           dict_service['contact_title3'] = contact_title3
                           
                           print("Category :"+category)
                           dict_service['category'] = category
                           
                           print("Category Info :"+category_info)
                           dict_service['category_info'] = category_info
                           with open('data.csv', 'a') as csvfile:
                                filewriter = csv.DictWriter(csvfile, delimiter=',', fieldnames=fields)
                                filewriter.writerow(dict_service)
                                csvfile.close()
                                #Write row to CSV
                                csvwriter.writerow(dict_service)
                           print('########################## Company Info Ended ########################')
                      except:
                           print('none')
                          
              except:
                  print("no tbody")
     except:
          print('no term found')
