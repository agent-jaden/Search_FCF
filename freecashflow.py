#-*- coding:utf-8 -*-
# Parsing data from comp.fnguide
import urllib.request
import xlrd
import xlsxwriter
import os
from bs4 import BeautifulSoup

### PART I - Read Excel file
num_stock = 2003
#num_stock = 100
input_file = "basic_20170729.xlsx"
cur_dir = os.getcwd()

workbook = xlrd.open_workbook(os.path.join(cur_dir, input_file))
sheet_list = workbook.sheets()
sheet1 = sheet_list[0]

stock_cat_list = []
stock_name_list = []
stock_num_list = []
stock_url_list = []
stock_url2_list = []

net_income_list = []
per_list = []
pbr_list = []
pdr_list = []
op_cf_list = []
capex_list = []
fcf_list = []
market_cap_list = []
close_price_list = []
error_list = []

for i in range(num_stock):
#for i in range(1900,2000):
	stock_cat_list.append(sheet1.cell(i+1,0).value)
	stock_name_list.append(sheet1.cell(i+1,1).value)
	stock_num_list.append(int(sheet1.cell(i+1,2).value))
	url="http://comp.fnguide.com/SVO2/ASP/SVD_Finance.asp?pGB=1&gicode=A" + sheet1.cell(i+1,2).value + "&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701"
	#url="http://comp.fnguide.com/SVO2/ASP/SVD_Invest.asp?pGB=1&gicode=A" + sheet1.cell(i+1,2).value + "&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701"
	url2="http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A" + sheet1.cell(i+1,2).value + "&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701"
	
	stock_url_list.append(url)
	stock_url2_list.append(url2)

for j in range(num_stock):
	print(j, stock_name_list[j])
	if j%10 == 0: print (j)
	
#	if j%100 == 0:
#
#		l=0
#		while(l<100000):
#			print (l)
#			l = l+1	

	# Read from URL #1
	url = stock_url_list[j]
	#print(url)

	handle = None
	while handle == None:
		try:
			handle = urllib.request.urlopen(url)
			#print(handle)
		except:
			pass

	data = handle.read()
	#print(data)
	soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

	cashflow_statement = soup.find('div', {'id': 'divCashY'})

	if cashflow_statement != None:
	
		trs = cashflow_statement.find_all('tr', {'class': "rwf rowBold"})
		
		if len(trs) != 0:
			#Operating Cash Flow
			tds = trs[0].find_all('td')
			
			if tds[3].text == 'N/A(IFRS)' or tds[3].text == '\xa0': 
				op_cf_2017 = 0
			else:
				op_cf_2017 = int(tds[3].text.replace(',',''))
			if tds[2].text == 'N/A(IFRS)' or tds[2].text == '\xa0': 
				op_cf_2016 = 0
			else:
				op_cf_2016 = int(tds[2].text.replace(',',''))
			if tds[1].text == 'N/A(IFRS)' or tds[1].text == '\xa0': 
				op_cf_2015 = 0
			else:
				op_cf_2015 = int(tds[1].text.replace(',',''))
			if tds[0].text == 'N/A(IFRS)' or tds[0].text == '\xa0': 
				op_cf_2014 = 0
			else:
				op_cf_2014 = int(tds[0].text.replace(',',''))

			op_cf_list.append([op_cf_2014, op_cf_2015, op_cf_2016, op_cf_2017])
			
			capex_trs = cashflow_statement.find_all('tr',{'class':"c_grid3_11 rwf acd_dep2_sub"})

			# Tangible asset
			tds = capex_trs[8].find_all('td')
			if tds[3].text == 'N/A(IFRS)' or tds[3].text == '\xa0': 
				capex1_2017 = 0
			else:
				capex1_2017 = int(tds[3].text.replace(',',''))
			if tds[2].text == 'N/A(IFRS)' or tds[2].text == '\xa0': 
				capex1_2016 = 0
			else:
				capex1_2016 = int(tds[2].text.replace(',',''))
			if tds[1].text == 'N/A(IFRS)' or tds[1].text == '\xa0': 
				capex1_2015 = 0
			else:
				capex1_2015 = int(tds[1].text.replace(',',''))
			if tds[0].text == 'N/A(IFRS)' or tds[0].text == '\xa0': 
				capex1_2014 = 0
			else:
				capex1_2014 = int(tds[0].text.replace(',',''))

			# Non Tangible asset
			tds = capex_trs[9].find_all('td')
			if tds[3].text == 'N/A(IFRS)' or tds[3].text == '\xa0': 
				capex2_2017 = 0
			else:
				capex2_2017 = int(tds[3].text.replace(',',''))
			if tds[2].text == 'N/A(IFRS)' or tds[2].text == '\xa0': 
				capex2_2016 = 0
			else:
				capex2_2016 = int(tds[2].text.replace(',',''))
			if tds[1].text == 'N/A(IFRS)' or tds[1].text == '\xa0': 
				capex2_2015 = 0
			else:
				capex2_2015 = int(tds[1].text.replace(',',''))
			if tds[0].text == 'N/A(IFRS)' or tds[0].text == '\xa0': 
				capex2_2014 = 0
			else:
				capex2_2014 = int(tds[0].text.replace(',',''))

			# Land Asset
			tds = capex_trs[10].find_all('td')
			if tds[3].text == 'N/A(IFRS)' or tds[3].text == '\xa0': 
				capex3_2017 = 0
			else:
				capex3_2017 = int(tds[3].text.replace(',',''))
			if tds[2].text == 'N/A(IFRS)' or tds[2].text == '\xa0': 
				capex3_2016 = 0
			else:
				capex3_2016 = int(tds[2].text.replace(',',''))
			if tds[1].text == 'N/A(IFRS)' or tds[1].text == '\xa0': 
				capex3_2015 = 0
			else:
				capex3_2015 = int(tds[1].text.replace(',',''))
			if tds[0].text == 'N/A(IFRS)' or tds[0].text == '\xa0': 
				capex3_2014 = 0
			else:
				capex3_2014 = int(tds[0].text.replace(',',''))

			capex_2014 = capex1_2014 + capex2_2014 + capex3_2014
			capex_2015 = capex1_2015 + capex2_2015 + capex3_2015
			capex_2016 = capex1_2016 + capex2_2016 + capex3_2016
			capex_2017 = capex1_2017 + capex2_2017 + capex3_2017

			capex_list.append([capex_2014, capex_2015, capex_2016, capex_2017])

			fcf_2014 = op_cf_2014 - capex_2014
			fcf_2015 = op_cf_2015 - capex_2015
			fcf_2016 = op_cf_2016 - capex_2016
			fcf_2017 = op_cf_2017 - capex_2017

			fcf_list.append([fcf_2014, fcf_2015, fcf_2016, fcf_2017])
		else:
			op_cf_list.append([0, 0, 0, 0])
			capex_list.append([0, 0, 0, 0])
			fcf_list.append([0, 0, 0, 0])
	else:
		op_cf_list.append([0, 0, 0, 0])
		capex_list.append([0, 0, 0, 0])
		fcf_list.append([0, 0, 0, 0])

	income_statement = soup.find('div', {'id':'divSonikY'})
	income_statement = soup.find('div', {'id':'divSonikY'})

	if income_statement != None:
		trs = income_statement.findAll('tr', {'class': "rwf rowBold"})

		if len(trs) == 4:
			# Net Income
			net_income_tds = trs[3].findAll('td')
			if net_income_tds[3].text == 'N/A(IFRS)' or net_income_tds[3].text == '\xa0': 
				recent_net_income = 0
			else:
				recent_net_income = int(net_income_tds[3].text.replace(',',''))
			if net_income_tds[2].text == 'N/A(IFRS)' or net_income_tds[2].text == '\xa0': 
				former_net_income = 0
			else:
				former_net_income = int(net_income_tds[2].text.replace(',',''))
		elif len(trs) == 3:
			# Net Income
			net_income_tds = trs[2].findAll('td')
			if net_income_tds[3].text == 'N/A(IFRS)' or net_income_tds[3].text == '\xa0': 
				recent_net_income = 0
			else:
				recent_net_income = int(net_income_tds[3].text.replace(',',''))
			if net_income_tds[2].text == 'N/A(IFRS)' or net_income_tds[2].text == '\xa0': 
				former_net_income = 0
			else:
				former_net_income = int(net_income_tds[2].text.replace(',',''))
				
		else:
			print("error", len(trs))
	else:
		recent_net_income = 0
		former_net_income = 0
		error_list.append(stock_name_list[j])

	net_income_list.append([former_net_income, recent_net_income])

	# Read from URL #2
	url2 = stock_url2_list[j]

	handle = None
	while handle == None:
		try:
			handle = urllib.request.urlopen(url2)
			#print(handle)
		except:
			pass

	data = handle.read()
	#print(data)
	soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

	#corpinfo = soup.find('div', {'class':'section ul_corpinfo'})
	corpinfo = soup.find('div', {'class':'corp_group2'})
	dds = corpinfo.findAll('dd')

	if dds[1].text == 'N/A(IFRS)' or dds[1].text == '\xa0' or dds[1].text == '-':
		per = 0.0
	else:
		per = float(dds[1].text.replace(',',''))
	if dds[7].text == 'N/A(IFRS)' or dds[7].text == '\xa0' or dds[7].text == '-':
		pbr = 0.0
	else:
		pbr = float(dds[7].text.replace(',',''))
	if dds[9].text == 'N/A(IFRS)' or dds[9].text == '\xa0' or dds[9].text == '-%':
		pdr = 0.0
	else:
		pdr = float(dds[9].text.replace('%','')) / 100

	per_list.append(per)
	pbr_list.append(pbr)
	pdr_list.append(pdr)

	ul_de = soup.find('div', {'id':'svdMainGrid1'})
	#print(ul_de)
	trs = ul_de.findAll('tr')
	td = trs[3].find('td')
	market_cap = float(td.text.replace(',',''))
	market_cap_list.append(market_cap)

	#spans = ul_de.findAll('span')
	#print(len(spans))

	main_chart = soup.find('span', {'id':'svdMainChartTxt11'})
	close_price = int(main_chart.text.replace(',',''))
	close_price_list.append(close_price)

print(error_list)

# Write an Excel file2
workbook_name = "Search_FCF.xlsx"

workbook = xlsxwriter.Workbook(workbook_name)
if os.path.isfile(os.path.join(cur_dir, workbook_name)):
	os.remove(os.path.join(cur_dir, workbook_name))
workbook = xlsxwriter.Workbook(workbook_name)

worksheet_result = workbook.add_worksheet('result')
filter_format = workbook.add_format({'bold':True,
									'fg_color': '#D7E4BC'
									})

percent_format = workbook.add_format({'num_format': '0.00%'})

roe_format = workbook.add_format({'bold':True,
								  'underline': True,
								  'num_format': '0.00%'})

num_format = workbook.add_format({'num_format':'0.00'})
num2_format = workbook.add_format({'num_format':'#,##0'})
num3_format = workbook.add_format({'num_format':'#,##0.00',
								  'fg_color':'#FCE4D6'})


worksheet_result.set_column('A:A', 10)
worksheet_result.set_column('B:B', 20)

worksheet_result.autofilter(0,0,2003,26)

worksheet_result.write(0, 0, "분류", filter_format)
worksheet_result.write(0, 1, "종목명", filter_format)
worksheet_result.write(0, 2, "PER")
worksheet_result.write(0, 3, "PBR")
worksheet_result.write(0, 4, "시가배당률")
worksheet_result.write(0, 5, "시가총액")
worksheet_result.write(0, 6, "OP Cashflow 2014")
worksheet_result.write(0, 7, "OP Cashflow 2015")
worksheet_result.write(0, 8, "OP Cashflow 2016")
worksheet_result.write(0, 9, "OP Cashflow 2017")
worksheet_result.write(0, 10, "순이익 2016")
worksheet_result.write(0, 11, "순이익 2017")
worksheet_result.write(0, 12, "CAPEX 2014")
worksheet_result.write(0, 13, "CAPEX 2015")
worksheet_result.write(0, 14, "CAPEX 2016")
worksheet_result.write(0, 15, "CAPEX 2017")
worksheet_result.write(0, 16, "FCF 2014")
worksheet_result.write(0, 17, "FCF 2015")
worksheet_result.write(0, 18, "FCF 2016")
worksheet_result.write(0, 19, "FCF 2017")
worksheet_result.write(0, 20, "P/FCF (2014)")
worksheet_result.write(0, 21, "P/FCF (2015)")
worksheet_result.write(0, 22, "P/FCF (2016)")
worksheet_result.write(0, 23, "P/FCF (2017E)")

for k in range(num_stock):
	worksheet_result.write(k+1, 0, stock_cat_list[k])
	worksheet_result.write(k+1, 1, stock_name_list[k])
	worksheet_result.write(k+1, 2, per_list[k])
	worksheet_result.write(k+1, 3, pbr_list[k])
	worksheet_result.write(k+1, 4, pdr_list[k], percent_format)
	worksheet_result.write(k+1, 5, market_cap_list[k])
	worksheet_result.write(k+1, 6, op_cf_list[k][0])
	worksheet_result.write(k+1, 7, op_cf_list[k][1])
	worksheet_result.write(k+1, 8, op_cf_list[k][2])
	worksheet_result.write(k+1, 9, op_cf_list[k][3])
	worksheet_result.write(k+1, 10, net_income_list[k][0])
	worksheet_result.write(k+1, 11, net_income_list[k][1])
	worksheet_result.write(k+1, 12, capex_list[k][0])
	worksheet_result.write(k+1, 13, capex_list[k][1])
	worksheet_result.write(k+1, 14, capex_list[k][2])
	worksheet_result.write(k+1, 15, capex_list[k][3])
	worksheet_result.write(k+1, 16, fcf_list[k][0])
	worksheet_result.write(k+1, 17, fcf_list[k][1])
	worksheet_result.write(k+1, 18, fcf_list[k][2])
	worksheet_result.write(k+1, 19, fcf_list[k][3])
	if fcf_list[k][0] > 0:
		worksheet_result.write(k+1, 20, float(market_cap_list[k]/fcf_list[k][0]), num_format)
	else:
		worksheet_result.write(k+1, 20, 0.0, num_format)
	if fcf_list[k][1] > 0:
		worksheet_result.write(k+1, 21, float(market_cap_list[k]/fcf_list[k][1]), num_format)
	else:
		worksheet_result.write(k+1, 21, 0.0, num_format)
	if fcf_list[k][2] > 0:
		worksheet_result.write(k+1, 22, float(market_cap_list[k]/fcf_list[k][2]), num_format)
	else:
		worksheet_result.write(k+1, 22, 0.0, num_format)
	if fcf_list[k][3] > 0:
		worksheet_result.write(k+1, 23, float(market_cap_list[k]/(fcf_list[k][3]*2)), num_format)
	else:
		worksheet_result.write(k+1, 23, 0.0, num_format)




