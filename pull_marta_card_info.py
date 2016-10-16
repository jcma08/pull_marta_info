import pandas as pd
from splinter import Browser
import time


def run(filename=True):

	########################################################################
	# read data

	#filename='MARTA Breeze card numbers.xlsx'
	#filename='MARTA Breeze card numbers_1.xls'
	dataset=pd.read_excel(filename,sheetname=0,header=None)


	########################################################################
	# input params

	url='https://balance.breezecard.com/breezeWeb/jsp/web/cardnumberweb.jsp'

	columns=[
		'cardnumber',
		'protected_balance',
		'expiration_date',
		'product_name',
		'remaining_rides',
		'stored_value',
		'pending_autoload_transactions']
	cardinformation=pd.DataFrame(columns=columns)

	n = dataset.shape[0]
	
	for i in range(n):

		cardnumber=dataset.ix[i].values[0]
		#cardnumer='0164 1487 1502 5743 2323'
		if( len(cardnumber) < 16 ):
			print "invalid card length: %s\n" % cardnumber
			temp_df=pd.DataFrame([[
				cardnumber,
				'NA',
				'NA',
				'NA',
				'NA',
				'NA',
				'NA']],
				columns=columns)
			cardinformation=cardinformation.append(temp_df)
			continue

		browser = Browser('chrome')
		browser.visit(url)
		browser.fill('cardnumber', cardnumber)
		browser.find_by_name('submitButton').click()
		text = browser.find_by_tag('tr')
		# 2. breezecard bulk of information
		temp_txt = text[2].value.split('\n')

		if len(temp_txt) == 11:
			
			txt = [cardnumber]

			
			# 5. is our card balance protected?
			temp_val = temp_txt[2].split(':')[1]
			txt.append(temp_val)
			
			# 6. card expiration date
			temp_val = temp_txt[3].split(':')[1]
			txt.append(temp_val)
			
			# 7. product name
			temp_val = temp_txt[5]
			txt.append(temp_val)
			
			# 9. remaining rides
			temp_val = temp_txt[6]
			txt.append(temp_val)
			
			# 11. store value
			temp_val = temp_txt[7].split(':')[1]
			txt.append(temp_val)
			
			# 11. store value
			temp_val = temp_txt[9]
			txt.append(temp_val)


			temp_df=pd.DataFrame([txt],columns=columns)
			cardinformation = cardinformation.append(temp_df)
			
		elif len(temp_txt) == 10:
			txt = [cardnumber]

			# 5. is our card balance protected?
			temp_val = temp_txt[2].split(':')[1]
			txt.append(temp_val)
			
			# 6. card expiration date
			temp_val = temp_txt[3].split(':')[1]
			txt.append(temp_val)
			
			# 7. product name
			temp_val = temp_txt[5]
			txt.append(temp_val)
			
			# 9. remaining rides
			temp_val = 0
			txt.append(temp_val)
			
			# 11. store value
			temp_val = temp_txt[6].split(':')[1]
			txt.append(temp_val)
			
			# 11. store value
			temp_val = temp_txt[8]
			txt.append(temp_val)

			temp_df=pd.DataFrame([txt],columns=columns)
			cardinformation = cardinformation.append(temp_df)

		else:
			temp_df=pd.DataFrame([[
					cardnumber,
					'NA',
					'NA',
					'NA',
					'NA',
					'NA',
					'NA']],columns=columns)
			cardinformation=cardinformation.append(temp_df)
	
		browser.quit()

	output = 'output_' + time.strftime("%H_%M_%S") + '.xlsx'
	cardinformation.to_excel(
		output,
		header=True,
		index=False
	)





########################################################################
# make api call






########################################################################
# write to file


if __name__=='__main__':
	# get filename
	filename=raw_input('Please type in the name of the file.\n')

	run(filename)


	raw_input('Press enter to close.')



