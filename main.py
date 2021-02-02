import mammoth
import codecs
from bs4 import BeautifulSoup
from xml.sax import saxutils as su
import re
import sys
import json

def docxTohtmlwithClasses(input_filename,output_filename,outputfilename):
	# ---------Read docx and converting to html with before sentence split conditions --------------#
	f = open(input_filename, 'rb')
	b = open(outputfilename, 'wb')
	document = mammoth.convert_to_html(f)
	b.write(document.value.encode('utf8'))
	f.close()
	b.close()
	f = codecs.open(outputfilename, 'rb', encoding='utf-8', errors='ignore')
	html_doc = f.read()
	soup = BeautifulSoup(html_doc, 'html.parser')
	for tag in soup("img"):
		tag.decompose()
	strhtm = soup.prettify()
	
	### List ###
	ul_tag = soup.findAll('ul')
	ol_tag = soup.findAll('ol')
	for each_ulTag in ul_tag:
		each_ulTag['class'] = 'slideTitle-False'
		hr_tag = soup.new_tag("hr", attrs={'class': 'listColor'})
		each_ulTag.insert_after(hr_tag)
	for each_olTag in ol_tag:
		each_olTag['class'] = 'slideTitle-False'
		hr_tag = soup.new_tag("hr", attrs={'class': 'listColor'})
		each_olTag.insert_after(hr_tag)
	
	### Table ###
	table_tag = soup.findAll('table')
	for each_tableTag in table_tag:
		each_tableTag['class'] = 'slideTitle-False'
		each_tableTag.replace_with(
			str(each_tableTag).replace('<table>', '<hr class ="tableColor"><table>').replace('</table>',
			                                                                                 '</table><hr class ="tableColor">'))
	
	### Heading ###
	h_tag = soup.findAll(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
	for each_hTag in h_tag:
		each_hTag['class'] = 'slideTitle-True'
		hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
		each_hTag.insert_before(hr_tag)
	
	##p_tag data
	ptag = soup.findAll('p')
	
	# ### End of paragraph ###
	for each_pTag in ptag:
		# print(":::", each_pTag)
		if len(each_pTag.get_text(strip=True)) == 0:
			each_pTag.decompose()
		
		if '</strong></p>' in str(each_pTag) or '</em></p>' in str(each_pTag) or '</em></strong></p>' in str(each_pTag):
			each_pTag['class'] = 'slideTitle-True'
			hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
			each_pTag.insert_before(hr_tag)
		
		elif ': </em></p>' in each_pTag or ':</em></strong></p>' in each_pTag:
			each_pTag['class'] = 'slideTitle-True'
			hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
			each_pTag.insert_before(hr_tag)
		
		elif '<p> </p>' in each_pTag:
			each_pTag.decompose()
		
		elif '<None></None>' in str(each_pTag):
			each_pTag.decompose()
		
		else:
			if each_pTag.text.endswith(':') == True or each_pTag.text.endswith(': ') == True:
				each_pTag['class'] = 'slideTitle-True'
				hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
				each_pTag.insert_before(hr_tag)
			else:
				each_pTag['class'] = 'slideTitle-False'
				hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
				each_pTag.insert_after(hr_tag)
	
	strhtm1 = su.unescape(str(soup))
	with open(output_filename, "w", encoding='utf-8') as f1:
		f1.write(strhtm1)
		f1.write(
			'<style> .headingColor{border-top : 1px solid red;} .listColor{border-top : 1px solid yellow;} .paraColor{border-top : 1px solid black;} .tableColor{border-top : 1px solid blue;} </style>')
	f1.close()
	
	# placing the split in the para's ###
	
	afterSplitFile = codecs.open(output_filename, 'rb', encoding='utf-8', errors='ignore')
	html_doc1 = afterSplitFile.read()
	soup1 = BeautifulSoup(html_doc1, 'html.parser')
	
	for x in soup1.findAll():
		if len(x.get_text(strip=True)) == 1:
			x.decompose()
	
	P_TAG = soup1.findAll('p')
	H_TAG = soup1.findAll(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
	
	
	# heading split
	def split_into_sentences_hTag(text):
		text = text.replace('&amp;', '&')
		if 'h1' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h1><h1>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h1><h1>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h1><h1>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h1><h1>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h1><h1>", text)
		if 'h2' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h2><h2>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h2><h2>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h2><h2>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h2><h2>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h2><h2>", text)
		if 'h3' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h3><h3>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h3><h3>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h3><h3>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h3><h3>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h3><h3>", text)
		if 'h4' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h4><h4>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h4><h4>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h4><h4>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h4><h4>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h4><h4>", text)
		if 'h5' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h5><h5>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h5><h5>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h5><h5>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h5><h5>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h5><h5>", text)
		if 'h6' in text:
			text = re.sub(r'(?<=\w\.)\s', "</h6><h6>", text)
			text = re.sub(r'(?<=\w\!)\s', "</h6><h6>", text)
			text = re.sub(r'(?<=\w\?)\s', "</h6><h6>", text)
			text = re.sub(r'(?<=\w\;)\s', "</h6><h6>", text)
		# text = re.sub(r'(?<=\w\:)\s', "</h6><h6>", text)
		return text
	
	
	for eachHTag in H_TAG:
		eachHTag['class'] = 'slideTitle-True'
		do1 = split_into_sentences_hTag(str(eachHTag))
		eachHTag.replace_with(do1)
	
	alphabets = "([A-Za-z])"
	prefixes = "(Mr|St|Mrs|Ms|Dr)[.]"
	suffixes = "(Inc|Ltd|Jr|Sr|Co)"
	starters = "(Mr|Mrs|Ms|Dr|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
	acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
	websites = "[.](com|net|org|io|gov)[.]*"
	websites_1 = r'href=[\'"]?([^\'" >]+)'
	name_initials = "([A-Z])[.]"
	special_char = '[!%")>?}:][.]'
	special_char_semiColon = '[!%")>?}:][;]'
	list_no = "^[0-9A-Za-z][.]"
	strongRegex = '</strong>[ ]*[–!%"><;?-},(){}:]'
	
	
	def split_into_sentences_pTag(text):
		# print(text)
		text = text.replace('&amp;', '&')
		
		# new  added
		sent = re.split('(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!|\;)(\s|[A-Z][0-9]\).*)', text)
		sent_new = []
		for i in sent:
			# print("::",i)
			if len(i) == 1:
				i = i.replace(i, '</p><p class="slideTitle-False">')
				sent_new.append(i)
			else:
				sent_new.append(i)
		return ''.join(sent_new)
		
	# paragraph spliting
	for eachPTAG in P_TAG:
		# print(eachPTAG)
		# eachPTAG['class'] = 'slideTitle-False'
		strong_tag = eachPTAG.find('strong')
		# print(strong_tag)
		if len(eachPTAG.get_text(strip=True)) == 0:
			eachPTAG.replace_with(
				str(eachPTAG).replace('<p></p>', ''))
		
		# list numbers/alphabets
		if re.search(r'^[\t]*[0-9A-Z]+[.]', eachPTAG.text) != None:
			eachPTAG['class'] = 'slideTitle-False'
			list_no_alpha = eachPTAG
			eachPTAG.replace_with(list_no_alpha)
		
		elif strong_tag != None:
			if '</strong>:</p>' in eachPTAG:
				eachPTAG['class'] = 'slideTitle-True'
				do1 = split_into_sentences_pTag(str(eachPTAG))
				eachPTAG.replace_with(do1)
			else:
				# eachPTAG['class'] = 'slideTitle-True'
				text1 = str(eachPTAG).replace('</strong>.', '</strong>.</p><p class="slideTitle-False">').replace('.</strong>', '.</strong></p><p class="slideTitle-False">').replace('<strong> </strong>', ' ')
				do1 = split_into_sentences_pTag(text1)
				eachPTAG.replace_with(do1)
		else:
			if (eachPTAG).text.endswith(':'):
				eachPTAG['class'] = 'slideTitle-True'
			else:
				eachPTAG['class'] = 'slideTitle-False'
				text = str(eachPTAG)
				if "Ph.D" in text: text = text.replace("Ph.D.", "Ph<prd>D<prd>")
				if "Dr." in text: text = text.replace("Dr.", "Dr<prd>")
				text = re.sub("\s" + alphabets + "[.] ", " \\1<prd>", text)
				text = re.sub(acronyms + " " + starters, "\\1<stop> \\2", text)
				text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>", text)
				text = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", text)
				text = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", text)
				text = re.sub(" " + suffixes + "[.]", " \\1<prd>", text)
				text = re.sub(" " + alphabets + "[.]", " \\1<prd>", text)
				text = text.replace("<prd>", ".")
				text = text.replace("<stop>", ".")
				do1 = split_into_sentences_pTag(text)
				eachPTAG.replace_with(do1)
	
	strhtmNew = su.unescape(str(soup1))
	return strhtmNew


def htmlTojson(outputfilename):
	f = codecs.open(outputfilename, 'rb', encoding='utf-8', errors='ignore')
	html_doc = f.read()
	soup = BeautifulSoup(html_doc, 'html.parser')
	presentation = []
	dic_data = {}
	testLst = []
	for i in soup.findAll(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol']):
		testLst.append(i)
	titlelst = []
	sentlst = []
	otherlst = []
	for i in range(len(testLst)):
		if testLst[i]['class'] == ['slideTitle-True']:
			titlelst.append(i)
		elif testLst[i]['class'] == ['slideTitle-False']:
			sentlst.append(i)
		else:
			otherlst.append(i)
	otherTrueCase = sentlst[-1] + 1
	titlelst.append(otherTrueCase)
	pairLst = list(zip(titlelst[:-1], titlelst[1:]))
	
	lst = []
	j = 0
	for m in range(len(titlelst) - 1):
		dic_data['slideTitle'] = testLst[titlelst[m]].text
		for i in range(titlelst[m] + 1, titlelst[m + 1]):
			lst.append(testLst[i].text)
		dic_data['sentences'] = lst
		lst = []
		presentation.append(dic_data)
		dic_data = {}
	return presentation


if __name__ == "__main__":
	input_filename = sys.argv[1]
	jsonFile = sys.argv[2]
	
	output_filename = 'Testing.html'
	outputfilename = 'filenametest.html'
	result = docxTohtmlwithClasses(input_filename,output_filename,outputfilename)
	with open(output_filename, "w", encoding='utf-8') as f1:
		f1.write(result)
	response = htmlTojson(output_filename)
	with open(jsonFile, "w", encoding='utf-8') as outfile:
		json.dump(response, outfile,ensure_ascii=False)
