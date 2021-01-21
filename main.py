import mammoth
import codecs
from bs4 import BeautifulSoup
from xml.sax import saxutils as su
import re
import sys

#fetching parameters from command line arguments
input_filename = sys.argv[1]
output_filename = sys.argv[2]

outputfilename = 'filenametest.html'

f = open(input_filename, 'rb')
b = open(outputfilename, 'wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()

f = codecs.open(outputfilename, 'r', encoding='utf-8')
html_doc = f.read()
soup = BeautifulSoup(html_doc, 'html.parser')
for tag in soup("img"):
	tag.decompose()
strhtm = soup.prettify()

### List ###
ul_tag = soup.findAll('ul')
ol_tag = soup.findAll('ol')
for each_ulTag in ul_tag:
	each_ulTag.replace_with(str(each_ulTag).replace('</ul>', '</ul><hr class ="listColor">'))
for each_olTag in ol_tag:
	each_olTag.replace_with(str(each_olTag).replace('</ol>', '</ol><hr class ="listColor">'))

### Table ###
table_tag = soup.findAll('table')
for each_tableTag in table_tag:
	each_tableTag.replace_with(
		str(each_tableTag).replace('<table>', '<hr class ="tableColor"><table>').replace('</table>',
		                                                                                 '</table><hr class ="tableColor">'))

### Heading ###
h1_tag = soup.findAll("h1")
for each_h1Tag in h1_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h1Tag.insert_before(hr_tag)

h2_tag = soup.findAll("h2")
for each_h2Tag in h2_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h2Tag.insert_before(hr_tag)

h3_tag = soup.findAll("h3")
for each_h3Tag in h3_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h3Tag.insert_before(hr_tag)

h4_tag = soup.findAll("h4")
for each_h4Tag in h4_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h4Tag.insert_before(hr_tag)

h5_tag = soup.findAll("h5")
for each_h5Tag in h5_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h5Tag.insert_before(hr_tag)

h6_tag = soup.findAll("h6")
for each_h6Tag in h6_tag:
	hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
	each_h6Tag.insert_before(hr_tag)

for each_strongTag in soup.find_all('p'):
	para = each_strongTag.find_next_sibling('strong')
	hr_tag = soup.new_tag("hr", attrs={'class': 'strongTextColor'})
	each_strongTag.insert_before(hr_tag)

### End of paragraph ###
p_tag = soup.findAll('p')
for each_pTag in p_tag:
	each_pTag.replace_with(str(each_pTag).replace('</p>', '</p><hr class ="paraColor">'))

### Empty Paragraph ###
empty_ptag = soup.findAll('p')
for each_emptypTag in empty_ptag:
	each_emptypTag.replace_with(
		str(each_emptypTag).replace('<p></p>', '<p></p><hr class ="paraColor">').replace('<p>&nbsp</p>',
		                                                                                 '<p>&nbsp</p><hr class ="paraColor">'))

strhtm1 = su.unescape(str(soup))
text_file = open(output_filename, "w")
text_file.write(strhtm1)
text_file.write(
	'<style> .headingColor{border-top : 1px solid black;} .listColor{border-top : 1px solid yellow;} .paraColor{border-top : 1px solid green;} .tableColor{border-top : 1px solid blue;} .strongTextColor{border-top : 1px solid red;} </style>')
text_file.close()

###### taking the older html to perform operations after split #####

afterSplitFile = codecs.open(output_filename, 'r', encoding='utf-8')
html_doc1 = afterSplitFile.read()
soup1 = BeautifulSoup(html_doc1, 'html.parser')

P_TAG = soup1.findAll('p')
H_TAG = soup1.findAll(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])

alphabets = "([A-Za-z])"
prefixes = "(Mr|St|Mrs|Ms|Dr)[.]"
suffixes = "(Inc|Ltd|Jr|Sr|Co)"
starters = "(Mr|Mrs|Ms|Dr|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
websites = "[.](com|net|org|io|gov)"
websites_1 = "(www)[.]"


def split_into_sentences_pTag(text):
	text = re.sub(prefixes, "\\1<prd>", text)
	text = re.sub(websites, "<prd>\\1", text)
	text = re.sub(websites_1, "\\1<prd> ", text)
	if "Ph.D" in text: text = text.replace("Ph.D.", "Ph<prd>D<prd>")
	text = re.sub("\s" + alphabets + "[.] ", " \\1<prd> ", text)
	text = re.sub(acronyms + " " + starters, "\\1<stop> \\2", text)
	text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>", text)
	text = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", text)
	text = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", text)
	text = re.sub(" " + suffixes + "[.]", " \\1<prd>", text)
	text = re.sub(" " + alphabets + "[.]", " \\1<prd>", text)
	text = text.replace("<prd>", ".")
	if 'www.' in text:
		text = re.sub(websites_1, "\\1<prd> ", text)
	else:
		text = re.sub(r'(?<=\w\.)\s', "</p><p>", text)
		text = re.sub(r'(?<=\w\!)\s', "</p><p>", text)
		text = re.sub(r'(?<=\w\?)\s', "</p><p>", text)
		text = re.sub(r'(?<=\w\;)\s', "</p><p>", text)
	return text

def split_into_sentences_hTag(text):
	if 'h1' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h1><h1>", text)
	if 'h2' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h2><h2>", text)
	if 'h3' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h3><h3>", text)
	if 'h4' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h4><h4>", text)
	if 'h5' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h5><h5>", text)
	if 'h6' in text:
		text = re.sub(r'(?<=\w\.)\s', "</h6><h6>", text)
	return text


for eachPTAG in P_TAG:
	if eachPTAG.text.endswith(':') :
		hr_tag = soup1.new_tag("hr") #, attrs={'class': 'headingColor'})
		eachPTAG.insert_before(hr_tag)
	# print(eachPTAG.text)
	if re.search('^\d{1,3}.',eachPTAG.text) != None:
		do1 = eachPTAG
	else:
		do1 = split_into_sentences_pTag(str(eachPTAG))
	eachPTAG.replace_with(do1)
	
for eachHTag in H_TAG:
	do1 = split_into_sentences_hTag(str(eachHTag))
	eachHTag.replace_with(do1)

strhtmNew = su.unescape(str(soup1))
print(strhtmNew)
text_file = open(output_filename, "w")
text_file.write(strhtmNew)
text_file.close()

# For linux system, please uncomment : For windows , please comment
# os.remove(outputfilename)
