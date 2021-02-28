from flask import Flask, render_template, request, send_file
from bs4 import BeautifulSoup
import codecs
from xml.sax import saxutils as su
import re
import sys
import json
from itertools import groupby

app = Flask(__name__)

def docxTohtmlwithClasses(htmlGenerated,output_filename):
    # ---------Read docx and converting to html with before sentence split conditions --------------#
    soup = BeautifulSoup(htmlGenerated, 'html.parser')
    for x in soup.findAll():
        if len(x.get_text(strip=True)) == 0:
            x.decompose()

    for tag in soup():
        for attribute in ["class", "id", "name", "style","face", "align", "dir","lang","paraeid","paraid","xml:lang","role","cellspacing","cellpadding","colspan"]:
                    del tag[attribute]
    for tag in soup("img"):
        tag.decompose()

    for match in soup.findAll('span'):
        match.unwrap()
    for match1 in soup.findAll('font'):
        match1.unwrap()
    for match2 in soup.findAll('div'):
        match2.unwrap()
    for match3 in soup.findAll('br'):
        match3.unwrap()
    # for strong in soup.findAll('b'):
    
    strhtm = soup.prettify()

    ptag = soup.findAll('p')
    for eachPtag in ptag:
        eachPtag['class'] = 'slideTitle-False'
        eachPtag.replace_with(eachPtag)
    
    def split_into_sentences_Tag(text):
        # print(text)
        text = text.replace('&amp;', '&')
        # new  added
        sent = re.split('(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!|\;)(\s|[A-Z][0-9]\).*)', text)
        sent_new = []
        # print("SENTENCE:", sent)
        for i in sent:
            # print("::",i)
            if i != '\xa0' and i != '</p>' and i != '' and i != ' ':
                if '<p class="slideTitle-False">' in i:
                    a = "%s</%s>" % (i, 'p')
                else:
                    a = "<%s class='slideTitle-False'>%s</%s>" % ('p', i, 'p')
                sent_new.append(a)
            else:
                pass
        print("sent-new-before::",sent_new)
        # print(str(sent_new[-1]))
        if "<p class='slideTitle-False'>" in str(sent_new[-1]) or '<p class="slideTitle-False">' in str(sent_new[-1]):
            sent_new[-1] = str(sent_new[-1]).replace("<p class='slideTitle-False'>","<p class='slideTitle-True'>").replace('<p class="slideTitle-False">','<p class="slideTitle-True">')
        # print("sent-new-after::",sent_new)
        # print(''.join(sent_new))
        return ''.join(sent_new)
    
    ul_tag = soup.findAll('ul')
    ol_tag = soup.findAll('ol')
    
    for each_ulTag in ul_tag:
        siblingTag = each_ulTag.find_previous_sibling('p')
        # print(siblingTag)
        if siblingTag != None:
            siblingTag['class'] = 'slideTitle-False'
            # print("PulTAG:",siblingTag)
            t1 = split_into_sentences_Tag(str(siblingTag))
            
            if '<p class="slideTitle-True"></p>' in t1:
                siblingTag.replace_with(str(siblingTag).replace('<p class="slideTitle-True"></p>',''))
                siblingTag['class'] = 'slideTitle-True'
            elif '</p><p class="slideTitle-True">' in t1:
                siblingTag.replace_with(str(siblingTag).replace('</p><p class="slideTitle-True">','</p><p class="slideTitle-True">'))
            else:
                siblingTag['class'] = 'slideTitle-True'
                siblingTag.replace_with(t1)
        else:
            pass
        each_ulTag['class'] = 'slideTitle-False'
        hr_tag = soup.new_tag("hr", attrs={'class': 'listColor'})
        each_ulTag.insert_after(hr_tag)
    
    # print(soup)
    for each_olTag in ol_tag:
        siblingTag1 = each_olTag.find_previous_sibling('p')
        # print(siblingTag1)
        if siblingTag1 != None:
            siblingTag1['class'] = 'slideTitle-False'
            # print("PulTAG:", siblingTag)
            t11 = split_into_sentences_Tag(str(siblingTag1))
        
            if '<p class="slideTitle-True"></p>' in t11:
                siblingTag1.replace_with(str(siblingTag1).replace('<p class="slideTitle-True"></p>', ''))
                siblingTag1['class'] = 'slideTitle-True'
            elif '</p><p class="slideTitle-True">' in t11:
                siblingTag1.replace_with(
                    str(siblingTag1).replace('</p><p class="slideTitle-True">', '</p><p class="slideTitle-True">'))
            else:
                siblingTag1['class'] = 'slideTitle-True'
                siblingTag1.replace_with(t11)
        else:
            pass
        each_olTag['class'] = 'slideTitle-False'
        hr_tag = soup.new_tag("hr", attrs={'class': 'listColor'})
        each_olTag.insert_after(hr_tag)
    
    for eachOLtag in ol_tag:
        for OLTAG in eachOLtag.findAll('li'):
            OLTAG['class'] = 'slideTitle-False'
            OLTAG.replace_with(str(OLTAG).replace('<p>', '').replace('</p>', '').replace('<li>', '<li class="slideTitle-False">'))
            
    for eachULtag in ul_tag:
        for ULTAG in eachULtag.findAll('li'):
            ULTAG['class'] = 'slideTitle-False'
            ULTAG.replace_with(str(ULTAG).replace('<p>', '').replace('</p>', '').replace('<li>', '<li class="slideTitle-False">'))
    
    ### Table ###
    table_tag = soup.findAll('table')
    for each_tableTag in table_tag:
        each_tableTag['class'] = 'slideTitle-False'
        each_tableTag.replace_with(
            str(each_tableTag).replace('<table class="slideTitle-False">',
                                       '<hr class ="tableColor"><table class="slideTitle-False">').replace('</table>',
                                                                                                           '</table><hr class ="tableColor">').replace(
                '<td>', '<td class="slideTitle-False">'))
        
   
    ### Heading ###
    h_tag = soup.findAll(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
    for each_hTag in h_tag:
        each_hTag['class'] = 'slideTitle-True'
        hr_tag = soup.new_tag("hr", attrs={'class': 'headingColor'})
        each_hTag.insert_before(hr_tag)
    
    # print(soup)
    # ### End of paragraph ###
    print("SOUP::", soup)
    for each_pTag in ptag:
        # print(each_pTag)
        if '<p class="slideTitle-True"></p>' in str(each_pTag):
            each_pTag.replace_with(str(each_pTag).replace('<p class="slideTitle-True"></p>',''))
        
        else:
        
            if '<p class="slideTitle-True">' in str(each_pTag):
                each_pTag['class'] = 'slideTitle-True'
                # each_pTag.replace_with(each_pTag)
            else:
                # each_pTag['class'] = 'slideTitle-False'
                # else:
                # if ('</b></p>' in str(each_pTag) and '<p><b>' in str(each_pTag)) or ('</i></p>' in str(each_pTag) and '<p><i>' in str(each_pTag)) or ('</u></p>' in str(each_pTag) and '<p><u>' in str(each_pTag)) or '</u></i></p>' in str(each_pTag) or '</b></u></p>' in str(each_pTag) or '</u></b></p>' in str(each_pTag) :
                if ('</strong></p>' in str(each_pTag) and '<p class="slideTitle-False"><strong>' in str(each_pTag)) or ('</i></p>' in str(each_pTag) and '<p class="slideTitle-False"><i>' in str(each_pTag)) or ('</u></p>' in str(each_pTag) and '<p class="slideTitle-False"><u>' in str(each_pTag)) or '</u></i></p>' in str(each_pTag) or '</strong></u></p>' in str(each_pTag) or '</u></strong></p>' in str(each_pTag) :
                    each_pTag['class'] = 'slideTitle-True'
                    hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
                    each_pTag.insert_before(hr_tag)
                
                # elif ': </b></p>' in each_pTag or ':</b></p>' in each_pTag:
                elif ': </strong></p>' in each_pTag or ':</strong></p>' in each_pTag:
                    each_pTag['class'] = 'slideTitle-True'
                    hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
                    each_pTag.insert_before(hr_tag)
                
                else:
                    if each_pTag.text.endswith(':') == True or each_pTag.text.endswith(': ') == True:
                        each_pTag['class'] = 'slideTitle-True'
                        hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
                        each_pTag.insert_before(hr_tag)
                    else:
                        if each_pTag['class'] == 'slideTitle-True':
                            pass
                        else:
                            each_pTag['class'] = 'slideTitle-False'
                            hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
                            each_pTag.insert_after(hr_tag)
            # each_pTag.replace_with(str(each_pTag).replace('<p class="slideTitle-False"></p>', ''))
            # each_pTag.replace_with(str(each_pTag).replace('<p class="slideTitle-True"></p>', ''))

    
    strhtm1 = su.unescape(str(soup))
    print("strhtm1::",strhtm1)
    with open(output_filename, "w", encoding='utf-8') as f1:
        f1.write(strhtm1)
        f1.write(
            '<style> .headingColor{border-top : 1px solid red;} .listColor{border-top : 1px solid yellow;} .paraColor{border-top : 1px solid black;} .tableColor{border-top : 1px solid blue;} </style>')
    f1.close()
    
    # placing the split in the para's ###
    
    afterSplitFile = codecs.open(output_filename, 'rb', encoding='utf-8', errors='ignore')
    html_doc1 = afterSplitFile.read()
    soup1 = BeautifulSoup(html_doc1, 'html.parser')
   
    P_TAG = soup1.findAll('p')
    H_TAG = soup1.findAll(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
    
    # heading split
    def split_into_sentences_hTag(text):
        text = text.replace('&amp;', '&')
        if 'h1' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h1><h1 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h1><h1 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h1><h1 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h1><h1 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h1><h1>", text)
        if 'h2' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h2><h2 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h2><h2 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h2><h2 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h2><h2 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h2><h2>", text)
        if 'h3' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h3><h3 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h3><h3 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h3><h3 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h3><h3 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h3><h3>", text)
        if 'h4' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h4><h4 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h4><h4 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h4><h4 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h4><h4 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h4><h4>", text)
        if 'h5' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h5><h5 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h5><h5 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h5><h5 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h5><h5 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h5><h5>", text)
        if 'h6' in text:
            text = re.sub(r'(?<=\w\.)\s', "</h6><h6 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\!)\s', "</h6><h6 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\?)\s', "</h6><h6 class = 'slideTitle-True'>", text)
            text = re.sub(r'(?<=\w\;)\s', "</h6><h6 class = 'slideTitle-True'>", text)
        # text = re.sub(r'(?<=\w\:)\s', "</h6><h6>", text)
        return text
    
    for eachHTag in H_TAG:
        # eachHTag['class'] = 'slideTitle-True'
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
        if "<p class='slideTitle-True'></strong></p></p>" in str(eachPTAG):
            eachPTAG.replace_with(str(eachPTAG).replace("<p class='slideTitle-True'></strong></p></p>", ''))
            
        if '<p class="slideTitle-True"></p>' in str(eachPTAG):
            eachPTAG.replace_with(str(eachPTAG).replace('<p class="slideTitle-True"></p>',''))
        
        else:
            strong_tag = eachPTAG.find('strong')
            if len(eachPTAG.get_text(strip=True)) == 0:
                eachPTAG.replace_with(
                    str(eachPTAG).replace('<p></p>', ''))
            
            # list numbers/alphabets
            if re.search(r'^[\t]*[0-9A-Z]+[.]', eachPTAG.text) != None:
                list_no_alpha = eachPTAG
                eachPTAG.replace_with(list_no_alpha)
            
            elif strong_tag != None:
                if '</strong>:</p>' in eachPTAG:
                    eachPTAG['class'] = 'slideTitle-True'
                    do1 = split_into_sentences_pTag(str(eachPTAG).strip())
                    eachPTAG.replace_with(do1)
    
                elif ('</strong></p>' in str(eachPTAG) and '<p class="slideTitle-False"><strong>' in str(eachPTAG)) or (
                        '</i></p>' in str(eachPTAG) and '<p class="slideTitle-False"><i>' in str(eachPTAG)) or (
                        '</u></p>' in str(eachPTAG) and '<p class="slideTitle-False"><u>' in str(eachPTAG)) or '</u></i></p>' in str(
                        eachPTAG) or '</strong></u></p>' in str(eachPTAG) or '</u></strong></p>' in str(eachPTAG):
                    eachPTAG['class'] = 'slideTitle-True'
                    eachPTAG.replace_with(eachPTAG)
                
                else:
                    text1 = str(eachPTAG).replace('</strong>.', '</strong>.</p><p class="slideTitle-False">').replace(
                        '.</strong>', '.</strong></p><p class="slideTitle-False">').replace('<strong> </strong>', ' ')
                    do1 = split_into_sentences_pTag(text1)
                    eachPTAG.replace_with(do1)
            else:
                if (eachPTAG).text.endswith(':'):
                    eachPTAG['class'] = 'slideTitle-True'
                    eachPTAG.replace_with(eachPTAG)
                
                else:
                    # if '<p class="slideTitle-False"></p>' in str(eachPTAG):
                    #     eachPTAG.replace_with(str(eachPTAG).replace('<p class="slideTitle-False"></p>', ''))
                    #
                    # elif '<p class="slideTitle-True"></p>' in str(eachPTAG):
                    #     eachPTAG.replace_with(str(eachPTAG).replace('<p class="slideTitle-True"></p>', ''))
                    #
                    # else:
                    text = str(eachPTAG).strip()
                    if "Ph.D" in text: text = text.replace("Ph.D.", "Ph<prd>D<prd>")
                    if "Dr." in text: text = text.replace("Dr.", "Dr<prd>")
                    text = re.sub("\s" + alphabets + "[.] ", " \\1<prd>", text)
                    text = re.sub(acronyms + " " + starters, "\\1<stop> \\2", text)
                    text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>",
                                  text)
                    text = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", text)
                    text = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", text)
                    text = re.sub(" " + suffixes + "[.]", " \\1<prd>", text)
                    text = re.sub(" " + alphabets + "[.]", " \\1<prd>", text)
                    text = text.replace("<prd>", ".")
                    text = text.replace("<stop>", ".")
                    do1 = split_into_sentences_pTag(text)
                    eachPTAG.replace_with(do1)
    strhtmNew = su.unescape(str(soup1))
    # print(strhtmNew)
    
    with open(output_filename, "w", encoding='utf-8') as f1:
        f1.write(strhtmNew)
        f1.write(
            '<style> .headingColor{border-top : 1px solid red;} .listColor{border-top : 1px solid yellow;} .paraColor{border-top : 1px solid black;} .tableColor{border-top : 1px solid blue;} </style>')
    f1.close()

    afterSplitFile1 = codecs.open(output_filename, 'rb', encoding='utf-8', errors='ignore')
    html_doc11 = afterSplitFile1.read()
    soup11 = BeautifulSoup(html_doc11, 'html.parser')
    newPTag = soup11.findAll('p')
    
    for each_pTag in newPTag:
        if ('</strong></p>' in str(each_pTag) and '<p class="slideTitle-False"><strong>' in str(each_pTag)) or ('</i></p>' in str(each_pTag) and '<p class="slideTitle-False"><i>' in str(each_pTag)) or ('</u></p>' in str(each_pTag) and '<p class="slideTitle-False"><u>' in str(each_pTag)) or '</u></i></p>' in str(each_pTag) or '</strong></u></p>' in str(each_pTag) or '</u></strong></p>' in str(each_pTag):
            each_pTag['class'] = 'slideTitle-True'
            hr_tag = soup.new_tag("hr", attrs={'class': 'paraColor'})
            each_pTag.insert_before(hr_tag)
    
        elif ': </strong></p>' in each_pTag or ':</strong></p>' in each_pTag:
            each_pTag['class'] = 'slideTitle-True'
            each_pTag.replace_with(each_pTag)
    
        else:
            if each_pTag.text.endswith(':') == True or each_pTag.text.endswith(': ') == True:
                each_pTag['class'] = 'slideTitle-True'
                each_pTag.replace_with(each_pTag)
            else:
                if '<p class="slideTitle-True">' in str(each_pTag):
                    each_pTag['class'] = 'slideTitle-True'
                else:
                    each_pTag['class'] = 'slideTitle-False'
                    each_pTag.replace_with(each_pTag)
        
        each_pTag.replace_with(str(each_pTag).replace('<p class="slideTitle-False"></p>',''))

    strhtmNew1 = su.unescape(str(soup11))
    return strhtmNew1
    

def htmlTojson(outputfilename):
    f = codecs.open(outputfilename, 'rb', encoding='utf-8', errors='ignore')
    html_doc = f.read()
    soup = BeautifulSoup(html_doc, 'html.parser')
    presentation = []
    dic_data = {}
    dic_data1 = {}
    testLst = []
    for i in soup.findAll(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'hr','td']):
        testLst.append(i)
    # print(testLst)
    
    titlelst = []
    sentlst = []
    otherlst = []
    
    for i in range(len(testLst)):
        # print(testLst[i])
        if testLst[i]['class'] == ['slideTitle-True']:
            titlelst.append(i)
        elif testLst[i]['class'] == ['slideTitle-False']:
            sentlst.append(i)
        else:
            otherlst.append(i)
            
    print(titlelst, sentlst, otherlst)

    if sentlst[-1] > otherlst[-1]:
        otherTrueCase = sentlst[-1] + 1
    else:
        otherTrueCase = otherlst[-1] + 1
    
    titlelst.append(otherTrueCase)
    
    # print(titlelst,sentlst,otherlst)
    pairLst = list(zip(titlelst[:-1], titlelst[1:]))
    lst = []
    lst1 = []
    j = 0
    k = 0
    res = []
    newLst1 = []
    for q in sentlst:
        if q < titlelst[0]:
            dic_data1 = {}
            dic_data1['slideTitle'] = ''
            dic_data1['sentences'] = testLst[q].text
            newLst1.append(dic_data1)
        else:
            pass
        presentation.append(newLst1)
        dic_data1 = {}
        newLst1 = []
    
    for m in range(len(titlelst) - 1):
        for i in range(titlelst[m] + 1, titlelst[m + 1]):
            lst.append(testLst[i].text)
            result = [list(g) for k, g in groupby(lst, key=bool) if k]
        # print(result)
        newLst = []
        if result == []:
            dic_data = {}
            dic_data['slideTitle'] = testLst[titlelst[m]].text
            dic_data['sentences'] = []
            newLst.append(dic_data)
        else:
            for kk in result:
                kk = [x.strip() for x in kk if x.strip()]
                dic_data = {}
                dic_data['slideTitle'] = testLst[titlelst[m]].text
                dic_data['sentences'] = (kk)
                newLst.append(dic_data)
        lst = []
        presentation.append(newLst)
        dic_data = {}
    flat_list = [item for sublist in presentation for item in sublist]
    return flat_list


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        output_filename = 'index.html'
        htmlGenerated = request.form.get('editordata')
        filenameEntered = request.form.get('filenamedata')

        result = docxTohtmlwithClasses(htmlGenerated,output_filename)
        with open(output_filename, "w", encoding='utf-8') as f1:
            f1.write(result)
        response = htmlTojson(output_filename)
        with open(filenameEntered, "w", encoding='utf-8') as outfile:
            json.dump(response, outfile, ensure_ascii=False)
        return "Your JSON is generated and saved as "+str(filenameEntered)+". You can click on back button, to go back to text area space!!"
    return render_template('index.html')
    

if __name__ == '__main__':
    app.run(debug=True, port=8003)

