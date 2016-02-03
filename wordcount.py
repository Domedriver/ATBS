import docx, sys, re, unidecode

fileName = ' '.join(sys.argv[1:])
doc = docx.Document(fileName)

wordRegex = re.compile(r"[a-z']*") # should capture words and contractions

wordCount = {}
fullText = []
cleanText = []
maxWordLength = 0

for para in doc.paragraphs:
	if para.text != '':
		fullText.append(para.text)
for para in fullText:
	sentenceWords = para.split()
	for word in sentenceWords:		
		wordConvert = unidecode.unidecode(word)
		mo = wordRegex.search(wordConvert.lower())
		if mo is not None and mo.group() is not '':
			cleanText.append(mo.group().lower())
			if len(mo.group()) > maxWordLength:
				maxWordLength = len(mo.group())
	
for word in cleanText:
	wordCount.setdefault(word, 0)
	wordCount[word] += 1

wordSort = [(k,v) for v,k in sorted([(k, v) for v, k in wordCount.items() if k > 2], reverse=True)]

for thing in wordSort:
	word, number = thing
	print word.ljust(maxWordLength), ': ', str(number).rjust(2)
