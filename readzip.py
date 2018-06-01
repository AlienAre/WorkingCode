import zipfile

zfile = zipfile.ZipFile('C:\Users\Wangwe5\Downloads\C529.zip')
for finfo in zfile.infolist():
    ifile = zfile.open(finfo)
#        line_list = ifile.readlines()
#        print line_list

#print FileList
for txts in FileList:
	with open(txts) as f:
		for line in f:
			if line.strip():
				#print 'in strip'
				if TypePa.match(line[2:].lstrip(' ')):
					#print 'Process'
					ProcesTxt(txts)
					break
