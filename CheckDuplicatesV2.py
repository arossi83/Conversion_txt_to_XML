import re
import datetime

def CheckDuplicates(par,files,logf):
    seen=[]
    duplicates=[]
    for x in par:
        single=[]
        start=-1
        if x not in seen:
            while True:
                try:
                    start=par.index(x,start+1)
                    single.append(start)
                except:
                    if len(single)>1:
                        duplicates.append(single)
                    break
        seen.append(x)
    plainD=[]
    for kk in duplicates:
        for k in kk:
            plainD.append(k)
    #print(par)
    #print(duplicates)
    #print(plainD)

    cleanPar=[]
    cleanFile=[]
    for ix,x in enumerate(par):
        if ix not in plainD:
            cleanPar.append(x)
            cleanFile.append(files[ix])
    #logf.write(cleanPar)
    #logf.write(cleanFile)
    idToAdd=[]
    for d in duplicates:
        uniq=[]
        for idx in d:            
            fname=files[idx]
            Tstring=re.findall('[0-9]+_[0-9]+_202[0-9]_[0-9]+h[0-9]+m[0-9]+s',fname)[0]
            date = datetime.datetime.strptime(Tstring,"%d_%m_%Y_%Hh%Mm%Ss")
            timestamp = datetime.datetime.timestamp(date)
            #logf.write("%s - %d" % (fname,timestamp))
            uniq.append(timestamp)
        idToAdd.append(d[uniq.index(max(uniq))])

    #logf.write(idToAdd)
    for idx in idToAdd:
        cleanPar.append(par[idx])
        cleanFile.append(files[idx])
    #logf.write(cleanPar)
    #logf.write(cleanFile)
    logf.write("--> %d duplicates have been removed\n" % len(duplicates))
    if len(duplicates) > 0:
        rem = []
        for element in files:
            if element not in cleanFile:
                rem.append(element)
        logf.write("--> Removed Elements:\n")
        logf.write(rem)
        logf.write("\n")
    return cleanPar, cleanFile
