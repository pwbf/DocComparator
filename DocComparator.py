import docx2txt
import win32com
from win32com.client import Dispatch
from os import mkdir,rename,listdir,path,remove,system
from shutil import rmtree
from re import sub, search
from hashlib import sha3_256
import imagehash
from PIL import Image

BASE_PATH = './'
BASE_DIR = '_DOCSROOT'
DOCSROOT = BASE_PATH + BASE_DIR
THREADNUM = 8
SIMCHK = 20

FILE_HASHTABLE = {}
FILE_CHECKEDTABLE = []
FILE_DUPTABLE = []
CP_List = {}

IMAGE_HASHTABLE = {}
IMAGE_CHECKEDTABLE = []
IMAGE_DUPTABLE = []
IMAGE_TABLE = []
HASHCHK = [[0, []], [0, []], [0, []], [0, []], [0, []]]        #0~10; 11~15; 16~20; 21~25; 25~ [Count, [Name List]]
RF_List = {}

def filehasher(fname):
    return ((sha3_256(open(fname,'rb').read()).hexdigest()).upper())

def write2Log(logstr):
    with open("_DocComparator.log.txt", "a", encoding="utf-8") as f:
        f.write(logstr)
        f.close()
    

def moveFilesOut(FL):
    for Dirname in FL:
        subPath = DOCSROOT + "/" + Dirname
        subFolder = listdir(subPath)

        sp0 = Dirname.split(" ")
        UNAME = sp0[0]

        sp1 = (sp0[1]).split("_")
        NewFname = str(UNAME) + '-' + str(sp1[0])

        for f in subFolder:
            originPath = subPath + "/" + f
            if path.isfile(originPath):
                fext = path.splitext(originPath)[-1]
                tailNum = 1
                newPath = DOCSROOT + "/" + NewFname + fext

                while(True):
                    if(not path.isfile(newPath)):
                        rename(originPath, newPath)
                        print(newPath)
                        break
                    else:
                        newPath = DOCSROOT + "/" + NewFname + "_" + str( tailNum ) + fext
                        tailNum += 1

        rmtree(path.abspath(subPath))


def hashFiles(FL):
    for FNAME in FL:
        fullpath = DOCSROOT + "/" + FNAME
        hashed = filehasher(fullpath)
        print("FILE SHA3_256=> " + FNAME + " >> " + hashed)
        
        if hashed not in FILE_HASHTABLE:
            FILE_HASHTABLE.update({hashed : FNAME})
            FILE_CHECKEDTABLE.append([FNAME, hashed])
        else:
            dupFNAME = FILE_HASHTABLE[hashed]
            FILE_DUPTABLE.append([[FNAME, dupFNAME, hashed]])
            uid0 = FNAME.split('.')[0]
            uid1 = dupFNAME.split('.')[0]
            CP_List.update({uid0 : hashed})
            CP_List.update({uid1 : hashed})

def hashImages(DIRNAME):
    dirpath = DOCSROOT + "/" + DIRNAME
    subFLIST = listdir(dirpath)
    for index, FNAME in enumerate(subFLIST):
        if bool(search(r"\.(png)", FNAME)):
            fullpath = dirpath + "/" + FNAME
            hashed = filehasher(fullpath)
            hashVal = imagehash.average_hash(Image.open(fullpath)) 
            IMAGE_TABLE.append([dirpath, FNAME, hashed, hashVal, int(index % THREADNUM)])
            print("IMAGE SHA3_256=> " + FNAME + " >> " + hashed + " | value>> " + str(hashVal))
        
            if hashed not in IMAGE_HASHTABLE:
                IMAGE_HASHTABLE.update({hashed : [DIRNAME, FNAME]})
                IMAGE_CHECKEDTABLE.append([DIRNAME, FNAME, hashed])
            else:
                dupDIRNAME = IMAGE_HASHTABLE[hashed][0]
                dupFNAME = IMAGE_HASHTABLE[hashed][1]
                IMAGE_DUPTABLE.append([[dupDIRNAME, dupFNAME, DIRNAME, FNAME, hashed]])
                RF_List.update({dupDIRNAME : hashed})
                RF_List.update({DIRNAME : hashed})

def crossCheck(threadnum):
    print('Similarity Check: ')
    for f0 in IMAGE_TABLE:
        h0 = f0[3]
        for f1 in IMAGE_TABLE:
            print('Origin: ' + f0[0] + '>'+f0[1] + '> | Compare: ' + f1[0] + '>'+f1[1], end='')
            if(not ((f0[0] == f1[0]) and (f0[1] == f1[1]))):
                h1 = f1[3]
                chk = abs(h1 - h0)
                #if(chk < SIMCHK):
                #print('Origin: ' + f0[0] + '>'+f0[1] + '> | Compare: ' + f1[0] + '>'+f1[1] + ' | hashvalDiff: '+str(chk))
                print(' | hashvalDiff: '+str(chk))
                if(chk >= 0 and chk < 10):
                    HASHCHK[0][0] += 1
                    HASHCHK[0][1].append([f0[0], f0[1], f1[0], f1[1], chk])
                elif(chk >= 10 and chk < 15):
                    HASHCHK[1][0] += 1
                    HASHCHK[1][1].append([f0[0], f0[1], f1[0], f1[1], chk])
                elif(chk >= 15 and chk < 20):
                    HASHCHK[2][0] += 1
                    HASHCHK[2][1].append([f0[0], f0[1], f1[0], f1[1], chk])
                elif(chk >= 21 and chk < 25):
                    HASHCHK[3][0] += 1
                    HASHCHK[3][1].append([f0[0], f0[1], f1[0], f1[1], chk])
                elif(chk >= 25):
                    HASHCHK[4][0] += 1
                    HASHCHK[4][1].append([f0[0], f0[1], f1[0], f1[1], chk])
    print('')

def mkdirForFiles(FL):
    for FNAME in FL:
        originPath = DOCSROOT + "/" + FNAME

        if bool(search(r"\.(docx)", FNAME)):
            SubDirName = sub(r"\.(docx)","",FNAME)
            newPath = DOCSROOT + "/" + SubDirName
            mkdir(newPath)
            rename(originPath, newPath + "/" +  FNAME)
            docx2txt.process(newPath + "/" +  FNAME, newPath) 

        elif bool(search(r"\.(doc)", FNAME)):
            print("Convert=> " + FNAME + " to DOCX")
            doc2docx(path.abspath(originPath))
            FNAME += "x"
            originPath += "x"
            SubDirName = sub(r"\.(docx)","", FNAME)
            newPath = DOCSROOT + "/" + SubDirName
            mkdir(newPath)
            rename(originPath, newPath + "/" +  FNAME)
            docx2txt.process(newPath + "/" +  FNAME, newPath) 
        else:
            continue
        print("mkdir=> " + FNAME + " >> " + SubDirName)

def doc2docx(p):
    word = win32com.client.Dispatch('word.application')
    word.DisplayAlerts = 0
    doc = word.Documents.Open(p)
    doc.SaveAs(p+"x", 12)
    doc.Close()
    word.Quit()
    remove(p)


if not path.isdir(BASE_PATH + BASE_DIR):
    print("Directory not exist!")
    print("Create for you")
    print("Put documents inside and restart program")
    mkdir(BASE_DIR)
else:
    FILELST = listdir(DOCSROOT)
    if not FILELST:
        print("Directory is empty")
    else:
        if((input("Enter \"T\" if you want to move file out of folder automatically:")).upper() == "T"):
            print(">> Moving file out >>")
            moveFilesOut(FILELST)
            print("")

        FILELST = listdir(DOCSROOT)
        print(">> Checking file >>")
        hashFiles(FILELST)
        print("")

        print(">> Moving file >>")
        mkdirForFiles(FILELST)
        print("")
        print(">> Checking each image >>")
        DIRLST = listdir(DOCSROOT)
        if not FILELST:
            print("Directory is empty")
        else:
            for DIRNAME in DIRLST:
                if path.isdir(DOCSROOT + "/" + DIRNAME):
                    print("DIR:" + DIRNAME)
                    hashImages(DIRNAME)
                    print("")
                else:
                    print("FILE:" + DIRNAME)
                    continue
        print("")
        print(">> Checking each image similarity>>")
        crossCheck(0)
        print("")

        print("Checked File:")
        write2Log("Checked File:\n")
        for f in FILE_CHECKEDTABLE:
            print(f[0] + " >> " + f[1])
            write2Log(f[0] + " >> " + f[1] + "\n")
        print("")
        write2Log("\n")
        print("Checked Image:")
        write2Log("Checked Image:\n")
        for g in IMAGE_CHECKEDTABLE:
            print(g[0] + "\\ "+ g[1] + " >> " + g[2])
            write2Log(g[0] + "\\ "+ g[1] + " >> " + g[2] + "\n")
        print("")
        write2Log("\n")

        print("====================================================================")
        write2Log("====================================================================\n")
        print("Duplicated File:")
        write2Log("Duplicated File:\n")
        for f in FILE_DUPTABLE:
            print(">> [" + f[0][0] + "]  and  [" + f[0][1] + "] << are identical")
            write2Log(">> [" + f[0][0] + "]  and  [" + f[0][1] + "] << are identical" + "\n")
        print("")
        write2Log("\n")
        print("====================================================================")
        write2Log("====================================================================\n")
        print("Duplicated Image:")
        write2Log("Duplicated Image:\n")
        for g in IMAGE_DUPTABLE:
            print(">> [" + g[0][0] + ">" + g[0][1] + " ] and [ " + g[0][2] + ">" + g[0][3] + " ] are identical")
            write2Log(">> [" + g[0][0] + ">" + g[0][1] + " ] and [ " + g[0][2] + ">" + g[0][3] + " ] are identical" + "\n")
        print("")
        write2Log("\n")
        print("====================================================================")
        write2Log("====================================================================\n")
        print("Copastier:")
        write2Log("Copastier:\n")
        for c in CP_List:
            print(c)
            write2Log(c + "\n")
        print("")
        write2Log("\n")
        print("====================================================================")
        write2Log("====================================================================\n")
        print("Referencier:")
        write2Log("Referencier:\n")
        for r in RF_List:
            print(r)
            write2Log(r + "\n")
        print("")
        write2Log("\n")
        print("====================================================================")
        write2Log("====================================================================\n")
        print('Hash Diff spec:')
        print('0~10: '+str(HASHCHK[0][0]))
        for r in HASHCHK[0][1]:
            print(r)

        print('\n11~15: '+str(HASHCHK[1][0]))
        for r in HASHCHK[1][1]:
            print(r)

        print('\n16~20: '+str(HASHCHK[2][0]))
        for r in HASHCHK[2][1]:
            print(r)

        print('\n21~25: '+str(HASHCHK[3][0]))
        print('\n25~: '+str(HASHCHK[4][0]))
        print("")
system("pause")
