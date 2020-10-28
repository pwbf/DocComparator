import docx2txt
import win32com
from win32com.client import Dispatch
from os import mkdir,rename,listdir,path,remove,system
from shutil import rmtree
from re import sub, search
from hashlib import sha3_256

BASE_PATH = './'
BASE_DIR = '_DOCSROOT'
DOCSROOT = BASE_PATH + BASE_DIR

FILE_HASHTABLE = {}
FILE_CHECKEDTABLE = []
FILE_DUPTABLE = []

IMAGE_HASHTABLE = {}
IMAGE_CHECKEDTABLE = []
IMAGE_DUPTABLE = []

def filehasher(fname):
    return ((sha3_256(open(fname,'rb').read()).hexdigest()).upper())

def moveFilesOut(FL):
    for Dirname in FL:
        subPath = DOCSROOT + "/" + Dirname
        subFolder = listdir(subPath)

        sp0 = Dirname.split(" ")
        UNAME = sp0[0]

        sp1 = (sp0[1]).split("_")
        NewFname = sp1[0]

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

def hashImages(DIRNAME):
    dirpath = DOCSROOT + "/" + DIRNAME
    print(DIRNAME)
    subFLIST = listdir(dirpath)
    for FNAME in subFLIST:
        if bool(search(r"\.(png)", FNAME)):
            print(FNAME)
            fullpath = dirpath + "/" + FNAME
            hashed = filehasher(fullpath)
            print("IMAGE SHA3_256=> " + FNAME + " >> " + hashed)
        
            if hashed not in IMAGE_HASHTABLE:
                IMAGE_HASHTABLE.update({hashed : [DIRNAME, FNAME]})
                IMAGE_CHECKEDTABLE.append([FNAME, hashed])
            else:
                dupDIRNAME = IMAGE_HASHTABLE[hashed][0]
                dupFNAME = IMAGE_HASHTABLE[hashed][1]
                IMAGE_DUPTABLE.append([[dupDIRNAME, dupFNAME, DIRNAME, FNAME, hashed]])

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
                else:
                    print("FILE:" + DIRNAME)
                    continue
        print("")

        print("Checked File:")
        for f in FILE_CHECKEDTABLE:
            print(f[1] + " >> " + f[0])
        print("")
        print("Checked Image:")
        for g in IMAGE_CHECKEDTABLE:
            print(g[1] + " >> " + g[0])
        print("")

        print("====================================================================")
        print("Duplicated File:")
        for f in FILE_DUPTABLE:
            print(">> [" + f[0][0] + "]  and  [" + f[0][1] + "] << are identical")
        print("")
        print("====================================================================")
        print("Duplicated Image:")
        for g in IMAGE_DUPTABLE:
            print(">> [" + g[0][0] + ">" + g[0][1] + " ] and [ " + g[0][2] + ">" + g[0][3] + " ] are identical")
        print("")
system("pause")