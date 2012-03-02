'''
Created on Mar 2, 2012

@author: kernel-sanders
'''
import OleFileIO_PL
import string
import glob
import os

### Read the inFile and check if it is an OLE file, return this as an ole object
def readFile(inFile):
    if OleFileIO_PL.isOleFile(inFile):
        print 'File is an OLE'
        ole = OleFileIO_PL.OleFileIO(inFile)
    else:
        print 'File is NOT an OLE'
    return ole

### Check to see if "data" is within our defined set of uppercase, lowercase, numbers, ':', '.', and '/'
def isChar(data):
    allowed = set(string.ascii_lowercase + string.ascii_uppercase + ":./0123456789")
    if set(data) <= allowed:
        return True
    else:
        return False

### Print the ACII characters of a recovery *.dat file
def printFile(ole):
    for stream in ole.listdir():
        print stream
        streamData = ole.openstream(stream)
        binaryData = streamData.read()
        url = ''
        for possibleChar in binaryData:
            if isChar(possibleChar):
                url += possibleChar
        print url

### Locate all of the Internet Explorer recovery files on the hard drive where this script is run, return them as a list      
def findFiles():
    pathsToFiles = []
    locations = glob.glob(os.getcwd()[0:3] + 'Users\*\AppData\Local\Microsoft\Internet Explorer\Recovery')
    for location in locations:
        for folders in os.listdir(location):
            foldersPath = os.path.join(location,folders)
            for filesOrFolders in os.listdir(foldersPath):
                if filesOrFolders[-4:] == '.dat':
                    pathsToFiles.append(os.path.join(foldersPath,filesOrFolders))
                else:
                    for files in os.listdir(os.path.join(foldersPath,filesOrFolders)):
                        if files[-4:] == '.dat':
                            pathsToFiles.append(os.path.join(os.path.join(foldersPath,filesOrFolders),files))
                        else:
                            print 'An error occurred'
    return pathsToFiles
        

if __name__ == '__main__':
    filelist = findFiles()
    for files in filelist:
        print 'Trying to open: ' + files
        currentfile = readFile(files)
        printFile(currentfile)