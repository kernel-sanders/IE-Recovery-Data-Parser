'''
Created on Mar 2, 2012

@author: kernel-sanders
'''
import OleFileIO_PL
import string
import glob
import os
from optparse import OptionParser
from xml.dom.minidom import Document

### Read the inFile and check if it is an OLE file, return this as an ole object
def readFile(inFile):
    if OleFileIO_PL.isOleFile(inFile):
        ole = OleFileIO_PL.OleFileIO(inFile)
    return ole

### Check to see if "data" is within our defined set of printable ascii characters
def isChar(data):
    allowed = set(string.letters + string.punctuation)
    return set(data) <= allowed

### Parse the ACII characters of a recovery *.dat file and add them to the xmlElement as a "url" element's text
def writeXML(ole, xmlElement):
    for stream in ole.listdir():
        if str(stream).strip("[]'")[0:1] != "T" or str(stream).strip("[]'") == 'FrameList' or str(stream).strip("[]'") == 'ClosedTabList': # Only worry about the streams with useful information in them
            #print str(stream).strip("[]'")
            streamData = ole.openstream(stream)
            binaryData = streamData.read()
            printableData = ''
            for possibleChar in binaryData:
                if isChar(possibleChar):
                    printableData += possibleChar
            urlList = printableData.split("http")
            for urls in urlList:
                if urls != '' and (urls[0:1] == ':' or urls[0:1] == 's'):
                    urlElement = doc.createElement("url")
                    xmlElement.appendChild(urlElement)
                    urlText = doc.createTextNode('http' + urls)
                    urlElement.appendChild(urlText)

### Locate all of the Internet Explorer recovery files on the hard drive where this script is run, return them as a list      
def findFiles(searchDir):
    pathsToFiles = []
    if searchDir == None:
        locations = glob.glob(os.getcwd()[0:3] + 'Users\*\AppData\Local\Microsoft\Internet Explorer\Recovery')
        for location in locations:
            pathsReturned = walkDir(location)
            for paths in pathsReturned:
                pathsToFiles.append(paths)
    else:
        pathsToFiles = walkDir(searchDir)
    return pathsToFiles

# Walk the searchDir looking for *.dat files, return a list of paths to the files 
def walkDir(searchDir):
    paths = []
    for root, dirs, files in os.walk(searchDir):
        for currentFile in files:
            if currentFile[-4:] == '.dat':
                paths.append(os.path.join(root,currentFile))
    return paths
        

if __name__ == '__main__':
    parser = OptionParser()
    parser.add_option("-d", "--dir", dest="directory", help='User provided directory to search. If none is provided, the hard drive of the current working directory will be searched, and a "Username" attribute will be added to each "Recovery_File" XML element.')
    parser.add_option("-f", "--filename", dest="filename", help='User provided filename for output (none defaults to "IE_History.xml").')
    (options, args) = parser.parse_args()
    searchDir = options.directory
    filelist = findFiles(searchDir)
    if searchDir == None:
        includeUsername = True
    else:
        includeUsername = False
    filename = options.filename
    if filename == None:
        filename = 'IE_History.xml'
    doc = Document()
    main = doc.createElement("IE_Recovery_Data_Files")
    doc.appendChild(main)
    for files in filelist:
        fileName = files.split('\\')[-1:]
        xmlElement = doc.createElement("Recovery_File")
        xmlElement.setAttribute("id", str(fileName).strip("[]'"))
        if includeUsername:
            userName = str(files.split('\\')[2:3]).strip("[]'")
            xmlElement.setAttribute("Username", userName)
        main.appendChild(xmlElement)
        currentfile = readFile(files)
        writeXML(currentfile, xmlElement)
    f = open(filename, 'w')
    f.write(doc.toprettyxml())
    f.close()