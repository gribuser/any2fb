#!python
# -*- coding: cp866 -*-
"""\
Usage: any2fb2.py [options] [@]<file-or-mask> [file-or-mask...]

Options available:
  -o --output=<file|dir>       Set output file name or dir
  -m --rename=<mask>           Set output file name to match the given mask.
                               You can use the following sequences in the mask:
                                   %f - author's first name
                                   %l - author's last name
                                   %t - the book title
                               You must *NOT* supply the .fb2 extension here!
  -g --genre=<name>            Set genre tag for all converted books
  -f --preserve-form           Preserv <form> content
  -c --no-charset              Do not convert charset
  -e --no-epigraphs            Do not detect epigraphs
  -l --no-empty-lines          Do not create empty lines
  -d --no-description          Do not create description
  -r --max-fix-count=<n>       Set maximum fix count to n
  -u --no-quotes               Do not convert "quotes" to <<quotes>>
  -n --no-footnotes            Do not convert [text] and {text} into footnotes
  -i --no-italic               Do not detect _italic_ text
  -j --no-broken-para          Do not restore broken paragraphs
  -p --no-poems                Do not search poems
  -h --only-h-tags             Only use existing headers (<h1-6>)
  -s --ignore-indents          Ignore line indents (leading spaces)
     --no-dash                 Do not convert short -
  -t --type=<indented|lines|auto> Set text type to "indented" or "lines" or "auto" (default)
     --images=<none|local|dynamic> Remove all (none) or offsite (offsite) images from the document
                               Or leave even dynamic images (dynamic).
     --no-links                Delete all external links
     --follow=<n>              Set links folowing deepness (0-9)
     --follow-offsite          Follow off-site links
     --overwrite               Overwrite output file if it's already exists
     --out-encoding=<val>      Set encoding of output file (default is 'windows-1251')
     --log-encoding=<val>      Set encoding of screen log (stdout) (default is 'cp866')
     --log=<fileName>          Set log file name (default none)
     --fail-on-space           Fail converion if space char found in first or last name.
                               This will help with books with 2 or more authors.
     --copy-failed=<dir>       Copy failed source files to this directory.
     --recurce=<mask>          Recurce into subdirs and process files matching given mask.
     --exclude=<mask>          Exclude mask for file recursion (does not affect explicit file names).
  -q --quiet                   No info messages

  You can use @filename as input argument to make this script read options or file names 
  or masks from given file. The file must contain single file name or mask per line.

  You may also edit Regular Expressions in this script to get more control over import
"""
#------------------------------------------------------------------------
#  Copyright (C):
#
#  Any2FB2 ActiveX control and FBE plugin by GribUser.
#  Original any2fb2.vbs script by GribUser.
#  This any2fb2.py script by Alex Shabarshoff <mailto:shura@uc.ru>.
#
#  This script (any2fb2.py) is public domain.
#------------------------------------------------------------------------
#  $Id$
#
#  $Log$
#------------------------------------------------------------------------

import sys
import win32com.client
#from win32com.client import gencache
import os
import getopt
import dircache,fnmatch
import time
__all__ = ["Any2Fb2Error",
           "any2fb2", 
           "GetFB2Info", 
           "GetFB2InfoStr", 
           "GetFB2InfoDom", 
           "GetFB2HeaderDom", 
           "EscapeBookFileName"]

class Any2Fb2Error(Exception):
    opt = ''
    msg = ''
    def __init__(self, msg, opt):
        self.msg = msg
        self.opt = opt
        Exception.__init__(self, msg, opt)

    def __str__(self):
        return self.msg

def EscapeBookFileName( name ):
    name = name.replace('&quot;','').replace('&amp;','&').replace('&lt;','&').replace('&gt;','&').replace('"','').replace(':','').replace('|','').replace('\\',' ').replace('/',' ').replace('<','').replace('>','').replace('&',' and ').replace('?','')
    while name.find("..") >= 0:
            name = name.replace("..",".")
    while name.find("  ") >= 0:
            name = name.replace("  "," ")
    return name

class any2fb2_log:
	def __init__(self,fileName,log_encoding = "cp866", mute=0):
	        if log_encoding is None or log_encoding == "":
	        	self.log_encoding = "cp866"
	        else:
	        	self.log_encoding = log_encoding
	        	
		if not (fileName is None or fileName == ""):
			self.f = open(fileName,"a+t")
	        else:
	        	self.f = None
	        self.mute = mute
	        self.was_n = 1

	def __del__(self):
		if not (self.f is None):
			self.f.close()

	def write(self,s):
		if not self.f is None:
			if self.was_n:
				self.f.write( (time.asctime(time.localtime())+": ").encode(self.log_encoding) )
			self.f.write( s.encode(self.log_encoding) )
		if not self.mute:
			print s.encode(self.log_encoding),
	        self.was_n = s[-1] == '\n'

def _add_matched_file( fileName, newFiles ):
	newFiles.append( fileName )

def any2fb2(program_args,isInteractive = 0):

        long_opts=[
        "output=",
        "rename=",
        "genre=",
        "log=",
        "overwrite",
        "preserve-form",
        "no-charset",
        "no-epigraphs",
        "no-empty-lines",
        "no-description",
        "max-fix-count=",
        "no-quotes",
        "no-footnotes",
        "no-italic",
        "no-broken-para",
        "no-poems",
        "only-h-tags",
        "ignore-indents",
        "no-dash",
        "type=",
        "images=",
        "no-links",
        "follow=",
        "follow-offsite",
        "out-encoding=",
        "log-encoding=",
        "fail-on-space",
        "copy-failed=",
        "recurce=",
        "exclude=",
        "quiet"
        ]

        FBApp = win32com.client.Dispatch("any_2_fb2.any2fb2")

        mute         = 0
        true         = 1 == 1
        out_encoding = "windows-1251"
        log_encoding = "cp866"

        newFiles = []
        for f in program_args:
        	if f[0] == '@':
        	        try:
        			fl = open(f[1:],"rt")
        			lines = fl.readlines()
        			fl.close()
                		for l in lines:
                			if l == "" or l == "\n" or l[0] == '#':
                				continue
                		        newFiles.append(l[:-1])
        		except OSError:
        			newFiles.append(f)
        		continue
        	newFiles.append(f)
        program_args = newFiles

        try:
                option, files = getopt.getopt(program_args, "o:m:fceldr:unijphst:g:q", long_opts)
        except getopt.GetoptError,e:
        	if isInteractive:
                        print __doc__
                        print
                        print "Error:",e
                raise Any2Fb2Error(e.msg, e.opt)


        outdir        = ""
        fileNameMask  = ""
        fileGenre     = ""
        overwrite     = 0
        log           = ""
        fail_on_space = 0
        copyFailed    = None
        recurce       = []
        exclude       = []

        for arg,val in option:
                if arg   == "-f" or arg == "--preserve-form":   FBApp.PreserveForm       = true
                elif arg == "-m" or arg == "--rename":          fileNameMask             = unicode(val,log_encoding)
                elif arg == "-g" or arg == "--genre":           fileGenre                = unicode(val,log_encoding)
                elif arg == "-o" or arg == "--output":          outdir                   = unicode(val,log_encoding)
                elif arg == "--overwrite":                      overwrite                = 1
                elif arg == "-c" or arg == "--no-charset":      FBApp.noConvertCharset   = true
                elif arg == "-e" or arg == "--no-epigraphs":    FBApp.noEpigraphs        = true
                elif arg == "-l" or arg == "--no-empty-lines":  FBApp.noEmptyLines       = true
                elif arg == "-d" or arg == "--no-description":  FBApp.noDescription      = true
                elif arg == "-r" or arg == "--max-fix-count":   FBApp.FixCount           = int(val)
                elif arg == "-u" or arg == "--no-quotes":       FBApp.noQuotesConvertion = true
                elif arg == "-n" or arg == "--no-footnotes":    FBApp.noFootNotes        = true
                elif arg == "-i" or arg == "--no-italic":       FBApp.noItalic           = true
                elif arg == "-j" or arg == "--no-broken-para":  FBApp.noRestoreBrokenParagraphs = true
                elif arg == "-p" or arg == "--no-poems":        FBApp.noPoems            = true
                elif arg == "-h" or arg == "--only-h-tags":     FBApp.noHeaders          = true
                elif arg == "-s" or arg == "--ignore-indents":  FBApp.ignoreLineIndent   = true
                elif arg == "--no-defice":                      FBApp.noLongDashes       = true
                elif arg == "-t" or arg == "--type":            
                        if val == "indented":
                                                                FBApp.TextType           = 1
                        elif val == "lines":
                                                                FBApp.TextType           = 2
                        else:
                                                                FBApp.TextType           = 0
                elif arg == "--images":
                        if val == "none":
                                                                FBApp.noImages           = true
                        elif val == "local":
                                                                FBApp.noOffSiteImages    = true
                        elif val == "dynamic":
                                                                FBApp.leaveDinamicImages = true
                elif arg == "--no-links":                       FBApp.noExternalLinks    = true
                elif arg == "--follow":                         FBApp.FollowLinksDeep    = int(val)
                elif arg == "--follow-offsite":                 FBApp.FollowOffSiteLinks = true
                elif arg == "--out-encoding":                   out_encoding             = val
                elif arg == "--log-encoding":                   log_encoding             = val
                elif arg == "--fail-on-space":                  fail_on_space            = 1
                elif arg == "--copy-failed":                    import shutil; copyFailed = val
                elif arg == "-q" or arg == "--quiet":           mute                     = 1
                elif arg == "--log":                            log                      = val
                elif arg == "--recurce":                        recurce.append(val)
                elif arg == "--exclude":                        exclude.append(val)

        if log != "":
        	log = any2fb2_log( log, log_encoding, mute )
        else:
        	log = any2fb2_log( "",  log_encoding, mute )
        log.mute = 1
        log.write( " --------------- Any2FB2 started ---------------\n")
        log.mute = mute
#                if mute == 0:
#                        print "Using option", arg, val


        #Edit folowing lines to use regular expressions

        #FBApp.reOnlyFollowLinks='\.html'
        #FBApp.reNeverFollowLinks='adv|ban'
        #FBApp.reHeadersDetect='chapter\s\d+'
        #FBApp.reOnLoad='Frodo\nGoblins'
        #FBApp.reOnDone='<p>\n<p>\s\n@\nG'

        newFiles = []
        for f in files:
        	if not (f.find('*') >= 0 or f.find('?') >= 0 or f.find('[') >= 0):
        		newFiles.append(f)
        		continue
#        	if f[0] == '@':
#        		continue
                d,n = os.path.split(f)
                if n == "":
                        n = d
                        d = "."
                if d == "":
                        d = "."
       		_scan_path( d, 0, n, exclude, _add_matched_file, newFiles )

       	_scan_path( ".", 1, recurce, exclude, _add_matched_file, newFiles )

        if len(newFiles) == 0:
                if isInteractive:
                        print __doc__
                        print
                        if len(files) > 0:
                        	print "Error: No files matches the input file mask."
                        else:
                        	print "Error: No input files given."
                raise Any2Fb2Error("No files matches the input file mask", files)

#        if outdir != "" and (not os.path.isdir(outdir)) and len(newFiles) <= 1:
#                if isInteractive:
#                        print "Error: Directory",outdir,"does not exist, but mulifile operation requested."
#                raise Any2Fb2Error("Directory %s does not exist, but mulifile operation requested" % outdir, outdir)


        for f in newFiles:
            log.write( f )
            DOM = FBApp.Convert(f)

            if DOM is None:
                    log.write(" FAILURE\n")
                    try:
                    	log.write(FBApp.LOG+"\n")
                    except:
                    	pass
                    if not copyFailed is None:
                            d,n = os.path.split(f)
                            try:
                                    shutil.copy( f, copyFailed )
                            except OSError:
                                    pass
                    if len(newFiles) == 1:
                            raise Any2Fb2Error("Error converting file %s" % f, f)
            else:
                if outdir != "":
                        if os.path.isdir(outdir):
                                d,n = os.path.split(f)
                                n,e = os.path.splitext(n)
                                out = os.path.join( outdir, n+".fb2")
                        else:
                                out = outdir
                else:
                        n,e = os.path.splitext(f)
                        out = n + ".fb2"


                s = DOM.xml

                if fileGenre != "":
                        s = s.replace('<genre></genre>', '<genre>'+fileGenre+'</genre>')
                        assert s.find('<genre>'+fileGenre+'</genre>') >= 0
                
                
                if fileNameMask != "":
#                        first_name, last_name, title, genre = GetFB2InfoStr(s, out)
			try:
                        	first_name, last_name, title, genre = _GetFB2InfoMSXML(DOM)
                        	if fail_on_space and (last_name.find(" ") >= 0 or first_name.find(" ") >= 0):
                                	log.write(" FAILURE The space char found in last or first name (probably 2 or more authors)\n")
                                	if not copyFailed is None:
                                        	d,n = os.path.split(f)
                                        	try:
                                			shutil.copy( f, copyFailed )
                                		except OSError:
                                			pass
                                        if len(newFiles) == 1:
                                		raise e
                                	continue
                        except Any2Fb2Error,e:
                        	log.write(" FAILURE "+e.msg+"\n")
                        	if not copyFailed is None:
                                	d,n = os.path.split(f)
                                	try:
                        			shutil.copy( f, copyFailed )
                        		except OSError:
                        			pass
                                if len(newFiles) == 1:
                        		raise e
                        	continue
                        d,n = os.path.split(out)
                        n = fileNameMask.replace("%f", EscapeBookFileName(first_name)).replace("%l", EscapeBookFileName(last_name)).replace("%t", EscapeBookFileName(title)) + ".fb2"
                        while n.find("..") >= 0:
                                n = n.replace("..",".")
                        d,n = os.path.split(os.path.join(d,n))
                        out = os.path.join(d,n)

                s = s.replace('<?xml version="1.0"?>', '<?xml version="1.0" encoding="%s"?>' % out_encoding)

                d,n = os.path.split(out)
                try:
                        os.makedirs(d)
                except OSError:
                        pass

                if not overwrite and os.path.isfile(out):
                        log.write(" FAILURE File "+out+"already exists.\n")
                        if len(newFiles) == 1:
                                raise Any2Fb2Error("File %s already exists." % out, out)
                        continue

                f = open(out,"w+b")
                f.write(s.encode(out_encoding))
                f.close()

                if mute == 0:
                        log.write( " -> "+out+" OK\n")


def GetFB2Info(fileName):
        f = open(fileName,"rt")
        xml_string = f.read()
        f.close()
        return GetFB2InfoStr(xml_string)

def GetFB2HeaderDom(xml_string, fileName=""):
        pos = xml_string.find("<FictionBook")
        if pos < 0:
            raise Any2Fb2Error("Can't find the <FictionBook> tag", fileName)
        pos = xml_string.find("<body")
        if pos < 0:
            raise Any2Fb2Error("Can't find the <body> tag", fileName)

        # reduce book size to title-info only
        # this incredibly decreaces the parsing time!
        info = xml_string[:pos] + "</FictionBook>"

        import xml.dom.minidom

        doc = xml.dom.minidom.parseString( info )
        return doc

def GetFB2InfoStr(xml_string, fileName=""):
        return GetFB2InfoDom(GetFB2HeaderDom(xml_string, fileName))

def GetFB2InfoDom(doc):
        nn = doc.documentElement.getElementsByTagName("title-info")
        if nn is None or len(nn) == 0:
            raise Any2Fb2Error("Can't find the <title-info> tag", fileName)
        nn = nn[0]

        first_name = nn.getElementsByTagName("first-name")[0].firstChild.data
        last_name  = nn.getElementsByTagName("last-name")[0].firstChild.data
        title  = nn.getElementsByTagName("book-title")[0].firstChild.data
        genre  = nn.getElementsByTagName("genre")[0]

        return first_name, last_name, title, genre

def _GetFB2InfoMSXML(doc, fileName = ""):
        title_info = doc.documentElement.childNodes.item(0).childNodes.item(0)
        nn = title_info.childNodes
        first_name = last_name = genre = title = ""
        for i in range(0,nn.length):
        	n = nn.item(i).nodeName
        	if n == 'author':
        	        author = nn.item(i).childNodes
                        for j in range(0,author.length):
		               an = author.item(j).nodeName
                               if an == 'first-name':
                                       first_name = author.item(j).text
                               elif an == 'last-name':
                                       last_name = author.item(j).text
        	elif n == 'book-title':
        		title = nn.item(i).text
        	elif n == 'genre':
        		genre = nn.item(i).text

        if title == "":
            raise Any2Fb2Error("Can't find the <book-title> tag", fileName)
        if first_name == "":
            raise Any2Fb2Error("Can't find the <first-name> tag", fileName)
        if last_name == "":
            raise Any2Fb2Error("Can't find the <last-name> tag", fileName)
        return first_name, last_name, title, genre

def _scan_path( startPath, dive, docMasks, excludeMasks, func, arg ):
	contents = dircache.listdir( startPath )
        selectedFiles = {}
        for m in docMasks:
	       	selected = fnmatch.filter( contents, m )
        	for i in range(len(selected)):
        		selectedFiles[selected[i]] = 1

        for m in excludeMasks:
	       	selected = fnmatch.filter( contents, m )
        	for i in range(len(selected)):
        		selectedFiles[selected[i]] = 0
        selected = []
        for k in selectedFiles.keys():
        	if selectedFiles[k] == 1:
        		selected.append(k)

	for i in range( len(selected) ):
		if ( not os.path.isdir(startPath+'/'+selected[i]) ):
			func( os.path.join(startPath,selected[i]), arg )
	if dive:
		for i in range( len(contents) ):
			if os.path.isdir(startPath+'/'+contents[i]):
				_scan_path( os.path.join(startPath,contents[i]), dive, docMasks, excludeMasks, func, arg )


"""
#"-q",  
"-o",".",
"-f",  
"-c",  
"-e",  
"-l",  
"-d",  
"-u",  
"-n",  
"-i",  
"-j",  
"-p",  
"-h",  
"-s",  

"-r","555",
"-t","lines",
"-g","none",

"--preserve-form",
"--no-charset",
"--no-epigraphs",
"--no-empty-lines",
"--no-description",
"--no-quotes",
"--no-footnotes",
"--no-italic",
"--no-broken-para",
"--no-poems",
"--only-h-tags",
"--ignore-indents",
"--no-dash",
"--no-links",
"--follow-offsite",
#"--quiet",

"--max-fix-count=999",
"--type","indented",
"--images","dynamic",
"--follow","3",
"""

test_args = [
"--output","c:\\eBooks\\ForImport",
"--rename","%f %l\\%f %l. %t",
"--genre","SF",
"--overwrite",
"--out-encoding","windows-1251",
"--log","c:/eBooks/any2fb2.tmp",

"C:\\book\\FOUNDATION\\*.html",
#"@test.tmp"
#"outfile"
]

if __name__ == "__main__":
        try:
	        any2fb2(sys.argv[1:], 1)
#                any2fb2(test_args, 1)
        	sys.exit(0)
        except Any2Fb2Error:
        	sys.exit(1)
