' Original script by GribUser (http://www.gribuser.ru/)
' Modified by Fj (fj.mail@gmail.com) - enchanced command line parsing 
'	& more structured code

Option Explicit

Dim FBApp ' converter
Dim mute ' silent mode
Dim srcfile 'source file name
Dim dstfile 'destination file name

function SetRegexes
	'Edit folowing lines to use regular expressions

	'FBApp.reOnlyFollowLinks='\.html'
	'FBApp.reNeverFollowLinks='adv|ban'
	'FBApp.reHeadersDetect='chapter\s\d+'
	'FBApp.reOnLoad='Frodo'& vbCr &'Goblins'
	'FBApp.reOnDone='<p>'& vbCr &'<p>s'& vbCr &'@'& vbCr &'G'
end function

function OnError (errString)
	WScript.Echo(_
		"Error: " & errString & vbCr & vbCr &_
		"Usage: any2fb2.vbs <infile> [<outfile>] [options]" & vbCr &_
		"Options available:" & vbCr &_
		"-f Preserve <form> content" & vbCr &_
		"-c Do not convert charset" & vbCr &_
		"-e Do not detect epigraphs" & vbCr &_
		"-l Do not create empty lines" & vbCr &_
		"-d Do not create description" & vbCr &_
		"-r Set maximum fix count to 1000" & vbCr &_
		"-q Do not convert ""quotes"" to «quotes»" & vbCr &_
		"-n Do not convert [text] and {text} into footnotes" & vbCr &_
		"-i Do not detect _italic_ text" & vbCr &_
		"-j Do not restore broken paragraphs" & vbCr &_
		"-p Do not search poems" & vbCr &_
		"-h Only use existing headers (<h1-6>)" & vbCr &_
		"-s Ignore line indents (leading spaces)" & vbCr &_
		"-ld Do not convert short - into" & vbCr &_
		"-t1|-t2 Set text type to ""indented""|""with empty lines""" & vbCr &_
		"-g Remove all images from the document" & vbCr &_
		"-go Remove off-site images from the document" & vbCr &_
		"-gd Do not remove dynamic images" & vbCr &_
		"-el Delete all external links" & vbCr &_
		"-fd# Set links folowing deepness (0-9)" & vbCr &_
		"-fo Follow off-site links" & vbCr &_
		"-mute No info messages" & vbCr & vbCr &_
		"You may also edit Regular Expressions in this script to get more control over import")
	WScript.Quit 1
End function

function ParseSwitch(arg)
	Select Case arg
		Case "-f" FBApp.PreserveForm = True
		Case "-c" FBApp.noConvertCharset = True
		Case "-e" FBApp.noEpigraphs = true
		Case "-l" FBApp.noEmptyLines = true
		Case "-d" FBApp.noDescription = true
		Case "-r" FBApp.FixCount = 1000
		Case "-q" FBApp.noQuotesConvertion = true
		Case "-n" FBApp.noFootNotes = true
		Case "-i" FBApp.noItalic = true
		Case "-j" FBApp.noRestoreBrokenParagraphs = true
		Case "-p" FBApp.noPoems = true
		Case "-h" FBApp.noHeaders = true
		Case "-s" FBApp.ignoreLineIndent = true
		Case "-ld" FBApp.noLongDashes = true
		Case "-t1" FBApp.TextType = 1
		Case "-t2" FBApp.TextType = 2
		Case "-g" FBApp.noImages = true
		Case "-go" FBApp.noOffSiteImages = true
		Case "-gd" FBApp.leaveDinamicImages = true
		Case "-el" FBApp.noExternalLinks = true
		Case "-fd1" FBApp.FollowLinksDeep = 1
		Case "-fd2" FBApp.FollowLinksDeep = 2
		Case "-fd3" FBApp.FollowLinksDeep = 3
		Case "-fd4" FBApp.FollowLinksDeep = 4
		Case "-fd5" FBApp.FollowLinksDeep = 5
		Case "-fd6" FBApp.FollowLinksDeep = 6
		Case "-fd7" FBApp.FollowLinksDeep = 7
		Case "-fd8" FBApp.FollowLinksDeep = 8
		Case "-fd9" FBApp.FollowLinksDeep = 9
		Case "-fo" FBApp.FollowOffSiteLinks = true
		Case "-mute" mute=1
		Case Else OnError "unknown switch: " & arg
	End Select
End function

'Create converter

On Error resume next

Set FBApp = CreateObject("any_2_fb2.any2fb2")

If Err Then
  WScript.Echo "Unable to create ActiveX object ""any_2_fb2.any2fb2"". This may be because of the incorrect setup."&_
		vbCr & vbCr & "Please reinstall the program and try rinning this script again."
  WScript.Quit 1
end if

On Error GoTo 0

SetRegexes

'Parse command line
mute = 0
Dim args, arg

Set args = WScript.Arguments

For Each arg in args
	If Mid(arg, 1, 1) = "-" Then
		ParseSwitch(arg)
	ElseIf srcfile = Empty Then
		srcfile = arg
	ElseIf dstfile = Empty Then
		dstfile = arg
	Else
		OnError "Too many file names specified: "_
			& srcfile & ", " & dstfile & ", " & arg
	End If
Next

If srcfile = Empty Then
	OnError "No source file name specified"
End If 

If dstfile = Empty Then
	' Trim extension (if present) from srcfile and add .fb2
	Dim extStart
	extStart = InStrRev(srcfile, ".")
	
	If extStart <> 0 Then 
		dstfile = Left(srcfile, extStart - 1)
	Else
		dstfile = srcfile
	End If
	dstfile = dstfile & ".fb2"
End If

Dim DOM
set DOM = FBApp.Convert(srcfile)
if DOM is Nothing then
	if mute=0 then
		WScript.Echo("Conversion " & srcfile & " -> " & dstfile _
		& " failed: " & vbCr & FBApp.LOG)
	end if
	WScript.Quit 1
else
	DOM.save(dstfile)
	WScript.Quit 0 'comment this line to receive confirmation for successful file conversion
	if mute=0 then
		WScript.Echo("Conversion " & srcfile & " -> " & dstfile _
		& " successful")
	end if
	WScript.Quit 0
end if