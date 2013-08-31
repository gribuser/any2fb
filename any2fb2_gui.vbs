Set args = WScript.Arguments
if args.Count>0 then
	if args.Count>1 then
	  WScript.Echo "Only one parameter supported in interactive mode. Non-GUI script any2fb2.vbs supports all settings via parameters"
	  WScript.Quit 0
	end if
	if args(0)="/?" then
	  WScript.Echo "Usage: any2fb2.vbs <infile>"
	  WScript.Quit 0
	end if
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegWrite "HKCU\Software\Grib Soft\Any to FB2\1.0\LastOpenURI", args(0)
end if

On Error resume next

Set XMLDoc = WScript.CreateObject("MSXML2.DOMDocument.4.0")
If Err Then
Dim MsgText, Btns, oWShellExt
	msgText="Unable to create ActiveX object ""MSXML2.DOMDocument.4.0"". This may be because of missing MSXML 4.0."&_
		vbCr & "You must install MSXML 4.0 or later to use Any2FB. You can manually download MSXML from"&_
		vbCr & vbCr & "http://msdn.microsoft.com/xml/"
	Btns=48
	Err.Clear
	Set oWShellExt = WScript.CreateObject("Shell.Application")
	if Err<>0 then
	else
		msgText=msgText+vbCr & vbCr & "Or you can download (~5MB) and install MSXML 4.0 from our mirror NOW."&_
			vbCr & vbCr & "Do you want to install MSXML now?"
		Btns=Btns or 4
	end if
	
	msgBoxResult = msgbox(msgText, Btns)
	if msgBoxResult=6 then
		oWShellExt.ShellExecute "http://www.gribuser.ru/xml/fictionbook/2.0/software/msxml_4.0_sp2.msi"
	end if
  WScript.Quit 1
end if


Set FBApp = CreateObject("any_2_fb2.any2fb2")

If Err Then
  WScript.Echo "Unable to create ActiveX object ""any_2_fb2.any2fb2"". This may be because of the incorrect setup."&_
		vbCr & vbCr & "Please reinstall the program and try rinning this script again."
  WScript.Quit 1
end if

set DOM = FBApp.ConvertInteractive(0,True)

if DOM is Nothing then
   WScript.Quit 0
end if