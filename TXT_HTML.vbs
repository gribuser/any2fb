Public TitlePrinted As Boolean
Public TCR_HTMLerror As Boolean

Public Const max_chapt_len = 40
Const div_end = "</DIV>"
Const nbs1 = "&nbsp;"
Const nbs4 = "&nbsp;&nbsp;&nbsp;&nbsp;"
Const div_center = "<DIV align=center>"
Const div_jus_st = "<DIV align=justify>"
Const brl = "<BR>"

Dim SpaceNum As String

Private InLine1 As String
Private InLine2 As String

           'max line length to print centered
Public Const maxDigLen = 4 'for digits
Public Const maxSpecialLen = 40 'for special words (as Глава)
Public Const maxBetweenTwoEmptyLen = 90 'max line len between two empty line
Public Const maxStarLen = 20 'for ***
Public Const maxCapitalLen = 80 'for capital words
Public Const maxStartFromDig = 40 '80 'for chapters titles starting from digit
Const minAvLen = 200 'min averaged line length
Public Const MaxAuthorLen = 80 'max length of author name
Public Const MaxBookTitleLen = 80 ''max length of book title
    Public Const maxTitleLenForLit = 150 'if the line between two empty and if converted from lit-file
    Public Const MaxLenSubtitle = 80
    Public Const MaxSubtitLineNum = 3
Public Const MaxLinesToAccumulateFirstEmpty = 16
Public Const MaxLinesToAccumulateFirstNoEmpty = 3
Public Const MinEpigraphLineNum = 2
Public Const MaxEpigraphLineNum = 6
Public Const MaxEpigAutLen = 40
Public Const EpigMaxLineLenToJoin = 40
    Public Const MinVersesLineNum = 2
    Public Const MinVersesLineNumNoEmpty = 3
    Public Const MaxVersesLineNum = 16 'HAS TO BE = MaxLinesToAccumulateFirstEmpty
    Public Const MaxLenVerseNoEmpty = 35
    Public Const MaxLinesToAccumulateForVerses = 35 '16
Public MainTitleFound As Boolean
Public MainTitleFoundAsException As Boolean
Private AfterFindBookTitle As Boolean
Public Const MaxLenComaCapital = 40 'line1 ends by coma, line2 starts from capital


Public Function TXT_HTML(TxtFile As String, Optional FromRb As String, _
Optional NoHeader As Boolean = False, Optional SaveFileAs As String, Optional FromPdf = False) As String

On Error Resume Next

If NoHeader Then
Dim MakeChapterBookmarksOld As Boolean, MakeBookContentsOld As Boolean
MakeChapterBookmarksOld = MakeChapterBookmarks: MakeBookContentsOld = MakeBookContents
MakeChapterBookmarks = False: MakeBookContents = False
End If

TitlePrinted = False
CurBookSubtitle = ""
TxtToHtmlLastTitPos = 0
TxtToHtmlFirstAbsTitleFound = False
         
ContentsSignInserted = False
BookContentsStr = ""
AfterTitleFound = False

If InputFileExt <> "lit" Then SpacesNum = "  " Else SpacesNum = " "

NumContBook = 0

If InputFileExt = "lit" Then
 If EmptyLineLim(BookTitle) Then BookTitle = ""
 If EmptyLineLim(BookAuthor) Then BookAuthor = ""
Else
BookTitle = "": BookAuthor = ""
End If


Dim TmpFileName As String, TmpFileName1 As String, LineLen As Long, lineLen1 As Long, _
AvLineLen As Double, linesToTest As Integer, _
OutputFile As String, OutString As String, SearhStr As String, _
LineNum As Long, FormatRec As String, _
parFound As Long, handleIn As Long, HandleOut As Long, _
handleIn1 As Long, HandleOut1 As Long, _
ParLineRat As Double, _
Line1 As String, line2 As String, line3 As String, line4 As String, _
ParToPrint As String, _
AvLeadingSpacesNum As Double, EmptyNum As Integer, FirstParLine As Boolean, _
ContentStrLen As Long, StartTime As Long

               StartTime = timeGetTime

              'make tmp and output files names
TmpFileName = CurrentPath & "\" & "0000_000.tmp0"
TmpFileName1 = CurrentPath & "\" & "0000_001.tmp0"

If NoHeader = False Then
OutputFile = LastTmpDir & "\" & filecommand(GetShortNameNoExtension, TxtFile, "") & ".html0"
Else
OutputFile = SaveFileAs
End If

''''''''''''''''''debug.print OutputFile: End

               'CORRECT tmp-file CONTENT
    filecommand CopyTheFile, TxtFile, TmpFileName

             'detect charset
CharsetDetector TmpFileName

                           
                      If WhereToClean = "0" Or WhereToClean = "2" Then
                      '''''debug.print "WhereToClean=", WhereToClean, "txt to html: cleanig before"
                      CleanUpF.CleanUpBookFile FileContent
                      End If

''''''''''''''''''debug.print FileContent
         Form1.Status.Panels(2).Text = "...converting txt to html..."
         
If InputFileExt = "tcr" Then
 pos0 = InStr(1, FileContent, "PPL", vbTextCompare)
 If pos0 > 0 And pos0 < 3 Then FileContent = Replace(FileContent, "PPL", "", 1, 1, vbTextCompare)
 FileContent = Replace(FileContent, "<* >", "", 1, 1, vbTextCompare)
End If

                'REMOVE HTML SHIT
             'remove <head>
 pos0 = InStr(1, FileContent, "</HEAD>", vbTextCompare)
      If pos0 > 0 Then FileContent = Right(FileContent, Len(FileContent) - pos0 - 6)
             'remove often-used tags
    'FileContent = Remove_Tags(FileContent, "<u[^>]*>", " ", False)
     '   FileContent = Remove_Tags(FileContent, "</u[^>]*>", " ", False)
    'FileContent = Remove_Tags(FileContent, "<h[^>]*>", " ", False)
    '    FileContent = Remove_Tags(FileContent, "</h[^>]*>", " ", False)
    'FileContent = Remove_Tags(FileContent, "<a[^>]*>", " ", False)
    '    FileContent = Replace(FileContent, "</a>", " ", , , vbTextCompare)
    'FileContent = Remove_Tags(FileContent, "<dir[^>]*>", " ", False)
    
''''''''''''''''''debug.print InStr(1, FileContent, "<HTML>", vbTextCompare)
             
If InStr(1, FileContent, "<H", vbTextCompare) Or InStr(1, FileContent, "<dir>", vbTextCompare) _
Or InStr(1, FileContent, "<BODY>", vbTextCompare) Or InStr(1, FileContent, "<p", vbTextCompare) _
Or InStr(1, FileContent, "<DIV", vbTextCompare) Or InStr(1, FileContent, "<A", vbTextCompare) _
Or InStr(1, FileContent, "<PRE", vbTextCompare) _
Then
FileContent = Replace(FileContent, "&lt;", Chr(171))
FileContent = Replace(FileContent, "&gt;", Chr(187))
FileContent = Remove_Tags(FileContent, "<[^>]*>", " ")
FileContent = Remove_Tags(FileContent, "&[^;]{0,8};", " ")
Else
FileContent = Replace(FileContent, "&lt;", Chr(171))
FileContent = Replace(FileContent, "&gt;", Chr(187))
FileContent = Replace(FileContent, "<", Chr(171))
FileContent = Replace(FileContent, ">", Chr(187))
End If
'

           Form1.Status.Panels(2).Text = "...removing bad symbols..."

        RestoreLineEndsGlobal FileContent, False
 
For i = 0 To 8
FileContent = Replace(FileContent, Chr(i), "")
Next
For i = 11 To 12
FileContent = Replace(FileContent, Chr(i), "")
Next
For i = 14 To 31
FileContent = Replace(FileContent, Chr(i), "")
Next



    Form1.Status.Panels(2).Text = "...replacing non-standard symbols..."
                  'too many double line ends ->remove double
Dim StLen As Double
'StLen = Len(FileContent): If StLen < 1 Then StLen = 1
If GetOccurenceNumberFast(FileContent, LineEnd & LineEnd, True) > 500 Then
FileContent = Replace(FileContent, LineEnd & LineEnd, LineEnd)
End If
'End
   ' WriteStringFile "f:\tmp2\Html_Txt.htm", FileContent: End
 
            FileContent = Replace(FileContent, " *", "*")
 If LastCharset <> "Win1251lat" Then
 FileContent = Replace(FileContent, Chr$(132), Chr$(34))
 'FileContent = Replace(FileContent, Chr$(145), Chr$(39)) 'added
 'FileContent = Replace(FileContent, Chr$(146), Chr$(39)) 'added
 FileContent = Replace(FileContent, Chr$(147), Chr$(34))
 FileContent = Replace(FileContent, Chr$(148), Chr$(34))
 End If
FileContent = Replace(FileContent, Chr$(160), " ")
FileContent = Replace(FileContent, Chr$(150), "-")
'FileContent = Replace(FileContent, Chr(151), "-")
FileContent = Replace(FileContent, Chr$(172), "-")
FileContent = Replace(FileContent, Chr$(173), "")
'FileContent = Replace(FileContent, "--", "-")
FileContent = Replace(FileContent, "==", "-")
    If LastCharset <> "Win1251lat" Then
    FileContent = Replace(FileContent, Chr$(168), Chr$(197)) ' Ё -> Е
    FileContent = Replace(FileContent, Chr$(184), Chr$(229)) ' ё -> е
    End If
                 'FileContent = Replace(FileContent, "$", LineEnd)
FileContent = Replace(FileContent, Chr(9), " ")


    If InputFileExt = "kml" Then
    LastFormatType = "tabulators": FormatRec = LastFormatType:
    LineNum = GetOccurenceNumberFast(FileContent, LineEnd, True)
    FileContent = FileContent & LineEnd & LineEnd 'for correct print if on error exit
    'WriteStringFile TmpFileName, FileContent
    GoTo eRestOrig
    
    'FormatRec = "spaces"
    'LineNum = GetOccurenceNumberFast(FileContent, LineEnd, True)
    'FileContent = FileContent & LineEnd & LineEnd 'for correct print if on error exit
    'WriteStringFile TmpFileName, FileContent
    'GoTo eStartRec
    End If

         'find out number of -
NumOfPer = 100# * CDbl(PatternMatchNumber("-" & Chr(13), FileContent, True)) / CDbl(Len(FileContent))
RemovePerenos = False

                    'write corrected tmp-file
    FileContent = FileContent & LineEnd & LineEnd 'for correct print if on error exit
    WriteStringFile TmpFileName, FileContent
                             'user-defined FormatRec ->goto restore
             If LastFormatType <> "auto" Then
                If LastFormatType = "tabulators" Then FormatRec = "tabulators" Else _
                If LastFormatType = "simple" Then FormatRec = "spaces" Else _
                If LastFormatType = "line ends" Then FormatRec = "line ends" Else _
                FormatRec = "advanced"
             LineNum = GetOccurenceNumberFast(FileContent, LineEnd, True)
             GoTo eRestOrig
             End If
                              'short file
             If FileLen(TmpFileName) < 5000 Then
             LineNum = GetOccurenceNumberFast(FileContent, LineEnd, True)
             FormatRec = "advanced": GoTo eStartRec
             End If
   

        Form1.Status.Panels(2).Text = "...analysing text structure..."
                       'FIND PARAGRAPH RECOVERY TYPE
FormatRec = "advanced"
                
           'check averaged line length
  Dim SumAi2 As Double, SumAi As Double, LenVariance As Double, Line1Len As Long, _
  LineEndsNum As Long, TabsToLineNumRat As Double
   SumAi2 = CDbl(0): SumAi = CDbl(0): LineEndsNum = CLng(0)
parFound = 0: LineNum = 1: AvLineLen = 0#: AvLeadingSpacesNum = 0#
handleIn = FreeFile
Open TmpFileName For Input As #handleIn
Do Until EOF(handleIn)
 Line Input #handleIn, Line1
               Line1Len = Len(Line1)
            SumAi = SumAi + Line1Len
            SumAi2 = SumAi2 + Line1Len ^ 2#
 AvLineLen = AvLineLen + CDbl(Len(Line1))
   If Line1Len - Len(LTrim(Line1)) >= 1 Then
   AvLeadingSpacesNum = AvLeadingSpacesNum + CDbl(1)
   End If
 'AvLeadingSpacesNum = AvLeadingSpacesNum + Len(line1) - Len(LTrim(line1))
  LineNum = LineNum + CLng(1)
   If InStr(Line1, SpacesNum) = 1 Then parFound = parFound + 1
Loop 'Do Until EOF(handleIn)
Close #handleIn 'Open TmpFileName For Input As #handleIn

TabsToLineNumRat = AvLeadingSpacesNum / CDbl(LineNum)
AvLineLen = AvLineLen / CDbl(LineNum)
AvLeadingSpacesNum = AvLeadingSpacesNum / CDbl(LineNum)
      LenVariance = (Sqr(SumAi2 - SumAi ^ 2# / CDbl(LineNum)) / CDbl(LineNum)) / AvLineLen

            'small len variance ->remove hyps
 If LenVariance < 0.07 Then
  If NumOfPer >= 0.002 Then
  'RemovePerenos = True
  FileContent = Replace(FileContent, Chr(32) & "-" & LineEnd, Chr(28))
  FileContent = Replace(FileContent, "-" & LineEnd, "")
  FileContent = Replace(FileContent, Chr(28), Chr(32) & "-" & LineEnd)
  End If
 End If
 
            'small line len variance + tabulated pars ->tabulators
  If TabsToLineNumRat > 0.1 And LenVariance < 0.07 Then
  LastFormatType = "tabulators": FormatRec = LastFormatType: GoTo eRestOrig
  End If
        'number of line ends aproximatly equal to the number of tabs -> line ends
  If TabsToLineNumRat > 0.9 And TabsToLineNumRat < 1.1 Then
  LastFormatType = "line ends": FormatRec = LastFormatType: GoTo eRestOrig
  End If
  
  ''''''''''''''''''debug.print AvLineLen, LineNum / CDbl(Len(FileContent))
  
         'long lines -> line ends
  If AvLineLen > 160 Then
  LastFormatType = "line ends": FormatRec = LastFormatType: GoTo eRestOrig
  End If
  
 
             'take out leading spaces if there is more than 7
If AvLeadingSpacesNum > 0.85 And InputFileExt <> "lit" Then
'''''''''''''''''''''debug.print "take out leading spaces"
          Form1.Status.Panels(2).Text = "...removing leading spaces..."
   WriteStringFile TmpFileName1, FileContent
    handleIn1 = FreeFile
    Open TmpFileName1 For Input As #handleIn1
    HandleOut1 = FreeFile
    Open TmpFileName For Output As #HandleOut1
        Do Until EOF(handleIn1)
         Line Input #handleIn1, Line1
         Print #HandleOut1, LTrim(Line1)
        Loop
    Close #HandleOut1
    Close #handleIn1
FormatRec = "advanced"
 FileContent = ReadStringFile(TmpFileName)
GoTo eStartRec
End If 'If AvLeadingSpacesNum > 0.85 And InputFileExt <> "lit" Then
     

    If AvLineLen < minAvLen And LineNum > 0 Then
    ParLineRat = CDbl(parFound) / CDbl(LineNum)
    FormatRec = "advanced"
     If ParLineRat > 0.1 And ParLineRat < 0.5 Then FormatRec = "spaces"
    End If 'If AvLineLen < minAvLen And lineNum > 0 Then

    If InputFileExt = "lit" Then
     If parFound < 1 Then
     FormatRec = "advanced"
     Else
     FormatRec = "spaces"
     LineNum = GetOccurenceNumberFast(FileContent, LineEnd, True)
     End If
    End If 'If InputFileExt = "lit" Then

'

eStartRec:

  
eRestOrig:
                   Form1.Status.Panels(2).Text = "...converting txt to html..."

        If FormatRec = "spaces" Then LastFormatRec = "simple" Else LastFormatRec = FormatRec

FileContent = Replace(FileContent, Chr(151), "-")
FileContent = Replace(FileContent, "--", "-")



    'FileContent = Replace(FileContent, Chr(169), " ")
    'FileContent = RemoveRepeatedSymbols(FileContent, Chr(32))
 
                CleanBookGarbage FileContent, True
 
 FileContent = Replace(FileContent, Chr(1), "")
 FileContent = FileContent & LineEnd & TmpFileEnd & LineEnd & Chr(1) & Chr(1) & LineEnd
      WriteStringFile TmpFileName, FileContent

Dim TotalLines As Double, LineNum0 As Double

If LineNum <= 1 Then TotalLines = CDbl(1) Else: TotalLines = CDbl(100) / CDbl(LineNum)

''''''''''''''''''debug.print FormatRec: End
                         'RECOVER TITLES, EPIGRAPHS, VERSES, ETC
              Form1.Status.Panels(2).Text = "...txt to html...looking for titles..."
 
handleIn = FreeFile: Open TmpFileName For Input As #handleIn
          'open output file for writing
filecommand DeleteTheFile, TmpFileName1, ""
HandleOut = FreeFile: Open TmpFileName1 For Output As #HandleOut

Dim LineTmp As String, LetAsc As Integer, CurPercents As Integer, SearchRes As Boolean
            '
                'FIND BOOK AUTHOR AND TITLE
     TXT_HTMLFindBookTitle handleIn, HandleOut, LineNum, EmptyNum
     TXT_HTMLSubtitleFinger handleIn, HandleOut, "", 1, 60
     TXT_HTMLEpigraphFinger handleIn, HandleOut, "" ', , , True
eInpLine1:
'MainTitleFoundAsException = False
               CurPercents = CInt(CSng(LineNum) * TotalLines)
               If CurPercents > 400 Then GoTo eErrExit
               If CurPercents \ 10 Then
                If CurPercents <= 100 Then Form1.Status.Panels(1).Text = CurPercents & " %"
               End If
               
    Line Input #handleIn, InLine1:
    If InStr(InLine1, Chr(1)) Then GoTo EndFile0
    LineNum = LineNum + 1
eTitCheck:
            'If InStr(InLine1, Chr(32)) = 1 Then FirstSymIsSpace = True Else FirstSymIsSpace = False
      If TXT_HTMLTitleFinger(handleIn, HandleOut, InLine1, "JustCheck") Then
       'If RemovePerenos Then InLine1 = TXT_HTMLRemoveHyphs(InLine1)
          'If InputFileExt = "rb" Then GoTo eInpLine1
      TXT_HTMLTitlePrinter handleIn, HandleOut, InLine1, , , PrintAsSubtitle
       'TXT_HTMLEpigraphFinger handleIn, HandleOut, ""
       SearchRes = TXT_HTMLSubtitleFinger(handleIn, HandleOut, "", 1, 60)
            'If MainTitleFoundAsException = False Then
              If SearchRes Then TXT_HTMLEpigraphFinger handleIn, HandleOut, "", , , , True _
              Else TXT_HTMLEpigraphFinger handleIn, HandleOut, ""
           ' End If
      'MainTitleFoundAsException = False
      GoTo eInpLine1
      End If 'if TXT_HTMLTitleFinger(HandleOut, InLine2) Then
      
             If EmptyLine(InLine1) = False Then
              If FindVerses = 1 And Len(Trim(InLine1)) <= MaxLenVerseNoEmpty Then
               If TXT_HTMLVersesFinger(handleIn, HandleOut, InLine1) Then GoTo eInpLine1
              End If
            
            Print #HandleOut, InLine1
            TitlePrinted = False
            GoTo eInpLine1
            End If 'If EmptyLine(InLine1) = False Then
            
                            'line is empty-> look for subtitle
CheckAgain:
   ' If InputFileExt = "rb" Then GoTo eBrCheck
              SearchRes = TXT_HTMLSubtitleFinger(handleIn, HandleOut, "", 2, 60)
                'If MainTitleFoundAsException = False Then
                SearchRes = TXT_HTMLEpigraphFinger(handleIn, HandleOut, "", , , , True)
                'SearchRes = TXT_HTMLVersesFinger(handleIn, HandleOut, "")
               ' End If
              'SearchRes = TXT_HTMLVersesFinger(handleIn, HandleOut, "")
                If SearchRes Then GoTo eInpLine1
              
              'If FormatRec = "tabulators" Or FormatRec = "line ends" Then
              If FormatRec = "line ends" Then
            
eBrCheck0:                  'put <BR> instead of empty lines
              Line Input #handleIn, LineTmp: If InStr(LineTmp, Chr(1)) Then GoTo EndFile0
              LineNum = LineNum + 1
                If EmptyLine(LineTmp) Then GoTo eBrCheck0
                Print #HandleOut, "<BR>"
               InLine1 = LineTmp: GoTo eTitCheck
             End If 'If FormatRec = "tabulators" Then
             
             
eBrCheck:                   'put <BR> if it does not break a line
              Line Input #handleIn, LineTmp: If InStr(LineTmp, Chr(1)) Then GoTo EndFile0
              LineNum = LineNum + 1
                If EmptyLine(LineTmp) Then GoTo eBrCheck
                LetAsc = Asc(LTrim(LineTmp))
              If LetterIsSmall(LetAsc) Then
                         'join lines
              Print #HandleOut, LineTmp
              TitlePrinted = False
              Else
              Print #HandleOut, "<BR>"
               InLine1 = LineTmp: GoTo eTitCheck
              'Print #HandleOut, LineTmp
              End If 'If LetterIsSmall(LetAsc) Then
            ''''''''''''''''''debug.print InLine1: End
            
GoTo eInpLine1
'Loop

EndFile0:
    Close #HandleOut
    Close #handleIn
''''''''''''''''''debug.print Err.Number: End
    
OutString = ReadStringFile(TmpFileName1)
OutString = OutString & LineEnd & Chr(1) & Chr(1)

''''''''''''''''''debug.print OutString

        WriteStringFile TmpFileName, OutString
        
      '  WriteStringFile "f:\tmp\AfterTit.txt", OutString: End

              'filecommand CopyTheFile, TmpFileName, "F:\tmp3\after titles.txt": End
''''''''''''''''''debug.print FormatRec: End
               'RESTORE PARAGRAPHS
eGetPar:

'FormatRec = "tabulators"

               Form1.Status.Panels(2).Text = "...txt to html...formatting paragraphs..."
''''''''''''''''''debug.print lineNum, TotalLines, 1# / CSng(lineNum) * 100: End
TotalLines = 1# / CSng(LineNum) * 100#
EmptyNum = 0: LineNum0 = 0: LineNum = 0
handleIn = FreeFile: Open TmpFileName For Input As #handleIn
          'open output file for writing
HandleOut = FreeFile: Open OutputFile For Output As #HandleOut

If NoHeader = False Then
                'write html start
    fsd = ConvertFontSizeToDigit(NewFileFontSize)
      CurChar = "windows-1251"
      If LastCharset = "Win1251lat" Then CurChar = "windows-1252"
''''''''''''''''''debug.print LastCharset
    Print #HandleOut, "<HTML>" & LineEnd & _
    "<meta content=""text/html; charset=" & CurChar & """ http-equiv=""Content-Type"">" & _
    ReaderJustified & LineEnd & _
    "<BODY bgColor=" & NewFileBackColor & ">" & LineEnd & _
    "<BASEFONT size=""" & fsd & """" & " face=""" & NewFileFontName & """" & ">" & LineEnd & _
    "<BODY TEXT=" & NewFileForeColor & ">" & LineEnd & _
    "<DIV></DIV>"
End If 'If NoHeader = False Then


      
                'ORIGINAL RECONSTRUCTION

If FormatRec = "line ends" Then
''''''''''''''''''debug.print "line ends"
eInpOneLine:
 If TXT_HTMLInputFirstLine(handleIn, HandleOut, LineNum) = False Then GoTo EndFile1
     CurPercents = CInt(CSng(LineNum) * TotalLines)
        If CurPercents > 400 Then GoTo eErrExit
      If CurPercents \ 10 Then
       If CurPercents <= 100 Then Form1.Status.Panels(1).Text = CurPercents & " %"
      End If
     'TXT_HTMLRemoveHyphs InLine1, RemovePerenos
     TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
     GoTo eInpOneLine
End If 'If FormatRec = "line ends" Then


          'loop over file lines
eInputTwoLines:
 If TXT_HTMLInputFirstLine(handleIn, HandleOut, LineNum) = False Then GoTo EndFile1
 If TXT_HTMLInputSecondLine(handleIn, HandleOut, LineNum, InLine1) = False Then GoTo EndFile1
  GoTo eInputNothing
  
EmptyCheck:
FirstParLine = True

eInputLine2:
 If TXT_HTMLInputSecondLine(handleIn, HandleOut, LineNum, InLine1, FirstParLine) = False Then
    If EmptyLine(InLine1) = False Then TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
  GoTo EndFile1
 End If
 
eInputNothing:

      CurPercents = CInt(CSng(LineNum) * TotalLines)
        If CurPercents > 400 Then
        GoTo eErrExit
        End If
      If CurPercents \ 10 Then
       If CurPercents <= 100 Then Form1.Status.Panels(1).Text = CurPercents & " %"
      End If
      
If FormatRec = "tabulators" Then
    
''''''''''''''''''debug.print AscCod, InLine2
             AscCod = Asc(Trim(InLine2) & " ")
     If LetterIsSmall(AscCod) And FromPdf = False Then
     'TXT_HTMLRemoveHyphs InLine1, RemovePerenos
     InLine1 = RTrim(InLine1) & " " & Trim(InLine2)
     GoTo eInputLine2
     End If

             AscCod = Asc(InLine2 & " ")
     If AscCod <> 32 And AscCod <> 133 Then
     'TXT_HTMLRemoveHyphs InLine1, RemovePerenos
     InLine1 = InLine1 & " " & LTrim(InLine2)
     GoTo eInputLine2
     End If 'If AscCod <> 32 Then
     
     TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
      ' '''''''''''''''''debug.print InLine1
     InLine1 = InLine2
     GoTo EmptyCheck
     
End If 'If FormatRec = "tabulators" Then
      
      
                'SPACE RECONSTRUCTION
If FormatRec = "spaces" Then

         AscCod = Asc(InLine2 & " ")
     If AscCod <> 32 Then
     'TXT_HTMLRemoveHyphs InLine1, RemovePerenos
     InLine1 = InLine1 & " " & LTrim(InLine2)
           FirstParLine = False
     GoTo eInputLine2
     End If 'If AscCod <> 32 Then
      
     
If FirstParLine Then GoTo ePrr

   pEnd = Asc(Right(" " & Trim(InLine1), 1))
   
    ''''''''''''''''''debug.print pend
   If pEnd = 32 Then
   InLine1 = InLine1 & " " & LTrim(InLine2)
      FirstParLine = False
      GoTo eInputLine2
      
   End If
   
    pst = Asc(Left(Trim(InLine2), 1) & " ")
        'print if line2 starts from [
    If pst = 91 Then GoTo ePrr
        'line1 ends by . line2 starts from .  ->print line1
    If pEnd = 46 And pst = 46 Then
    ''''''''''''''''''debug.print InLine2
    GoTo ePrr
    End If
    
    
                  'for english texts
   If LastCharset = "Win1251lat" Then
    If (pst = 39 And (pEnd = 46 Or pEnd = 133 Or pEnd = 39)) _
    Or (pst = 34 And (pEnd = 46 Or pEnd = 133 Or pEnd = 34)) _
    Or ((pst = 145 Or pst = 146) And (pEnd = 46 Or pEnd = 133 Or pEnd = 145 Or pEnd = 146)) _
    Or ((pst = 147 Or pst = 148) And (pEnd = 46 Or pEnd = 133 Or pEnd = 147 Or pEnd = 148)) _
    Or ((pst = 171) And (pEnd = 46 Or pEnd = 133 Or pEnd = 187)) _
    Or (((pst >= 65 And pst <= 91) Or (pst >= 192 And pst <= 223)) And (pEnd = 34 Or pEnd = 39 Or (pEnd >= 145 And pEnd <= 148) Or pEnd = 187)) _
    Then
    GoTo ePrr
    End If
   End If
   
                  'starts from ", ends by . or ... or "
   If (pst = 34) And (pEnd = 46 Or pEnd = 133 Or pEnd = 34) Then
   'If (pst = 34 Or pst = 145 Or pst = 147) _
   'And (pend = 46 Or pend = 133 Or pend = 34 Or pend = 146 Or pend = 148) Then
   GoTo ePrr
   End If
      
       'repeat if the line is not ended by .,; etc
    '          ;            !           .             :
   If (pEnd <> 33 And pEnd <> 46 And pEnd <> 58 And pEnd <> 59 _
   And pEnd <> 63 And pEnd <> 44 And pEnd <> 133 And pEnd <> 187) Then
   InLine1 = InLine1 & " " & LTrim(InLine2)
      FirstParLine = False
      'If TXT_HTMLInputSecondLine(handleIn, HandleOut, lineNum, InLine1, FirstParLine) = False Then GoTo EndFile1
      'GoTo eInputNothing
      GoTo eInputLine2
   End If 'If (pend <> 33 And pend <> 46 And pend <> 58
      
        'line1 ends by "," line2 starts from capital
      If pEnd = 44 Then
       If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223)) Then
            'epigraph?->print line1
         If Len(Trim(InLine1)) < MaxLenComaCapital Then GoTo ePrr
       InLine1 = InLine1 & " " & LTrim(InLine2)
        FirstParLine = False
        'If TXT_HTMLInputSecondLine(handleIn, HandleOut, lineNum, InLine1, FirstParLine) = False Then GoTo EndFile1
        'GoTo eInputNothing
        GoTo eInputLine2
       End If
      End If
    
          'pst <> "-"
      If pst <> 45 And pst <> 32 Then
                  'repeat if line2 does not start from capital letter
'If ((pst >= 0 And pst <= 33) Or (pst >= 35 And pst <= 48)
       If ((pst >= 0 And pst <= 33) Or (pst >= 35 And pst <= 41) Or (pst >= 43 And pst <= 48) _
       Or (pst >= 58 And pst <= 64) Or (pst >= 91 And pst <= 171) _
       Or (pst >= 172 And pst <= 191) Or pst >= 224) Then
         InLine1 = InLine1 & " " & LTrim(InLine2)
         FirstParLine = False
         'If TXT_HTMLInputSecondLine(handleIn, HandleOut, lineNum, InLine1, FirstParLine) = False Then GoTo EndFile1
         'GoTo eInputNothing
         GoTo eInputLine2
       End If 'If ((pst >= 0 And pst <= 33) Or (pst >= 35 And pst <= 48)
      End If 'If pst <> 45 Then
    
    If pst = 45 Then
    '''''''''''''''''''debug.print InLine2
           'pst = "-", repeat if the letter after "-" is not capital
         pcheck = Replace(InLine2, " ", "")
        If Len(pcheck) > 2 Then
        pst = Asc(Mid(pcheck, 2, 1) & " ")
                 'If pst = Chr(34) Then GoTo ePrr
           If ((pst >= 0 And pst <= 64) Or (pst >= 91 And pst <= 191) _
           Or pst >= 224) Then
           InLine1 = InLine1 & " " & LTrim(InLine2)
           FirstParLine = False
           'If TXT_HTMLInputSecondLine(handleIn, HandleOut, lineNum, InLine1, FirstParLine) = False Then GoTo EndFile1
           'GoTo eInputNothing
           GoTo eInputLine2
           End If 'If ((pst >= 0 And pst <= 64) Or (pst >= 91 And pst <= 191)
         End If 'If Len(pcheck) > 2 Then
      End If 'If pst = 45 Then
         
ePrr:
                 'print line1 and continue
    
     TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
      ' '''''''''''''''''debug.print InLine1
     InLine1 = InLine2
     GoTo EmptyCheck
     
     
End If 'If FormatRec = "spaces" Then

                'ADVANCED RECONSTRUCTION
If FormatRec = "advanced" Then

   If InLine1 = "" Then
   InLine1 = InLine2
   FirstParLine = True
   GoTo eInputLine2
   End If
  
             'single letter line1->sum up
   If Len(Trim(InLine1)) = 1 Then
    let0 = Trim(InLine1)
    If let0 <> "0" And let0 <> "1" And let0 <> "2" And let0 <> "3" And let0 <> "4" _
       And let0 <> "5" And let0 <> "6" And let0 <> "7" And let0 <> "8" _
       And let0 <> "9" And let0 <> "I" And let0 <> "X" And let0 <> "V" And let0 <> "L" Then
         InLine1 = InLine1 & " " & LTrim(InLine2)
         FirstParLine = False
         GoTo eInputLine2
    End If 'If let0
   End If 'If Len(Trim(InLine1)) = 1 Then
            'single letter line2->sum up
   If Len(Trim(InLine2)) = 1 Then
   ''''''''''''''''''''debug.print InLine2
     let0 = Trim(InLine2)
    If let0 <> "0" And let0 <> "1" And let0 <> "2" And let0 <> "3" And let0 <> "4" _
       And let0 <> "5" And let0 <> "6" And let0 <> "7" And let0 <> "8" _
       And let0 <> "9" And let0 <> "I" And let0 <> "X" And let0 <> "V" And let0 <> "L" Then
         InLine1 = InLine1 & " " & LTrim(InLine2)
         FirstParLine = False
         GoTo eInputLine2
    End If 'If let0
   End If 'If Len(Trim(InLine2)) = 1 Then


          
  'AscCod = Asc(InLine2 & " ")
  
   
   InLine2 = Trim(InLine2)
    pst = Asc(Left(Trim(InLine2), 1) & " ")
          'print line1 if line2 starts from [
    If pst = 91 Then
     TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
     InLine1 = InLine2
     FirstParLine = True
     GoTo eInputLine2
    End If
    
    '       pend = Asc(Right(" " & InLine1, 1))
   ' If pend = 32 Then
    
   '    If Len(Trim(InLine1)) < MaxLenComaCapital Then GoTo eEprFirstLine
   ' GoTo eSumm
   ' End If
           'check line1 end
        pEnd = Asc(Right(" " & Trim(InLine1), 1))
   
                 'for english texts
   If LastCharset = "Win1251lat" Then
    If (pst = 39 And (pEnd = 46 Or pEnd = 133 Or pEnd = 39)) _
    Or (pst = 34 And (pEnd = 46 Or pEnd = 133 Or pEnd = 34)) _
    Or ((pst = 145 Or pst = 146) And (pEnd = 46 Or pEnd = 133 Or pEnd = 145 Or pEnd = 146)) _
    Or ((pst = 147 Or pst = 148) And (pEnd = 46 Or pEnd = 133 Or pEnd = 147 Or pEnd = 148)) _
    Or ((pst = 171) And (pEnd = 46 Or pEnd = 133 Or pEnd = 187)) _
    Or (((pst >= 65 And pst <= 91) Or (pst >= 192 And pst <= 223)) And (pEnd = 34 Or pEnd = 39 Or (pEnd >= 145 And pEnd <= 148) Or pEnd = 187)) _
    Then
    'If (pst = 39 Or pst = 34 Or pst = 145 Or pst = 146 Or pst = 171) And pend <> "," Then
    TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
         InLine1 = InLine2
     FirstParLine = True
     GoTo eInputLine2
    End If
   End If
   
   
    'If InStr(InLine1, "Если нахлынет") Or InStr(InLine2, "Если нахлынет") Then
   '  '''''''''''''''''debug.print "InLine1", InLine1: 'End
    ' '''''''''''''''''debug.print "InLine2", InLine2
   '  '''''''''''''''''debug.print , FirstParLine
    ' End If
      
       'line1 ends by   ! . : ; ? , ... >>
   If (pEnd = 33 Or pEnd = 46 Or pEnd = 58 Or pEnd = 59 _
   Or pEnd = 63 Or pEnd = 44 Or pEnd = 133 Or pEnd = 187) Then
            
e555:
               'check line2 start
           ' If pst = 32 Then GoTo eSumm
        'line1 ends by . line2 starts from . ->print line1
            If pEnd = 46 And pst = 46 Then GoTo eEprFirstLine
        'line1 ends by "," line2 starts from capital->summ up if line1 is not short
      If pEnd = 44 Then
       If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223)) Then
          'epigraph?->print InLine1
              If Len(InLine1) < MaxLenComaCapital Then GoTo eEprFirstLine
       InLine1 = InLine1 & " " & LTrim(InLine2)
         FirstParLine = False
         GoTo eInputLine2
       End If
      End If 'If pend = 44 Then
   
     
          'print line1 if line2 starts from the capital letter
     'If ((pst >= 65 And pst <= 91) Or (pst >= 192 And pst <= 223) _
    'Or pst = 34 Or pst = 171) Then
     If LetterIsCapital(pst) Or pst = 34 Or pst = 171 Then
eEprFirstLine:
     TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
     InLine1 = InLine2
     GoTo eInputLine2
     Else
                     'print line1 if line2 starts from "-" following by capital letter
            If pst = 45 Then
            pcheck = Replace(InLine2, " ", "")
              If Len(pcheck) > 2 Then
              pst = Asc(Mid(pcheck, 2, 1))
               If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223) _
               Or pst = 34 Or pst = 171) Then
                               'print line1
               GoTo eEprFirstLine
               '                      'line1 is a paragraph
                '   TXT_HTMLParagraphPrinter handleIn, HandleOut, InLine1
                '   InLine1 = InLine2
                 '  GoTo eInputLine2
                               
               End If 'If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223)
              End If 'If Len(pcheck) > 1 Then
            End If 'If pst = 45
                    'print line1 if line2 starts from digit following by .
            If LetterIsDigit(pst) Then
            pcheck = Replace(InLine2, " ", "")
              If Len(pcheck) > 2 Then
              pst = Asc(Mid(pcheck, 2, 1))
               If pst = 46 Then
                               'print line1
               GoTo eEprFirstLine
               End If 'If pst = 46 Then
              End If 'If Len(pcheck) > 1 Then
            End If 'If LetterIsDigit(pst) Then
     End If 'If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223)
              'line1 does not end by  ! . : ; ? , ... >>, line2 starts from capital, both lines are short
   Else
       If Len(InLine1) < MaxLenComaCapital And Len(InLine2) < MaxLenComaCapital Then
        If ((pst >= 65 And pst <= 90) Or (pst >= 192 And pst <= 223) _
       Or pst = 34 Or pst = 171) Then GoTo eEprFirstLine
       End If
   End If 'If (pend = 33 Or pend = 46 Or pend = 58 Or pend = 59
eSumm:

     'TXT_HTMLRemoveHyphs InLine1, RemovePerenos
     InLine1 = InLine1 & " " & LTrim(InLine2)
                        
         FirstParLine = False
         GoTo eInputLine2
         
     'GoTo eInputNothing 'repeat if line2 is not a new par
End If 'If FormatRec = "advanced" Then



                 
'Loop
  

EndFile1:
              On Error Resume Next
                           ''''''''''''''''''debug.print "end file1": ' End
         Close #HandleOut
         Close #handleIn
 
                    Form1.Status.Panels(2).Text = "...txt to html...cleaning book..."
                    
 OutString = ReadStringFile(OutputFile)
  RemoveDesignerSign OutString
 
  
 If NoHeader = False Then
 MakeDesignerSign OutString
 OutString = OutString & LineEnd & "</BASEFONT>" & LineEnd & "</BODY></HTML>"
 End If
 
  'WriteStringFile "f:\tmp3\before titles.txt", OutString

             'add book contents
 If MakeBookContents Then
 ''''''''''''''''''debug.print "MakeBookContents", BookContentsStr
   If BookContentsStr <> "" Then
   Dim ContName As String
    '''''''''''''''''''''debug.print LastCharset
    If LastCharset = "Win1251lat" Then
    ContName = "CONTENTS"
    Else
    ContName = Chr(209) & Chr(206) & Chr(196) & Chr(197) & Chr(208) & Chr(198) & Chr(192) & _
    Chr(205) & Chr(200) & Chr(197) ' "СОДЕРЖАНИЕ"
    End If
    ContName = "<FONT color=" & BookContentColor & ">" & ContName & "</FONT>"
    BookContentsStr = _
    "<SPAN id=BCONTENTS>" & LineEnd & _
    "<DIV align=center><B>" & ContName & "</B></DIV><BR>" & LineEnd & _
    BookContentsStr & LineEnd & _
    "</SPAN><BR>"
   OutString = Replace(OutString, ContentsSign, BookContentsStr)
   Else
   OutString = Replace(OutString, ContentsSign, "")
   End If 'If BookContentsStr <> "" Then
 End If 'If MakeBookContents Then
 
 ''''''''''''''''''debug.print ContentsSign: End
 
 OutString = Replace(OutString, "<BR>" & LineEnd, "<BR>")
 OutString = RemoveRepeatedSymbols(OutString, "<BR>")
 OutString = Replace(OutString, "<BR>", "<BR>" & LineEnd)
 OutString = Replace(OutString, "</H2></DIV>" & LineEnd & "<BR>", "</H2></DIV>")
 OutString = RemoveRepeatedSymbols(OutString, LineEnd)
 
 OutString = RemoveRepeatedSymbols(OutString, Chr(32))
 
 InsertPageBreaksToHtml0 OutString
                      
                      If WhereToClean = "1" Or WhereToClean = "2" Then
                      '''''debug.print "WhereToClean=", WhereToClean, "txt to html: cleanig after"
                      CleanUpF.CleanUpBookFile OutString
                      End If
                      
                      ReplaceBrByDiv OutString
        
 Call WriteStringFile(OutputFile, OutString)

 filecommand DeleteTheFile, TmpFileName, ""
 filecommand DeleteTheFile, TmpFileName1, ""
 
 If NoHeader Then
 MakeChapterBookmarks = MakeChapterBookmarksOld: MakeBookContents = MakeBookContentsOld
 End If
 
  '''''''''''''''''''''debug.print "TXT_HTML: after DeleteTheFile "
 TXT_HTML = OutputFile


GoTo end0
eErrExit:
TXT_HTML = ""


end0:
                    TxtHtmlConvertingTime = timeGetTime - StartTime
''''''''''''''''''debug.print "end0"
LastFormatType = "auto"
LastConverter = ""
TxtToHtmlFirstAbsTitleFound = False
End Function


Public Sub TXT_HTMLParagraphPrinter(handleIn As Long, HandleOut As Long, LineToPrint0 As String, _
Optional NoEmptyCheck As Boolean = False)
On Error Resume Next
   'MainTitleFoundAsException = False
Dim LineToPrint As String
  LineToPrint = Trim(LineToPrint0)
   If Len(LineToPrint) < 1 Then GoTo end0
   
    'If EmptyLineLim(LineToPrint) Then GoTo end0
   ReplaceUnderscoreByItalic LineToPrint
 'print paragraph
  If Left$(LineToPrint, 1) = "-" Then
  LineToPrint = div_jus_st & nbs4 & "-" & nbs1 & _
  Trim(Right$(LineToPrint, Len(LineToPrint) - 1)) & div_end '& LineEnd
  Else
  LineToPrint = div_jus_st & nbs4 & LineToPrint & div_end '& LineEnd
  End If
Print #HandleOut, LineToPrint
 
AfterTitleFound = False
end0:
End Sub


Public Sub TXT_HTMLEpigraphPrinter(handleIn As Long, HandleOut As Long, LinesArr() As String, _
Optional PrintAsParagraph As Boolean, Optional LastLineIsAuthor As Boolean, _
Optional LinesNumToPrint As Integer = 0)

On Error Resume Next
'''''''''''''''''''debug.print "TXT_HTMLSubtitleFinger"

Dim StrToPrint As String, CurLine As String, AllLinesEmpty As Boolean, LinNum As Integer
AllLinesEmpty = True
StrToPrint = ""
If LinesNumToPrint <> 0 Then LinNum = LinesNumToPrint Else LinNum = UBound(LinesArr)
'LinNum = UBound(LinesArr)
  For j = 1 To LinNum
   CurLine = LinesArr(j)
   '''''''''''''''''''debug.print j, CurLine
   ReplaceUnderscoreByItalic CurLine, True
   If EmptyLineLim(CurLine) = False Then
   AllLinesEmpty = False
    If LastLineIsAuthor And j = LinNum And PrintAsParagraph = False Then GoTo eCheck
    StrToPrint = StrToPrint & div_jus_st & CurLine & div_end & LineEnd
   End If
  Next
eCheck:
    If AllLinesEmpty = False Then
     If PrintAsParagraph Then
     Print #HandleOut, StrToPrint & "<BR>": GoTo end0
     End If 'If PrintAsParagraph Then
     StrToPrint = "<FONT color=" & EpigraphColor & ">" & StrToPrint & "</FONT>"
      
     If LastLineIsAuthor Then
     LinesArr(LinNum) = "<FONT color=" & TextAuthorColor & ">" & LinesArr(LinNum) & LineEnd & "</FONT>"
     StrToPrint = StrToPrint & "<SPAN id=txtaut>" & div_jus_st & LinesArr(LinNum) & div_end & "</SPAN>"
     End If 'If LastLineIsAuthor Then
      
   StrToPrint = "<BR><SPAN id=epigraph><I>" & StrToPrint & "</I></SPAN><BR>"
       
     Print #HandleOut, StrToPrint
     TitlePrinted = True

ePrr:
     
    End If 'If AllLinesEmpty = False Then
   
end0:
End Sub


Public Sub TXT_HTMLVersesPrinter(handleIn As Long, HandleOut As Long, LinesArr() As String, _
Optional PrintAsParagraph As Boolean = False, Optional LastLineIsAuthor As Boolean = False, _
Optional LinesNumToPrint As Integer = 0)

On Error Resume Next
'''''''''''''''''''debug.print "TXT_HTMLSubtitleFinger"
LastLineIsAuthor = False

Dim StrToPrint As String, CurLine As String, AllLinesEmpty As Boolean, LinNum As Integer
AllLinesEmpty = True
StrToPrint = ""
If LinesNumToPrint <> 0 Then LinNum = LinesNumToPrint Else LinNum = UBound(LinesArr)
  For j = 1 To LinNum
   CurLine = LinesArr(j)
   ReplaceUnderscoreByItalic CurLine, True
   If EmptyLineLim(CurLine) = False Then
   AllLinesEmpty = False
    If LastLineIsAuthor And j = LinNum And PrintAsParagraph = False Then GoTo eCheck
    StrToPrint = StrToPrint & div_jus_st & CurLine & div_end & LineEnd
   End If
  Next
eCheck:
    If AllLinesEmpty = False Then
     If PrintAsParagraph Then
     Print #HandleOut, StrToPrint & "<BR>": GoTo end0
     End If 'If PrintAsParagraph Then
     StrToPrint = "<FONT color=" & VerseColor & ">" & StrToPrint & "</FONT>"
     
     If LastLineIsAuthor Then
     LinesArr(LinNum) = "<FONT color=" & TextAuthorColor & ">" & LinesArr(LinNum) & LineEnd & "</FONT>"
     StrToPrint = StrToPrint & "<SPAN id=txtaut>" & div_jus_st & LinesArr(LinNum) & div_end & "</SPAN>"
     End If 'If LastLineIsAuthor Then
      
 
 
 StrToPrint = "<BR><SPAN id=verse><I>" & StrToPrint & "</I></SPAN><BR>"
     
     Print #HandleOut, StrToPrint
     TitlePrinted = False
     

ePrr:
     
    End If 'If AllLinesEmpty = False Then
   
end0:
End Sub




Public Sub TXT_HTMLTitlePrinter(handleIn As Long, HandleOut As Long, LineToPrint0 As String, _
Optional NoEmptyCheck As Boolean = False, Optional NoBreakBefore As Boolean = False, _
Optional PrintAsSubtitle As Boolean = False, Optional GoToPrint As Boolean = False)
  'GoTo end0
On Error Resume Next
Dim LineToPrint As String, CurFileInPos As Long

 LineToPrint = Trim(LineToPrint0)
 If EmptyLineLim(LineToPrint) Then GoTo end0
 
 If InStr(LineToPrint, TmpFileEnd) = 1 Then GoTo end0
 
       If GoToPrint Then GoTo ePrint
 
   For i = 0 To UBound(AbsoluteTitleExceptions)
      If InStr(1, LineToPrint, AbsoluteTitleExceptions(i), vbTextCompare) Then
      TXT_HTMLParagraphPrinter handleIn, HandleOut, LineToPrint
      GoTo end0
      End If
    Next
 
            CurFileInPos = Seek(handleIn)
                 'absolute titles keywords
  If Len(LineToPrint) < maxSpecialLen Then
    For j = 0 To UBound(AbsoluteTitles)
      If (InStr(1, LineToPrint, AbsoluteTitles(j), vbTextCompare) = 1) Then
      PrintAsSubtitle = False: TxtToHtmlFirstAbsTitleFound = True:
      ''''''''''''''''''debug.print LineToPrint
      GoTo eBookMarks
      End If
    Next
  End If 'If (lineLen < maxSpecialLen And lineLen > 0) Then
  
                  'absolute titles exceptions
  If Len(LineToPrint) < maxSpecialLen Then
    For j = 0 To UBound(TitleExceptions)
      If (InStr(1, LineToPrint, TitleExceptions(j), vbTextCompare) = 1) Then
      PrintAsSubtitle = True: GoTo ePrint
      End If
    Next
  End If 'If (lineLen < maxSpecialLen And lineLen > 0) Then
 
    If TxtToHtmlFirstAbsTitleFound = False And CurFileInPos < TxtToHtmlMinFirstTitSt Then
    PrintAsSubtitle = False
    End If
    
    If TxtToHtmlFirstAbsTitleFound And CurFileInPos - TxtToHtmlLastTitPos < TxtToHtmlMinTitlesSep Then
    PrintAsSubtitle = True
    GoTo ePrint
    End If
    
    If CurFileInPos > TxtToHtmlMinFirstTitSt And CurFileInPos - TxtToHtmlLastTitPos < TxtToHtmlMinTitlesSep Then
    PrintAsSubtitle = True
    GoTo ePrint
    End If
    
  
    If PrintAsSubtitle Then GoTo ePrint
    
eBookMarks:

If CurBookSubtitle <> "" And InStr(LineToPrint, CurBookSubtitle) Then GoTo ePrint

If MakeChapterBookmarks = False And MakeBookContents = False Then GoTo ePrint
  
  For i = 0 To UBound(BookmarksExeptions)
   If InStr(LineToPrint, BookmarksExeptions(i)) = 1 Then GoTo ePrint
  Next
  
 Dim LineIni As String, BookMarkID As String, BookMarkName As String, BookMarkLinkId As String
 LineIni = LineToPrint
 NumContBook = NumContBook + 1
  BookMarkID = BookMarkIdSt & Trim(CStr(NumContBook))
  BookMarkName = BookMarkNameSt & Trim(CStr(NumContBook))

               'make chapter bookmarks
  If MakeChapterBookmarks Then
  LineToPrint = "<A id=" & BookMarkID & " name=" & BookMarkName & ">" & LineIni & "</A>"
  End If 'If MakeChapterBookmarks Then
              'make book contents
 If MakeBookContents Then
    If MakeChapterBookmarks = False Then
    LineToPrint = "<A id=" & BookMarkID & " name=" & BookMarkName & ">" & LineIni & "</A>"
    End If 'If MakeChapterBookmarks = False Then
     BookMarkLinkId = BookMarkLinkIdSt & Trim(CStr(NumContBook))
     BookContentsStr = BookContentsStr & _
     "<DIV align=justify>" & _
     "<A id=" & BookMarkLinkId & " href=""#" & BookMarkName & """>" & LineIni & "</A></DIV>" & LineEnd
 End If 'If MakeBookContents Then
         
ePrint:


 If PrintAsSubtitle = False Then
 LineToPrint = "<FONT color=" & ChapterTitleColor & ">" & LineToPrint & "</FONT>"
 Else
 LineToPrint = "<FONT color=" & SubtitleColor & ">" & LineToPrint & "</FONT>"
 End If


LineToPrint = "<B>" & LineToPrint & "</B>"
LineToPrint = div_center & LineToPrint & div_end & LineEnd

 If PrintAsSubtitle = False Then
    If NoBreakBefore Then
    LineToPrint = "<SPAN id=title>" & LineToPrint & "</SPAN><BR>"
    Else
    LineToPrint = "<BR><SPAN id=title>" & LineToPrint & "</SPAN><BR>"
    End If
  TxtToHtmlLastTitPos = CurFileInPos
 Else
 LineToPrint = "<BR><SPAN id=subtitle>" & LineToPrint & "</SPAN><BR>"
 End If 'If PrintAsSubtitle = False Then

Print #HandleOut, LineToPrint
If PrintAsSubtitle = False Then TitlePrinted = True Else TitlePrinted = False
AfterTitleFound = True

end0:
         PrintAsSubtitle = False
End Sub



Public Function TXT_HTMLSubtitleFinger(handleIn As Long, HandleOut As Long, LineAfterTit0 As String, _
Optional MaxLinesToAccumulate0 As Integer = MaxSubtitLineNum, Optional MaxTitLineLen0 As Integer, _
Optional CheckForEmpty As Boolean = False, Optional AfterEmptyCall As Boolean) As Boolean


'GoTo end0
On Error GoTo end0
TXT_HTMLSubtitleFinger = True
'MainTitleFoundAsException = False
'''''''''''''''''''debug.print "TXT_HTMLSubtitleFinger"
Dim line0 As String, line2 As String, LineAfterTit  As String, IniFilePos As Long, BeforeEmptyPos As Long, _
MaxLinesToAccumulate As Integer, ItNum As Integer, AccumLine As String, MaxTitLineLen As Integer, _
MaxEmptyLineToSkip As Integer

If MaxTitLineLen0 > 0 Then
MaxTitLineLen = MaxTitLineLen0
Else
  If InputFileExt = "lit" Then MaxTitLineLen = maxTitleLenForLit Else MaxTitLineLen = MaxLenSubtitle
End If 'If MaxTitLineLen0 > 0 Then

MaxLinesToAccumulate = MaxLinesToAccumulate0

MaxEmptyLineToSkip = 10

'MaxTitLineLen = 60
'MaxLinesToAccumulate = 4

   
eInpLine:
IniFilePos = Seek(handleIn): '''''''''''''''''''debug.print IniFilePos
     Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
''''''''''''''''''debug.print LineAfterTit
 If CheckForEmpty And EmptyLineLim(LineAfterTit) = False Then
 GoTo eFileBack
 End If
    'If LineAfterTit0 = "" Then Line Input #handleIn, LineAfterTit Else LineAfterTit = LineAfterTit0
            'first line after title is empty -> start line skipping
ItNum = 1
eSkip1:
eNoEmpty:
 'TXT_HTMLSubtitleFinger = False
     If EmptyLineLim(LineAfterTit) Then
           ItNum = ItNum + 1: If ItNum > MaxEmptyLineToSkip Then GoTo eFileBack
        Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
      GoTo eSkip1
     End If 'If EmptyLineLim(LineAfterTit) Then
''''''''''''''''''debug.print LineAfterTit: 'End
              'first non-empty line found->check if it is suitable for title
    If TXT_HTMLTitleFinger(handleIn, HandleOut, LineAfterTit) Then
    IniFilePos = Seek(handleIn)
       Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
    MainTitleFound = True
    ItNum = 1: GoTo eSkip1
    'GoTo end0
    End If

       'asc0 = Asc(LineAfterTit) & " ": If asc0 = 45 Then GoTo eFileBack
              ''first is suitable for title->find first empty
     LineAfterTit = Trim(LineAfterTit)
     Dim LinesArr() As String
      ReDim Preserve LinesArr(1): LinesArr(1) = LineAfterTit
              'input next line
        BeforeEmptyPos = Seek(handleIn)
          Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
              'next line is not empty ->accumulate
      If EmptyLineLim(LineAfterTit) = False Then
       ItNum = 1
eAccum:
          ItNum = ItNum + 1:  If ItNum > MaxLinesToAccumulate Then GoTo eFileBack
       ReDim Preserve LinesArr(ItNum):  LinesArr(ItNum) = Trim(LineAfterTit)
       BeforeEmptyPos = Seek(handleIn)
          Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
       If EmptyLineLim(LineAfterTit) = False Then GoTo eAccum
      End If 'If EmptyLineLim(LineAfterTit) = False Then
               'empty line found -> find max line length
         Dim MaxLineLen0 As Integer, LinesNum As Integer
         LinesNum = UBound(LinesArr)
               MaxLineLen0 = Len(LinesArr(1))
                  If LinesNum <= 1 Then GoTo ePrint
            For j = 2 To LinesNum
             If Len(LinesArr(j)) > MaxLineLen0 Then MaxLineLen0 = Len(LinesArr(j))
            Next
ePrint:

'If InStr(LinesArr(1), "забью") Then '''''''''''''''''debug.print "here", LinesArr(1), MaxLineLen0, MaxTitLineLen0, LinesNum
                    'print lines
           If MaxLineLen0 < MaxTitLineLen0 Then
                    'check last line -> mayby
                    'check line ends if the subtitle is not after main title
            If MainTitleFound = False Then
            Dim FirstLet As String, LastLet As String, NumOfStopSymb As Integer
             For j = 1 To LinesNum
             FirstLet = Left(LinesArr(j), 1): LastLet = Right(LinesArr(j), 1)
             If LastLet = "," Or LastLet = ":" Or LastLet = ";" Or FirstLet = "-" Then GoTo eFileBack
             Next
               
             ' If j = LinesNum Then
             '  NumOfStopSymb = GetOccurenceNumberFast(LinesArr(LinesNum), ".")
              ' If NumOfStopSymb > 2 Then GoTo eFileBack
             ' End If 'If j = LinesNum Then
            End If 'If MainTitleFound = False Then
            
            NumOfStopSymb = GetOccurenceNumberFast(LinesArr(LinesNum), ".")
               If NumOfStopSymb >= 2 Then GoTo eFileBack
            For j = 1 To LinesNum
            If InStr(LinesArr(j), ".") And InStr(LinesArr(j), ",") Then GoTo eFileBack
            Next
            
            For j = 1 To LinesNum
            ' If InStr(LinesArr(j), "***") Or InStr(LinesArr(j), "* * *") Then
            ' MainTitleFoundAsException = True
            ' End If
             TXT_HTMLTitlePrinter handleIn, HandleOut, LinesArr(j)
            ' if instr()
            Next
             ''''''''''''''''''debug.print "pos after print", Seek(handleIn)
           '  IniFilePos = BeforeEmptyPos ' Seek(handleIn) - 2
             ''''''''''''''''''debug.print "pos after print", IniFilePos, BeforeEmptyPos
           ItNum = 1: LineAfterTit = "":
           CheckForEmpty = True
           TXT_HTMLSubtitleFinger = True
           IniFilePos = Seek(handleIn) - 2: ItNum = -1: GoTo eNoEmpty
           'GoTo eSkip1
           Else
           
           'ItNum = 1: LineAfterTit = "": GoTo eSkip1
           GoTo eFileBack
           End If 'If EmptyLineLim(line2) Then
       
   GoTo end0
       
eFileBack:
TXT_HTMLSubtitleFinger = False
Seek #handleIn, IniFilePos

'TXT_HTMLEpigraphFinger handleIn, HandleOut, ""

end0:
MainTitleFound = False
''''''''''''''''''debug.print "TXT_HTMLSubtitleFinger: pos on exit", Seek(handleIn)
End Function


Public Function TXT_HTMLEpigraphFinger(handleIn As Long, HandleOut As Long, LineAfterTit0 As String, _
Optional MaxLinesToAccumulate0 As Integer = MaxLinesToAccumulateFirstEmpty, _
Optional MaxTitLineLen0 As Integer, _
Optional CheckForEmpty As Boolean = False, Optional AfterEmptyCall As Boolean) As Boolean
'GoTo end0
On Error GoTo end0
    TXT_HTMLEpigraphFinger = False
    If FindEpigraphs = 0 Then GoTo end0
'''''''''''''''''''debug.print "TXT_HTMLEpigraphFinger"
Dim line0 As String, line2 As String, LineAfterTit  As String, IniFilePos As Long, BeforeEmptyPos As Long, _
MaxLinesToAccumulate As Integer, ItNum As Integer, AccumLine As String, MaxTitLineLen As Integer, _
MaxEmptyLineToSkip As Integer, LastLineIsAut As Boolean, NoVersePrint As Boolean, _
SpaceStNum As Integer, NoSpaceStNum As Integer, EpigWasFound As Boolean

If MaxTitLineLen0 = 0 Then MaxTitLineLen = MaxEpigraphLen Else MaxTitLineLen = MaxTitLineLen0
    MaxLinesNum = MaxLinesNum0

MaxLinesToAccumulate = MaxLinesToAccumulate0

MaxEmptyLineToSkip = 10


eInpLine:
 '''''''''''''''''''debug.print IniFilePos

    IniFilePos = Seek(handleIn): '''''''''''''''''''debug.print IniFilePos
     Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
'If InStr(LineAfterTit, "Кукол") Then
''''''''''''''''''debug.print LineAfterTit, MaxLinesToAccumulate: End
'End If
''''''''''''''''''debug.print LineAfterTit
 If CheckForEmpty And EmptyLineLim(LineAfterTit) = False Then
 GoTo eFileBack
 End If
    
 ''''''''''''''''''debug.print "LineAfterTit", LineAfterTit
               ' If EmptyLineLim(LineAfterTit) = False Then GoTo eFileBack
            'first line after title is empty -> start line skipping
ItNum = 1: SpaceStNum = 0: NoSpaceStNum = 0
eSkip1:
     If EmptyLineLim(LineAfterTit) Then
eNoEmpty:
     TXT_HTMLEpigraphFinger = False
          ItNum = ItNum + 1: If ItNum > MaxEmptyLineToSkip Then GoTo eFileBack
       Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
      GoTo eSkip1
     Else
      If ItNum = 1 And AfterEmptyCall = False Then MaxLinesToAccumulate = MaxLinesToAccumulateFirstNoEmpty
     End If 'If EmptyLineLim(LineAfterTit) Then
     
   'If InStr(LineAfterTit, "Троицын день") Then
   ''''''''''''''''''debug.print LineAfterTit
   'End
   'End If
     
 ' If InStr(LineOld, "свадебный") Then
 ' '''''''''''''''''debug.print LineAfterTit
'  End
 ' End If
     
     
     If TXT_HTMLTitleFinger(handleIn, HandleOut, Trim(LineAfterTit)) Then
     TXT_HTMLSubtitleFinger handleIn, HandleOut, "", 1, 60
     GoTo end0
     End If
     
     'If TXT_HTMLSubtitleFinger(handleIn, HandleOut, "", 1, 60, , True) Then GoTo end0
     
'If InStr(LineAfterTit, "изваяньем") Then
''''''''''''''''''debug.print LineAfterTit, Len(Trim(LineAfterTit))
'End
'End If
''''''''''''''''''debug.print "LineAfterTit", LineAfterTit
               'line is not empty->find first empty after it
 If Asc(LineAfterTit) = 32 Then SpaceStNum = SpaceStNum + 1 Else NoSpaceStNum = NoSpaceStNum + 1
     LineAfterTit = Trim(LineAfterTit)
     Dim LinesArr() As String
      ReDim Preserve LinesArr(1): LinesArr(1) = LineAfterTit
              'input next line
        BeforeEmptyPos = Seek(handleIn)
          Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
              'next line is not empty ->accumulate
      If EmptyLineLim(LineAfterTit) = False Then
       ItNum = 1
eAccum:
          ItNum = ItNum + 1:  If ItNum > MaxLinesToAccumulate Then GoTo eFileBack
 If Asc(LineAfterTit) = 32 Then SpaceStNum = SpaceStNum + 1 Else NoSpaceStNum = NoSpaceStNum + 1
       ReDim Preserve LinesArr(ItNum):  LinesArr(ItNum) = Trim(LineAfterTit)
       BeforeEmptyPos = Seek(handleIn)
         Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
       If EmptyLineLim(LineAfterTit) = False Then GoTo eAccum
      End If 'If EmptyLineLim(LineAfterTit) = False Then
                      'mixture of space and no space lines starts->exit
        If SpaceStNum >= 2 And NoSpaceStNum >= 2 Then GoTo eFileBack
                      'empty line found -> find max line length
               Dim LinesNum As Integer
               LinesNum = UBound(LinesArr)
''''''''''''''''''debug.print "LinesNum", LinesNum
                     If LinesNum < MinEpigraphLineNum Then
                     ''''''''''''''''''debug.print LinesArr(1), IniFilePos: End
                     GoTo eFileBack
                     End If
'If InStr(LinesArr(1), "золотым") Then
''''''''''''''''''debug.print LinesArr(1), LinesNum, MinEpigraphLineNum, SpaceStNum, NoSpaceStNum ' AllLetAreCapital, LinesNum, MaxLineLen0, MaxVerseLen
''''''''''''''''''debug.print Asc(LinesArr(LinesNum)), LetterIsCapital(Asc(LinesArr(LinesNum)))
'End
'End If
                       'number of capitals in any line excluding the last one >4 ->exit
                   Dim CapLetNum As Integer
               For j = 1 To LinesNum - 1
                CapLetNum = 0
                For l = 1 To Len(LinesArr(j))
                 If LetterIsCapital(Asc(Mid(LinesArr(j), l, 1))) Then CapLetNum = CapLetNum + 1
                 If CapLetNum > 4 Then GoTo eFileBack
                Next
               'LinesArr(1) = LinesArr(1) & " " & LinesArr(j)
               Next
                       'last line does not start from capital, " ( [ { << print as par->exit
           Dim LetAsc As Integer, AllLetAreCapital As Boolean
              LetAsc = Asc(LinesArr(LinesNum))
                 If LetterIsCapital(LetAsc) Or LetAsc = 40 Or LetAsc = 91 _
                 Or LetAsc = 123 Or LetAsc = 171 Then
                 GoTo eCont
                 Else
         
                  GoTo eFileBack
                 ' End If
                 End If
         
eCont:
                     'if all lines start from capital->print as verses
   ''''''''''''''''''debug.print MaxLineLen0, MaxTitLineLen
                     'continue
               MaxLineLen0 = Len(LinesArr(1))
                   'If LinesNum <= 1 Then GoTo ePrint
            AllLetAreCapital = True
            For j = 1 To LinesNum
                'If GetOccurenceNumberFast(LinesArr(j), ".", True) > 1 Then GoTo eFileBack
               ' If GetOccurenceNumberFast(LinesArr(j), Chr(32), True) < 1 Then GoTo eFileBack
              If AllLetAreCapital Then
                LetAsc = Asc(LinesArr(j))
                   If LetterIsCapital(LetAsc) = False Then AllLetAreCapital = False
                   If LetAsc = 45 Or LetAsc = 171 Then AllLetAreCapital = True
                   'last line starts from ( [ { "->consider as capital
                If j = LinesNum Then
                 If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                 AllLetAreCapital = True
                 LastLineIsAut = True
                 End If 'If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                End If 'If j = LinesNum Then
                     'analize first line
                If j = 1 Then
                     'first line starts from  ( [ { ->exit
                 If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                 GoTo eFileBack
                 End If
                       'first line starts from  " or . or ... ->consider as capital
                 If LetAsc = 34 Or LetAsc = 46 Or LetAsc = 133 Then
                 AllLetAreCapital = True
                  'NoVersePrint = True
                 End If 'If LetAsc = 34 Then
                End If 'If j = 1 Then
              End If 'If AllLetAreCapital Then
                If Len(LinesArr(j)) > MaxLineLen0 Then MaxLineLen0 = Len(LinesArr(j))
             ''''''''''''''''''debug.print j, LinesArr(j)
            Next
            
'If InStr(LinesArr(1), "И вот теперь") Then
''''''''''''''''''debug.print LinesArr(1), LinesNum, AllLetAreCapital, MaxEpigraphLineNum, NoVersePrint, MaxLineLen0, MaxTitLineLen
''''''''''''''''''debug.print LinesArr(LinesNum), InStr(LinesArr(LinesNum), ".")
'End
'End If
         
ePrint:

           If TitlePrinted And LinesNum <= 3 Then GoTo eEpigCheck
           
                         'print as verses if all if all lines start from capital
           If AllLetAreCapital And NoVersePrint = False _
           And LinesNum >= MinVersesLineNum Then 'And MaxLineLen0 <= MaxVerseLen Then
           LastLineIsAut = False: 'GoTo ePrintVerse
            Dim posp As Long, pst As Long, SymStr(), RightLet As String, ExpSymNum As Integer
                            'analize lines structure
                   '2 lines, last line has more than 2 capitals ->last line is text author
             If LinesNum = 2 And GetNumberOfCapitals(LinesArr(LinesNum), 2) >= 2 Then
             GoTo eEpigCheck
             End If
                            'any of lines
             SymStr = Array("!", "?", ":", ";", "-", ".", Chr(133))
             For l = 1 To LinesNum - 1
             
              For j = 0 To UBound(SymStr)
              ExpSymNum = 0
              ExpSymNum = ExpSymNum + GetOccurenceNumberFast(LinesArr(l), CStr(SymStr(j)))
              If ExpSymNum >= 2 Then GoTo eEpigCheck
              Next
             Next
                            'there is ? in the last line->no text author
             If InStr(LinesArr(LinesNum), "?") Or InStr(LinesArr(LinesNum), "!") Then
             LastLineIsAut = False: GoTo ePrintVerse
             End If
                            'last line ends by ... ->no text author
             SymStr = Array(Chr(133), "!", "?", ":", ";", "-")
             For j = 0 To UBound(SymStr)
              If Right(LinesArr(LinesNum), 1) = SymStr(j) Then
              LastLineIsAut = False: GoTo ePrintVerse
              End If
             Next
           
           'GoTo ePrintVerse
              If LastLineIsAut Then GoTo ePrintVerse
                         'LinesNum is odd ->last line is text author
            'If LinesNum > 2 And NumberIsOdd(LinesNum) Then
            'LastLineIsAut = True: GoTo ePrintVerse
            'End If
             
                   'one of SymStr symbols inside last line->text author
               'As String
             SymStr = Array(Chr(34), Chr(171), ".")
             For j = 0 To UBound(SymStr)
               If j < 2 Then pst = 1 Else pst = 2
              posp = InStr(pst, LinesArr(LinesNum), SymStr(j))
              If posp > 1 And posp < Len(LinesArr(LinesNum)) - 2 Then
              LastLineIsAut = True: GoTo ePrintVerse
              End If
             Next
                    'last line starts from one of SymStr symbols ->text author
             SymStr = Array("(", "[", "{", Chr(171))
             For j = 0 To UBound(SymStr)
              If Left(LinesArr(LinesNum), 1) = SymStr(j) Then
              LastLineIsAut = True: GoTo ePrintVerse
              End If
             Next
                       'last line has more than 2 capitals ->last line is text author
             If GetNumberOfCapitals(LinesArr(LinesNum), 2) >= 2 Then
             LastLineIsAut = True: GoTo ePrintVerse
             End If
             
                      
ePrintVerse:
'LastLineIsAut = False
           TXT_HTMLVersesPrinter handleIn, HandleOut, LinesArr(), , LastLineIsAut
           TitlePrinted = False
           TXT_HTMLEpigraphFinger = True
           EpigWasFound = True
           IniFilePos = Seek(handleIn) - 2: ItNum = -1: SpaceStNum = 0: NoSpaceStNum = 0: GoTo eNoEmpty
           End If 'If AllLetAreCapital And LinesNum >= 3 And MaxLineLen0 <= MaxVerseLen Then
       
          
                   'the line before last does not end by . ... ? ! " ->exit
           'LetAsc = Asc(Right(LinesArr(LinesNum - 1), 1))
           ' If LetAsc = 34 Then LetAsc = Asc(Right(LinesArr(LinesNum - 1), 2))
           ' '''''''''''''''''debug.print LetAsc: End
           'If LetAsc <> 46 And LetAsc <> 133 And LetAsc <> 33 And LetAsc <> 63 _
           'Then GoTo eFileBack
           
eEpigCheck:


             'number of lines > MaxEpigraphLineNum ->exit
        If LinesNum > MaxEpigraphLineNum Then GoTo eFileBack

                          'print epigraph
      If MaxLineLen0 <= MaxTitLineLen Then
'If InStr(LinesArr(1), "И вот теперь") Then
''''''''''''''''''debug.print LinesArr(1), AllLetAreCapital, LinesNum, MaxEpigraphLineNum, NoVersePrint, MaxLineLen0, MaxTitLineLen
''''''''''''''''''debug.print LinesArr(LinesNum), InStr(LinesArr(LinesNum), ".")
'End
'End If
                      'check epigraph array: join all lines excluding the last one
            If AllLetAreCapital = False And MaxLineLen0 > EpigMaxLineLenToJoin Then
             Dim HypFound As Boolean
                    LinesArr(1) = RemoveHypsFromSingleLine(LinesArr(1), HypFound)
              For j = 2 To LinesNum - 1
                 'LinesArr(1) = LinesArr(1) & " " & LinesArr(j)
                If HypFound Then
                LinesArr(1) = LinesArr(1) & RemoveHypsFromSingleLine(LinesArr(j), HypFound)
                Else
                LinesArr(1) = LinesArr(1) & " " & RemoveHypsFromSingleLine(LinesArr(j), HypFound)
                End If
              Next
              LinesArr(2) = LinesArr(LinesNum): LinesNum = 2
            End If 'If AllLetAreCapital = False Then
 'End
ePrint1:
                    'analize last line
                LastLineIsAut = False
            If Len(LinesArr(LinesNum)) > MaxEpigAutLen Then GoTo ePrintEpig
            
            If Right(Asc(LinesArr(LinesNum)), 1) = 133 Then
            LastLineIsAut = False: GoTo ePrintEpig
            End If
                  'last line is single word starting from capital->text author
            If LetterIsCapital(Asc(LinesArr(LinesNum))) _
            And InStr(LinesArr(LinesNum), Chr(32)) <= 0 Then
            LastLineIsAut = True: GoTo ePrintEpig
            End If
            
                  'one of SymStr symbols inside last line->text author
            SymStr = Array(Chr(34), Chr(171), ".")
             For j = 0 To UBound(SymStr)
               If j < 2 Then pst = 1 Else pst = 2
              posp = InStr(pst, LinesArr(LinesNum), SymStr(j))
              If posp > 1 And posp < Len(LinesArr(LinesNum)) - 2 Then
              LastLineIsAut = True: GoTo ePrintEpig
              End If
             Next
                    'last line starts from one of SymStr symbols ->text author
             SymStr = Array("(", "[", "{", Chr(171))
             For j = 0 To UBound(SymStr)
              If Left(LinesArr(LinesNum), 1) = SymStr(j) Then
              LastLineIsAut = True: GoTo ePrintEpig
              End If
             Next
                    'last line contains more than 2 capitals ->text author
             If GetNumberOfCapitals(LinesArr(LinesNum), 2) >= 2 Then
             LastLineIsAut = True: GoTo ePrintEpig
             End If
ePrintEpig:
            If TitlePrinted Then
            TXT_HTMLEpigraphPrinter handleIn, HandleOut, LinesArr(), , LastLineIsAut, LinesNum
            Else
                If AllLetAreCapital Then
                TXT_HTMLVersesPrinter handleIn, HandleOut, LinesArr(), , LastLineIsAut, LinesNum
                TitlePrinted = False
                Else
                GoTo eFileBack
                End If
            End If
           TXT_HTMLEpigraphFinger = True
           EpigWasFound = True
           IniFilePos = Seek(handleIn) - 2: ItNum = -1: SpaceStNum = 0: NoSpaceStNum = 0:  GoTo eNoEmpty
           Else
                       'print as paragraph
           'If LinesNum > 1 Then
           'For j = 1 To LinesNum
           'TXT_HTMLParagraphPrinter handleIn, HandleOut, LinesArr(j)
           'Next
           'IniFilePos = Seek(handleIn) - 2: ItNum = -1: SpaceStNum = 0: NoSpaceStNum = 0:  GoTo eNoEmpty
           'Else
           GoTo eFileBack
           'End If
        End If 'If MaxLineLen0 < MaxTitLineLen Then
       
   GoTo end0
       
eFileBack:
Seek #handleIn, IniFilePos
  If LinesNum >= 1 And LinesNum <= 2 And EpigWasFound Then
   If TXT_HTMLSubtitleFinger(handleIn, HandleOut, "", LinesNum, 60) Then
   TXT_HTMLEpigraphFinger = True: GoTo end0
   End If
  End If
TXT_HTMLEpigraphFinger = False

end0:
'''''''''''''''''''debug.print "finished"
''''''''''''''''''debug.print "TXT_HTMLEpigraphFinger: pos on exit", Seek(handleIn)
End Function


Public Function TXT_HTMLVersesFinger(handleIn As Long, HandleOut As Long, LineAfterTit0 As String) As Boolean
'GoTo end0
On Error GoTo end0
    TXT_HTMLVersesFinger = False
    If FindVerses = 0 Then GoTo end0
'''''''''''''''''''debug.print "TXT_HTMLVersesFinger"
Dim line0 As String, line2 As String, LineAfterTit  As String, IniFilePos As Long, BeforeEmptyPos As Long, _
MaxLinesToAccumulate As Integer, ItNum As Integer, AccumLine As String, MaxTitLineLen As Integer, _
MaxEmptyLineToSkip As Integer, LastLineIsAut As Boolean, NoVersePrint As Boolean, _
SpaceStNum As Integer, NoSpaceStNum As Integer, PrintAsPar As Boolean, _
LetAsc As Integer, AllLetAreCapital As Boolean

MaxLinesToAccumulate = MaxLinesToAccumulateForVerses

MaxEmptyLineToSkip = 10
LineAfterTit = LineAfterTit0

IniFilePos = Seek(handleIn)
   
      
'If InStr(LineAfterTit, "И вот теперь") Then
''''''''''''''''''debug.print LineAfterTit, Len(Trim(LineAfterTit))
'End
'End If
''''''''''''''''''debug.print "LineAfterTit", LineAfterTit
               'line is not empty->find first empty after it
 If Asc(LineAfterTit) = 32 Then SpaceStNum = SpaceStNum + 1 Else NoSpaceStNum = NoSpaceStNum + 1
     LineAfterTit = Trim(LineAfterTit):
         If Len(LineAfterTit) > MaxVerseLen Then PrintAsPar = True 'GoTo eFileBack
          LetAsc = Asc(LineAfterTit)
         If (LetterIsCapital(LetAsc) = False Or LetAsc = 45) _
         And (LetAsc <> 133) Then GoTo eFileBack
     Dim LinesArr() As String
      ReDim Preserve LinesArr(1): LinesArr(1) = LineAfterTit
            Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
              'next line is not empty ->accumulate
      If EmptyLineLim(LineAfterTit) = False Then
       ItNum = 1
eAccum:
          ItNum = ItNum + 1:  If ItNum > MaxLinesToAccumulate Then GoTo eFileBack
 If Asc(LineAfterTit) = 32 Then SpaceStNum = SpaceStNum + 1 Else NoSpaceStNum = NoSpaceStNum + 1
       ReDim Preserve LinesArr(ItNum): LinesArr(ItNum) = Trim(LineAfterTit)
        If Len(LinesArr(ItNum)) > MaxVerseLen Then PrintAsPar = True  'GoTo eFileBack
           LetAsc = Asc(LinesArr(ItNum))
         If (LetterIsCapital(LetAsc) = False Or LetAsc = 45) _
         And (LetAsc <> 133) Then GoTo eFileBack
           Line Input #handleIn, LineAfterTit: If InStr(LineAfterTit, Chr(1)) Then GoTo eFileBack
       If EmptyLineLim(LineAfterTit) = False Then GoTo eAccum
      End If 'If EmptyLineLim(LineAfterTit) = False Then
                    'empty line found -> check if the lines are verses
               Dim LinesNum As Integer
               LinesNum = UBound(LinesArr)
                   'too few lines->exit
         If LinesNum <= MinVersesLineNumNoEmpty Then GoTo eFileBack
                  'mixture of space and no space lines starts->exit
        If SpaceStNum <> LinesNum And NoSpaceStNum <> LinesNum Then GoTo eFileBack
                            
                       'one of the lines > MaxVerseLen ->print as par
                    If PrintAsPar Then
                     For j = 1 To LinesNum
                     TXT_HTMLParagraphPrinter handleIn, HandleOut, LinesArr(j)
                     Next
                     TXT_HTMLVersesFinger = True
                     Seek #handleIn, Seek(handleIn) - 2: GoTo end0
                    End If

                    
'If InStr(LinesArr(1), "И вот теперь") Then
''''''''''''''''''debug.print LinesArr(1), LinesNum, MinEpigraphLineNum, SpaceStNum, NoSpaceStNum ' AllLetAreCapital, LinesNum, MaxLineLen0, MaxVerseLen
''''''''''''''''''debug.print Asc(LinesArr(LinesNum)), LetterIsCapital(Asc(LinesArr(LinesNum)))
'End
'End If
                     'any line starts from - ->exit
               For l = 1 To LinesNum '- 1
                If Asc(LinesArr(l)) = 45 Then GoTo eFileBack
               Next

                       'number of capitals in any line excluding the last one >=3 ->exit
                   Dim CapLetNum As Integer
               For j = 1 To LinesNum - 1
                CapLetNum = 0
                For l = 1 To Len(LinesArr(j))
                 If LetterIsCapital(Asc(Mid(LinesArr(j), l, 1))) Then CapLetNum = CapLetNum + 1
                 If CapLetNum >= 3 Then GoTo eFileBack
                Next
               'LinesArr(1) = LinesArr(1) & " " & LinesArr(j)
               Next
                       'last line does not start from capital, " ( [ { << ->exit
                 LetAsc = Asc(LinesArr(LinesNum))
                 If LetterIsCapital(LetAsc) Or LetAsc = 40 Or LetAsc = 91 _
                 Or LetAsc = 123 Or LetAsc = 171 Then
                 GoTo eCont
                 Else
                 GoTo eFileBack
                 End If
         
eCont:
                     'if all lines start from capital->print as verses
                  MaxLineLen0 = Len(LinesArr(1))
                   'If LinesNum <= 1 Then GoTo ePrint
            AllLetAreCapital = True
            For j = 1 To LinesNum
               If AllLetAreCapital Then
                LetAsc = Asc(LinesArr(j))
                   If LetterIsCapital(LetAsc) = False Then AllLetAreCapital = False
                   If LetAsc = 45 Or LetAsc = 171 Then AllLetAreCapital = True
                   'last line starts from ( [ { "->consider as capital
                If j = LinesNum Then
                 If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                 AllLetAreCapital = True
                 LastLineIsAut = True
                 End If 'If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                End If 'If j = LinesNum Then
                     'analize first line
                If j = 1 Then
                     'first line starts from  ( [ { ->exit
                 If LetAsc = 40 Or LetAsc = 91 Or LetAsc = 123 Then
                 GoTo eFileBack
                 End If
                       'first line starts from  " or . or ... ->consider as capital
                 If LetAsc = 34 Or LetAsc = 46 Or LetAsc = 133 Then
                 AllLetAreCapital = True
                  'NoVersePrint = True
                 End If 'If LetAsc = 34 Then
                End If 'If j = 1 Then
              End If 'If AllLetAreCapital Then
             
            Next
            
     If AllLetAreCapital = False Then GoTo eFileBack
            
'If InStr(LinesArr(1), "И вот теперь") Then
''''''''''''''''''debug.print LinesArr(1), LinesNum, AllLetAreCapital, MaxEpigraphLineNum, NoVersePrint, MaxLineLen0, MaxTitLineLen
''''''''''''''''''debug.print LinesArr(LinesNum), InStr(LinesArr(LinesNum), ".")
'End
'End If
         
ePrint:
           
                         'check for text author line
               LastLineIsAut = False: 'GoTo ePrintVerse
            Dim posp As Long, pst As Long, SymStr(), RightLet As String, ExpSymNum As Integer
                            'analize lines structure
                    'any of lines contains more than 2 symbols or starts from - ->exit
             SymStr = Array("!", "?", ":", ";", "-", ".", Chr(133))
             For l = 1 To LinesNum - 1
                'if asc(LinesArr(l))
              For j = 0 To UBound(SymStr)
              ExpSymNum = 0
              ExpSymNum = ExpSymNum + GetOccurenceNumberFast(LinesArr(l), CStr(SymStr(j)))
              If ExpSymNum >= 2 Then GoTo eFileBack
              Next
             Next
                            'last line ends by ... ->no text author
             SymStr = Array(Chr(133), "!", "?", ":", ";", "-")
             For j = 0 To UBound(SymStr)
              If Right(LinesArr(LinesNum), 1) = SymStr(j) Then
              LastLineIsAut = False: GoTo ePrintVerse
              End If
             Next
          
              If LastLineIsAut Then GoTo ePrintVerse
                         
                     '2 lines, last line has more than 2 capitals ->last line is text author
             If LinesNum = 2 And GetNumberOfCapitals(LinesArr(LinesNum), 2) >= 2 Then
             GoTo eFileBack
             End If
                   'one of SymStr symbols inside last line->text author
               'As String
             SymStr = Array(Chr(34), Chr(171), ".")
             For j = 0 To UBound(SymStr)
               If j < 2 Then pst = 1 Else pst = 2
              posp = InStr(pst, LinesArr(LinesNum), SymStr(j))
              If posp > 1 And posp < Len(LinesArr(LinesNum)) - 2 Then
              LastLineIsAut = True: GoTo ePrintVerse
              End If
             Next
                    'last line starts from one of SymStr symbols ->text author
             SymStr = Array("(", "[", "{", Chr(171))
             For j = 0 To UBound(SymStr)
              If Left(LinesArr(LinesNum), 1) = SymStr(j) Then
              LastLineIsAut = True: GoTo ePrintVerse
              End If
             Next
                       'last line has more than 2 capitals ->last line is text author
             If GetNumberOfCapitals(LinesArr(LinesNum), 2) >= 2 Then
             LastLineIsAut = True: GoTo ePrintVerse
             End If
             
                      
ePrintVerse:

           TXT_HTMLVersesPrinter handleIn, HandleOut, LinesArr(), , LastLineIsAut
           TXT_HTMLVersesFinger = True
           Seek #handleIn, Seek(handleIn) - 2: GoTo end0
                             
       
   GoTo end0
       
eFileBack:
Seek #handleIn, IniFilePos
TXT_HTMLVersesFinger = False

end0:
'''''''''''''''''''debug.print "finished"
'''''''''''''''''''debug.print "finished"
''''''''''''''''''debug.print "TXT_HTMLVersesFinger: pos on exit", Seek(handleIn)
End Function


Public Function TXT_HTMLTitleFinger(handleIn As Long, HandleOut As Long, LineToPrint0 As String, _
Optional JustCheck As String) As Boolean
Dim LineToPrint As String, FirstSymIsSpace As Boolean, AscCod As Integer, LastLet As String, LineLen As Integer

TXT_HTMLTitleFinger = True
PrintAsSubtitle = False

On Error Resume Next

'GoTo end0

        If InStr(LineToPrint0, Chr(32)) = 1 Then FirstSymIsSpace = True Else FirstSymIsSpace = False
 LineToPrint = Trim(LineToPrint0)
  
  LineLen = Len(LineToPrint)
         If LineLen < 1 Then GoTo e30
                      'last letter is : ; ->exit
 LastLet = Right(LineToPrint, 1)
 If LastLet = ":" Or LastLet = ";" Then GoTo e30
 
                     '____
If LineLen <= 40 Then
     If InStr(1, LineToPrint, "_______") = 1 Then
    PrintAsSubtitle = True
    If Len(JustCheck) <= 0 Then TXT_HTMLTitlePrinter handleIn, HandleOut, LineToPrint, , , PrintAsSubtitle, True
            GoTo e44
     End If
End If 'If LineLen <= maxStarLen Then
 
                     '---
If LineLen <= maxStarLen Then
 '  If InStr(1, LineToPrint, Chr(150) & Chr(150) & Chr(150) & Chr(150)) = 1 Then 'Or InStr(1, LineToPrint, "- - -") = 1 Then
     If InStr(1, LineToPrint, "---") = 1 Or InStr(1, LineToPrint, "- - -") = 1 Then
'''''''''''''debug.print "here", LineToPrint
    PrintAsSubtitle = True
    If Len(JustCheck) <= 0 Then TXT_HTMLTitlePrinter handleIn, HandleOut, LineToPrint, , , PrintAsSubtitle, True
            GoTo e44
     End If
End If 'If LineLen <= maxStarLen Then
 
                      'first let is - ... ->exit
 AscCod = Asc(LineToPrint & " ")
 If AscCod = 45 Or AscCod = 133 Then GoTo e30

                      'digits
If LineLen <= maxDigLen Then
'AscCod = Asc(LineToPrint & " ")
     If (AscCod >= 49 And AscCod <= 57) Then
          GoTo ePrint
     End If ''if(n_elements(in) eq lineLen) then begin
End If 'IF (lineLen le maxDigLen And lineLen gt 0) Then begin
                    'stars, xxx
If LineLen <= maxStarLen Then
     If InStr(LineToPrint, "***") = 1 Or InStr(LineToPrint, "* * *") = 1 Or _
     InStr(1, LineToPrint, "xxx", vbTextCompare) = 1 Or InStr(1, LineToPrint, "x x x", vbTextCompare) = 1 _
     Then
     PrintAsSubtitle = True:     GoTo ePrint
     End If
End If 'If LineLen <= maxStarLen Then

                    'more than 2 stars
If LineLen <= maxStarLen And GetOccurenceNumberFast(LineToPrint, "*", True) > 2 Then
   If InStr(LineToPrint, "/") Or LetterIsSmall(Asc(LineToPrint & " ")) Then GoTo e30
     PrintAsSubtitle = True:   GoTo ePrint
End If 'If GetOccurenceNumberFast(LineToPrint, "*", True) > 2 Then

                    'SPECIAL WORDS
If (LineLen < maxSpecialLen And SpecialWordsList) Then
    For j = 1 To UBound(SpecialWordsArr)
      If (InStr(LineToPrint, SpecialWordsArr(j)) = 1) Then
      GoTo ePrint
      End If
    Next
End If 'If (lineLen < maxSpecialLen And lineLen > 0) Then
                    'CAPITAL LETTERS
 If (LineLen <= maxCapitalLen) Then
  If LineToPrint = UCase(LineToPrint) Then
     ' AscCod = Asc(LineToPrint & " ")
      If AscCod = 34 Or AscCod = 42 Or AscCod = 171 _
       Or (AscCod >= 48 And AscCod <= 57) _
       Or (AscCod >= 65 And AscCod <= 90) _
       Or (AscCod >= 192 And AscCod <= 223) Then
                GoTo ePrint
      End If 'If (AscCod >= 8 And AscCod <= 57)
   End If 'If LineToPrint = UCase(LineToPrint) Then
 End If 'If (lineLen <= maxCapitalLen) Then
                     'starts from digits:  1. Ужасти Ужасного Короля
If LineLen <= maxStartFromDig And LineLen > 4 Then
''''''''''''''''''debug.print "here", LineToPrint
On Error GoTo e30
                  'GoTo e30
sym1 = CheckSymbolType(Left(LineToPrint, 1))
If sym1 <> "digit" Then GoTo e30
        'check number of points to see if it is reference list or not: 16. Велидов А.С. Похождения террориста, Одиссея Якова Блюмкина
  PointNum = GetOccurenceNumberFast(LineToPrint, Chr(46))
   If PointNum > 2 Then GoTo e30
        'check number of ,
    If PointNum > 1 Then
     If GetOccurenceNumberFast(LineToPrint, Chr(44)) > 1 Then GoTo e30
    End If
  
sym2 = CheckSymbolType(Mid(LineToPrint, 2, 1))
 If sym2 = "point" Then GoTo eMakeTit
 
 If sym2 = "digit" Then
 sym3 = CheckSymbolType(Mid(LineToPrint, 3, 1))
    If sym3 = "point" Then
      sym4 = CheckSymbolType(Mid(LineToPrint, 4, 1))
      If sym4 = "space" Then GoTo eMakeTit
    End If
 End If
 
GoTo e30

eMakeTit:
      GoTo ePrint
End If 'If lineLen <= maxStartFromDig And lineLen > 4 Then

GoTo e30

ePrint:

'TXT_HTMLTitlePrinter handleIn, HandleOut, LineToPrint, , , PrintAsSubtitle

If Len(JustCheck) <= 0 Then TXT_HTMLTitlePrinter handleIn, HandleOut, LineToPrint, , , PrintAsSubtitle
GoTo e44

e30:
TXT_HTMLTitleFinger = False
e44:
 If TXT_HTMLTitleFinger Then MainTitleFound = True
end0:
End Function



Public Function RemoveHypsFromSingleLine(LineStr As String, HypFound As Boolean) As String
On Error Resume Next

Dim FileStr As String
HypFound = False
FileStr = Replace(LineStr, LineEnd, " ")
FileStr = RTrim(RemoveRepeatedSymbols(FileStr, Chr(32)))
 If Right(FileStr, 1) <> "-" Then GoTo end0
 If Asc(Right(FileStr, 2)) = 32 Then GoTo end0
 
   LineStr = Left(FileStr, Len(FileStr) - 1)
   HypFound = True
   
end0:
 RemoveHypsFromSingleLine = LineStr
End Function