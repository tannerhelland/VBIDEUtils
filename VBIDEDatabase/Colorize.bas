Attribute VB_Name = "Colorize_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 30/10/98
' * Time             : 14:47
' * Module Name      : Colorize_Module
' * Module Filename  : Colorize.bas
' **********************************************************************
' * Comments         : Colorize in black, blue, green the VB keywords
' *
' *
' **********************************************************************

Option Explicit

Private gsHeader        As String

Private gsVBBlackKeywords As String
Private gsVBBlueKeyWords As String
Private gsVBRedKeyWords As String
Private gsVBRemKeyWords As String

Private gsJSBlackKeywords As String
Private gsJSBlueKeyWords As String
Private gsJSRedKeyWords As String
Private gsJSRemKeyWords As String

Private gsJavaBlackKeywords As String
Private gsJavaBlueKeyWords As String
Private gsJavaRedKeyWords As String
Private gsJavaRemKeyWords As String

Public Function ColorizeVBCode(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : Colorize_Module
   ' * Module Filename  : Colorize.bas
   ' * Procedure Name   : ColorizeVBCode
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Colorize in black, blue, green the VB keywords
   ' *
   ' *
   ' **********************************************************************

   Dim sBuffer          As String
   Dim nI               As Long
   Dim nJ               As Long
   Dim sTmpWord         As String
   Dim nStartPos        As Long
   Dim nSelLen          As Long
   Dim nWordPos         As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim nStart           As Long
   Dim nLen             As Long
   Dim nLenBuffer       As Long

   Dim sOutput          As String
   Dim sLine            As String
   Dim bColor           As Boolean

   Dim sChar            As String
   Dim sOperator        As String
   Dim bOperator        As Boolean
   Dim sCharTmp         As String

   Dim bComment         As Boolean

   sBuffer = sText

   sOutput = ""
   sTmpWord = ""
   sLine = ""
   nLenBuffer = Len(sBuffer)
   For nI = 1 To nLenBuffer
      sChar = Mid$(sBuffer, nI, 1)

      Select Case sChar
         Case "A" To "Z", "a" To "z", "_":
            If sTmpWord = "" Then nStartPos = nI
            sTmpWord = sTmpWord & sChar

         Case "{", "}":
            sOutput = sOutput & "\" & Mid$(sBuffer, nI, 1)

         Case Chr$(34):
            ' *** Inside a string
            nSelLen = 1
            nStart = nI
            For nJ = 1 To nLenBuffer
               If (Mid$(sBuffer, nI + 1, 1) = Chr$(34)) Then
                  nSelLen = nSelLen - 1
                  nI = nI + 1
                  Exit For
               Else
                  nSelLen = nSelLen + 1
                  nI = nI + 1
               End If
            Next
            sLine = Mid$(sBuffer, nStart, nSelLen + 2)
            sLine = Replace(sLine, "\", "\\")
            sLine = Replace(sLine, "{", "\{")
            sLine = Replace(sLine, "}", "\}")
            sOutput = sOutput & sLine

         Case Chr$(39):
            ' *** Comment
COLOR_COMMENTS:
            bComment = True
            If Len(sTmpWord) > 0 Then GoTo COLOR_WORD

            If Len(sTmpWord) > 0 Then
               sOutput = sOutput & sTmpWord
               sTmpWord = ""
            End If

            nStart = nI - 1
            nSelLen = 0
            For nJ = 1 To nLenBuffer
               If Mid$(sBuffer, nI, 2) = vbCrLf Then
                  Exit For
               Else
                  nSelLen = nSelLen + 1
                  nI = nI + 1
               End If
            Next
            sLine = Mid$(sBuffer, nStart + 1, nSelLen)

            nLen = nSelLen
            ' *** Green
            sLine = Replace(sLine, "\", "\\")
            sLine = Replace(sLine, "{", "\{")
            sLine = Replace(sLine, "}", "\}")
            sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 "

            bComment = False

            '         Case "<", ">", "=", "-", "+", "/", "*", "&", "\":
            '            ' *** Operators
            'COLOR_OPERATOR:
            '            bOperator = True
            '            sOperator = sChar
            '            If Len(sTmpWord) > 0 Then GoTo COLOR_WORD
            '
            '            If Len(sTmpWord) > 0 Then
            '               sOutput = sOutput & sTmpWord
            '               sTmpWord = ""
            '            End If
            '
            '            If sOperator = "\" Then sOperator = "\\"
            '
            '            ' *** 2 chars for this operator
            '            If InStr(1, "=<>&|+-", Mid$(sBuffer, nI + 1)) Then
            '               sOperator = sOperator & Mid$(sBuffer, nI + 1)
            '               nI = nI + 1
            '            End If
            '
            '            nWordPos = InStr(1, gsJSRedKeyWords, "*" & sOperator & "*", 1)
            '            If nWordPos <> 0 Then
            '               ' *** Red
            '               sOutput = sOutput & "\plain\f2\fs20\cf3 " & sOperator & "\plain\f2\fs20\cf0 "
            '               bColor = True
            '            Else
            '               sOutput = sOutput & sOperator
            '            End If
            '            sOperator = ""
            '            bOperator = False

         Case Else
COLOR_WORD:
            If Not (Len(sTmpWord) = 0) Then
               If (bOperator) Or (bComment) Then
                  sCharTmp = ""
               Else
                  sCharTmp = sChar
               End If
               bColor = False
               nStart = nStartPos - 1
               nLen = Len(sTmpWord)
               'nWordPos = InStr(1, gsVBBlackKeywords, "*" & sTmpWord & "*", 1)
               'If nWordPos <> 0 Then
               '   ' *** Black
               '   sOutput = sOutput & "\plain\f2\fs20\cf1 " & sTmpWord & "\plain\f2\fs20\cf0 " & sChar
               '   bColor = True
               'End If
               nWordPos = InStr(1, gsVBBlueKeyWords, "*" & sTmpWord & "*", 1)
               If nWordPos <> 0 Then
                  ' *** Blue
                  If nI < nLenBuffer Then
                     sOutput = sOutput & "\plain\f2\fs20\cf1 " & sTmpWord & "\plain\f2\fs20\cf0 " & sCharTmp
                  Else
                     sOutput = sOutput & "\plain\f2\fs20\cf1 " & sTmpWord & "\plain\f2\fs20\cf0 "
                  End If
                  bColor = True
               End If
               If UCase$(sTmpWord) = gsJSRemKeyWords Then
                  nStart = nI - 4
                  nLen = 3
                  For nJ = 1 To nLenBuffer
                     If Mid$(sBuffer, nI, 2) = vbCrLf Then
                        Exit For
                     Else
                        nLen = nLen + 1
                        nI = nI + 1
                     End If
                  Next
                  ' *** Green
                  sLine = Mid$(sBuffer, nStart + 1, nLen)
                  sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 " & Mid$(sBuffer, nI, 1)
                  bColor = True
               End If
               'If bColor = False Then sOutput = sOutput & sTmpWord & Mid$(sBuffer, nI, 1)
               If bColor = False Then
                  If bOperator And Mid$(sBuffer, nI, 1) = sOperator Then
                     sOutput = sOutput & sTmpWord
                  Else
                     sOutput = sOutput & sTmpWord & Mid$(sBuffer, nI, 1)
                  End If
               End If

            Else
               If nI > 1 Then
                  If Mid$(sBuffer, nI - 1, 2) = vbCrLf Then
                     sOutput = sOutput & "\par "
                  Else
                     sOutput = sOutput & sChar
                  End If
               Else
                  sOutput = sOutput & sChar
               End If
            End If
            sTmpWord = ""
            '            If Len(sOperator) > 0 Then GoTo COLOR_OPERATOR
            If bComment Then GoTo COLOR_COMMENTS
      End Select
   Next
   If Len(sTmpWord) <> 0 Then GoTo COLOR_WORD

   nStart = 0

   ColorizeVBCode = gsHeader & sOutput & "}"

End Function

Public Function ColorizeJavaScriptCode(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : Colorize_Module
   ' * Module Filename  : Colorize.bas
   ' * Procedure Name   : ColorizeJavaScriptCode
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Colorize in black, blue, green the Javascript  keywords
   ' *
   ' *
   ' **********************************************************************

   Dim sBuffer          As String
   Dim nI               As Long
   Dim nJ               As Long
   Dim sTmpWord         As String
   Dim nStartPos        As Long
   Dim nSelLen          As Long
   Dim nWordPos         As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim nStart           As Long
   Dim nLen             As Long

   Dim sOutput          As String
   Dim sLine            As String
   Dim bColor           As Boolean

   Dim sChar            As String
   Dim sOperator        As String

   sBuffer = sText

   sOutput = ""
   sTmpWord = ""
   sLine = ""
   For nI = 1 To Len(sBuffer)
      sChar = Mid$(sBuffer, nI, 1)

      Select Case sChar
         Case "A" To "Z", "a" To "z", "_":
            If sTmpWord = "" Then nStartPos = nI
            sTmpWord = sTmpWord & sChar

         Case "{", "}":
            sOutput = sOutput & "\" & Mid$(sBuffer, nI, 1)

         Case Chr$(34):
            ' *** Inside a string
            nSelLen = 1
            nStart = nI
            For nJ = 1 To Len(sBuffer)
               If (Mid$(sBuffer, nI + 1, 1) = Chr$(34)) Then
                  nSelLen = nSelLen - 1
                  nI = nI + 1
                  Exit For
               Else
                  nSelLen = nSelLen + 1
                  nI = nI + 1
               End If
            Next
            sLine = Mid$(sBuffer, nStart, nSelLen + 2)
            sLine = Replace(sLine, "\", "\\")
            sOutput = sOutput & sLine

         Case "/":
            If sChar = "/" And Mid$(sBuffer, nI + 1, 1) <> "/" Then GoTo COLOR_OPERATOR

            nStart = nI - 1
            nSelLen = 0
            For nJ = 1 To Len(sBuffer)
               If Mid$(sBuffer, nI, 2) = vbCrLf Then
                  Exit For
               Else
                  nSelLen = nSelLen + 1
                  nI = nI + 1
               End If
            Next
            sLine = Mid$(sBuffer, nStart + 1, nSelLen)

            nLen = nSelLen
            ' *** Green
            sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 "

         Case "<", ">", "=", "!", "-", "+", "/", "*", "|", "&", "\":
            ' *** Operators
COLOR_OPERATOR:
            sOperator = sChar

            If Len(sTmpWord) > 0 Then
               sOutput = sOutput & sTmpWord
               sTmpWord = ""
            End If

            If sOperator = "\" Then sOperator = "\\"

            ' *** 2 chars for this operator
            If InStr(1, "=<>&|+-", Mid$(sBuffer, nI + 1, 1)) Then
               sOperator = sOperator & Mid$(sBuffer, nI + 1, 1)
               nI = nI + 1
            End If

            nWordPos = InStr(1, gsJSRedKeyWords, "*" & sOperator & "*", 1)
            If nWordPos <> 0 Then
               ' *** Red
               sOutput = sOutput & "\plain\f2\fs20\cf3 " & sOperator & "\plain\f2\fs20\cf0 "
               bColor = True
            Else
               sOutput = sOutput & sOperator
            End If

         Case Else
COLOR_WORD:
            If Not (Len(sTmpWord) = 0) Then
               bColor = False
               nStart = nStartPos - 1
               nLen = Len(sTmpWord)
               nWordPos = InStr(1, gsJSBlueKeyWords, "*" & sTmpWord & "*", 1)
               If nWordPos <> 0 Then
                  ' *** Blue
                  sOutput = sOutput & "\plain\f2\fs20\cf1 " & sTmpWord & "\plain\f2\fs20\cf0 " & sChar
                  bColor = True
               End If
               If UCase$(sTmpWord) = gsJSRemKeyWords Then
                  nStart = nI - 4
                  nLen = 3
                  For nJ = 1 To Len(sBuffer)
                     If Mid$(sBuffer, nI, 2) = vbCrLf Then
                        Exit For
                     Else
                        nLen = nLen + 1
                        nI = nI + 1
                     End If
                  Next
                  ' *** Green
                  sLine = Mid$(sBuffer, nStart + 1, nLen)
                  sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 " & Mid$(sBuffer, nI, 1)
                  bColor = True
               End If
               If bColor = False Then sOutput = sOutput & sTmpWord & Mid$(sBuffer, nI, 1)
            Else
               If nI > 1 Then
                  If Mid$(sBuffer, nI - 1, 2) = vbCrLf Then
                     sOutput = sOutput & vbCrLf & "\par "
                  Else
                     sOutput = sOutput & sChar
                  End If
               Else
                  sOutput = sOutput & sChar
               End If
            End If
            sTmpWord = ""
      End Select
   Next
   If Len(sTmpWord) <> 0 Then GoTo COLOR_WORD

   nStart = 0

   ColorizeJavaScriptCode = gsHeader & sOutput & "\par }"

End Function

Public Function ColorizeJavaCode(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : Colorize_Module
   ' * Module Filename  : Colorize.bas
   ' * Procedure Name   : ColorizeJavaCode
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Colorize in black, blue, green the Javascript  keywords
   ' *
   ' *
   ' **********************************************************************

   Dim sBuffer          As String
   Dim nI               As Long
   Dim nJ               As Long
   Dim sTmpWord         As String
   Dim nStartPos        As Long
   Dim nSelLen          As Long
   Dim nWordPos         As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim nStart           As Long
   Dim nLen             As Long

   Dim sOutput          As String
   Dim sLine            As String
   Dim bColor           As Boolean

   Dim sChar            As String
   Dim sCharTmp         As String
   Dim sOperator        As String
   Dim bOperator        As Boolean

   Dim bMultipleRem     As Boolean

   sBuffer = sText

   bMultipleRem = False
   sOutput = ""
   sTmpWord = ""
   sLine = ""
   For nI = 1 To Len(sBuffer)
      sChar = Mid$(sBuffer, nI, 1)

      Select Case sChar
         Case "A" To "Z", "a" To "z", "_":
            If sTmpWord = "" Then nStartPos = nI
            sTmpWord = sTmpWord & sChar

         Case Chr$(34):
            ' *** Inside a string
            nSelLen = 1
            nStart = nI
            For nJ = 1 To Len(sBuffer)
               If (Mid$(sBuffer, nI + 1, 1) = Chr$(34)) Then
                  nSelLen = nSelLen - 1
                  nI = nI + 1
                  Exit For
               Else
                  nSelLen = nSelLen + 1
                  nI = nI + 1
               End If
            Next
            sLine = Mid$(sBuffer, nStart, nSelLen + 2)
            sLine = Replace(sLine, "\", "\\")
            sOutput = sOutput & sLine

         Case "{", "}":
            sOutput = sOutput & "\" & Mid$(sBuffer, nI, 1)

         Case "/":
            If sChar = "/" And Mid$(sBuffer, nI + 1, 1) = "*" Then
               bMultipleRem = True
            ElseIf sChar = "/" And Mid$(sBuffer, nI + 1, 1) <> "/" Then
               GoTo COLOR_OPERATOR
            End If

            If bMultipleRem = False Then
               nStart = nI - 1
               nSelLen = 0
               For nJ = 1 To Len(sBuffer)
                  If Mid$(sBuffer, nI, 2) = vbCrLf Then
                     Exit For
                  Else
                     nSelLen = nSelLen + 1
                     nI = nI + 1
                  End If
               Next
               sLine = Mid$(sBuffer, nStart + 1, nSelLen)
            Else
               nSelLen = InStr(nI, sBuffer, "*/")
               If nSelLen = 0 Then
                  nSelLen = nI
                  sLine = Mid$(sBuffer, nI, Len(sBuffer) - nI)
               Else
                  nSelLen = nSelLen + 2
                  sLine = Mid$(sBuffer, nI, nSelLen - nI)
                  sLine = Replace(sLine, vbCrLf, "\par ")
               End If
               nI = nSelLen - 1
            End If

            nLen = nSelLen
            ' *** Green
            sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 "

         Case "+", "-", ">", "<", "=", "/", "\":
            ' *** Operators
COLOR_OPERATOR:
            bOperator = True
            sOperator = sChar
            If Len(sTmpWord) > 0 Then GoTo COLOR_WORD

            ' *** 2 chars for this operator
            If InStr(1, "=<>&|+-", Mid$(sBuffer, nI + 1, 1)) Then
               sOperator = sOperator & Mid$(sBuffer, nI + 1, 1)
               nI = nI + 1
            End If

            If sOperator = "\" Then sOperator = "\\"

            nWordPos = InStr(1, gsJSRedKeyWords, "*" & sOperator & "*", 1)
            If nWordPos <> 0 Then
               ' *** Red
               sOutput = sOutput & "\plain\f2\fs20\cf3 " & sOperator & "\plain\f2\fs20\cf0 "
               bColor = True
            Else
               sOutput = sOutput & sOperator
            End If
            sOperator = ""
            bOperator = False

         Case Else
COLOR_WORD:
            If Len(sTmpWord) > 0 Then
               If bOperator = True Then
                  sCharTmp = ""
               Else
                  sCharTmp = sChar
               End If
               bColor = False
               nStart = nStartPos - 1
               nLen = Len(sTmpWord)
               nWordPos = InStr(1, gsJSBlueKeyWords, "*" & sTmpWord & "*", 1)
               If nWordPos <> 0 Then
                  ' *** Blue
                  sOutput = sOutput & "\plain\f2\fs20\cf1 " & sTmpWord & "\plain\f2\fs20\cf0 " & sCharTmp
                  bColor = True
               End If
               If UCase$(sTmpWord) = gsJSRemKeyWords Then
                  nStart = nI - 4
                  nLen = 3
                  For nJ = 1 To Len(sBuffer)
                     If Mid$(sBuffer, nI, 2) = vbCrLf Then
                        Exit For
                     Else
                        nLen = nLen + 1
                        nI = nI + 1
                     End If
                  Next
                  ' *** Green
                  sLine = Mid$(sBuffer, nStart + 1, nLen)
                  sOutput = sOutput & "\plain\f2\fs20\cf2 " & sLine & "\plain\f2\fs20\cf0 " & Mid$(sBuffer, nI, 1)
                  bColor = True
               End If
               If bColor = False Then
                  If bOperator And Mid$(sBuffer, nI, 1) = sOperator Then
                     sOutput = sOutput & sTmpWord
                  Else
                     sOutput = sOutput & sTmpWord & Mid$(sBuffer, nI, 1)
                  End If
               End If
            Else
               If nI > 1 Then
                  If Mid$(sBuffer, nI - 1, 2) = vbCrLf Then
                     sOutput = sOutput & vbCrLf & "\par "
                  Else
                     sOutput = sOutput & sChar
                  End If
               Else
                  sOutput = sOutput & sChar
               End If
            End If
            sTmpWord = ""
            If Len(sOperator) > 0 Then GoTo COLOR_OPERATOR
      End Select
   Next
   If Len(sTmpWord) <> 0 Then GoTo COLOR_WORD

   nStart = 0

   ColorizeJavaCode = gsHeader & sOutput & "\par }"

End Function

Public Sub InitColorize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : Colorize_Module
   ' * Module Filename  : Colorize.bas
   ' * Procedure Name   : InitColorize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Initialize the VB keywords
   ' *
   ' *
   ' **********************************************************************

   ' *** Header of RTF file
   gsHeader = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fmodern Courier New;}{\f3\fmodern Courier New;}}" & vbCrLf
   gsHeader = gsHeader & "{\colortbl\red0\green0\blue0;\red0\green0\blue127;\red0\green127\blue0;\red127\green0\blue0;}" & vbCrLf
   gsHeader = gsHeader & "\deflang2060\pard\plain\f2\fs20\cf0 " & vbCrLf

   'gsVBBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
   gsVBBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Optional*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"
   gsVBRedKeyWords = "*=*>*<*<=*>=*=<*=>*+*-*/***<>*&*"
   gsVBRemKeyWords = "REM"

   'gsJSBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
   gsJSBlueKeyWords = "*abstract*boolean*break*byte*case*catch*char*class*const*continue*default*delete*do*double*else*extends*false*final*finally*float*for*function*goto*if*implements*import*in*instanceof*int*interface*long*native*new*null*package*private*protected*public*return*short*static*super*switch*synchronized*this*throw*throws*transient*true*try*typeof*var*void*while*withAnchoranchors*Applet*applets*Area*Array*Button*Checkbox*Date*document*FileUpload*Form*forms*Frame*frames*Hidden*history*Image*images*Link*links*Area*location*Math*MimeType*mimeTypes*navigator*options*Password*Plugin*plugins*Radio*Reset*Select*String*Submit*Text*Textarea*window*"
   gsJSRedKeyWords = "*<*>*=*!*==*!=*=<*>=*<=*=>*+=*-=*-*+*/***&&*||*++*"
   gsJSRemKeyWords = "//"

   'gsJavaBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
   gsJavaBlueKeyWords = "*abstract*break*byte*boolean*catch*case*class*char*continue*default*double*do*else*extends*false*final*float*for*finally*if*import*implements*int*interface*instanceof*long*length*native*new*null*package*private*protected*public*final*return*switch*synchronizedshort*static*super*try*true*this*throw*throws*threadsafe*transient*void*while*"
   gsJavaRedKeyWords = "*+*-*>*<*<=*>=*=>*=<*<>*==*=*/*"
   gsJavaRemKeyWords = "//"

End Sub

Public Sub RemoveUnderScores(rtf As Control)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:48
   ' * Module Name      : Colorize_Module
   ' * Module Filename  : Colorize.bas
   ' * Procedure Name   : RemoveUnderScores
   ' * Parameters       :
   ' *                    rtf As Control
   ' **********************************************************************
   ' * Comments         : Remove all the "_" in VB Code
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Long
   Dim noldPos          As Long

   noldPos = rtf.SelStart

   ' *** Remove any underscores followed by a vbCrLf
   nPos = InStr(1, rtf.Text, "_" & vbCrLf)
   Do While Not (nPos = 0)
      rtf.SelStart = nPos - 1
      rtf.SelLength = Len("_" & vbCrLf)
      rtf.SelText = ""
      nPos = InStr(nPos, rtf.Text, "_" & vbCrLf)
   Loop

   ' *** Remove any underscores followed by a space and a vbCrLf
   nPos = InStr(1, rtf.Text, "_ " & vbCrLf)
   Do While Not (nPos = 0)
      rtf.SelStart = nPos - 1
      rtf.SelLength = Len("_ " & vbCrLf)
      rtf.SelText = ""
      nPos = InStr(nPos, rtf.Text, "_ " & vbCrLf)
   Loop

   rtf.SelStart = noldPos

End Sub
