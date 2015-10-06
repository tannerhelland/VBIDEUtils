VERSION 5.00
Begin VB.Form frmDocHelp 
   Caption         =   "ActiveX Documenter"
   ClientHeight    =   6735
   ClientLeft      =   4305
   ClientTop       =   1830
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fDev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   7620
   Begin VB.TextBox lblStatus 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   " Ready."
      Top             =   6420
      Width           =   5835
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   5880
      ScaleHeight     =   255
      ScaleWidth      =   1695
      TabIndex        =   20
      Top             =   6420
      Width           =   1755
   End
   Begin VB.PictureBox picTab 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Index           =   0
      Left            =   180
      ScaleHeight     =   4095
      ScaleWidth      =   7215
      TabIndex        =   1
      Top             =   1080
      Width           =   7275
      Begin VB.ListBox lstMembers 
         Height          =   2595
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   1320
         Width           =   7155
      End
      Begin VB.ListBox lstGeneral 
         Height          =   840
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   240
         Width           =   7155
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   2595
      End
      Begin VB.Label lblMembers 
         Caption         =   "Members:"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   1080
         Width           =   2595
      End
   End
   Begin VB.ComboBox cboClass 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6180
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   780
      Width           =   1395
   End
   Begin VB.PictureBox picTab 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Index           =   2
      Left            =   300
      ScaleHeight     =   4095
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   1740
      Width           =   7275
      Begin TLibHelp.VBRichEdit rtfDocument 
         Height          =   3795
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   6694
      End
   End
   Begin TLibHelp.cToolbar tbrMain 
      Left            =   720
      Top             =   540
      _ExtentX        =   9234
      _ExtentY        =   873
   End
   Begin TLibHelp.TabControl tabMain 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TLibHelp.cReBar rbrMain 
      Left            =   60
      Top             =   60
      _ExtentX        =   10398
      _ExtentY        =   767
   End
   Begin VB.PictureBox picTab 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Index           =   1
      Left            =   60
      ScaleHeight     =   4095
      ScaleWidth      =   7215
      TabIndex        =   2
      Top             =   840
      Width           =   7275
      Begin VB.CommandButton cmdRemoveSuperClass 
         Caption         =   "<"
         Height          =   315
         Left            =   3660
         TabIndex        =   19
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdAddSuperClass 
         Caption         =   ">"
         Height          =   315
         Left            =   3660
         TabIndex        =   18
         Top             =   2700
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdRemoveInterface 
         Caption         =   "<"
         Height          =   315
         Left            =   3660
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdAddInterface 
         Caption         =   ">"
         Height          =   315
         Left            =   3660
         TabIndex        =   16
         Top             =   660
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ListBox lstSuperClass 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1815
         IntegralHeight  =   0   'False
         Left            =   4020
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ListBox lstInterface 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1815
         IntegralHeight  =   0   'False
         Left            =   4020
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ListBox lstAvailable 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   3765
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label lblSuper 
         Caption         =   "Create Super Classed Member:"
         Height          =   255
         Left            =   4020
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblCreate 
         Caption         =   "Create Empty, Equivalent Member:"
         Height          =   255
         Left            =   4020
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblAvail 
         Caption         =   "Available:"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save..."
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print..."
         Index           =   4
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print Preview..."
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   12
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Select All"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Invert Selection"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Clear Selection"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents..."
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "vbaccelerator on the &Web..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmDocHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum EDocumentationTypes
    eDocumentText
    eDocumentRtf
    eDocumentHTML
End Enum
Private Enum EDocumentSectionTypes
    eHeader
    eGeneralInformation
    eBody
    eFooter
End Enum

Private m_cTLI As TypeLibInfo

Private m_sInterfaces() As String
Private m_iTypeInfoIndex() As Long
Private m_iBelongsToInterface() As Long
Private m_bHidden() As Boolean
Private m_iCount As Long

Private m_bShowHelp As Boolean
Private m_sCmdLine As String

Private m_cMRU As cMRUFileList
Private m_sInitSaveDir As String
Private m_sInitOpenDir As String
Private m_lMax As Long
Private m_lValue As Long

Private Sub Status(ByVal sStatus As String)
    lblStatus.Text = " " & sStatus
    Me.Refresh
End Sub
Private Property Let ProgressMax(ByVal lMax As Long)
    m_lMax = lMax
End Property
Private Property Let ProgressValue(ByVal lValue As Long)
Dim hBrush As Long
Dim lColor As Long
Dim tR As RECT
Dim lRIght As Long

    ' Clear progress bar:
    tR.Right = picProgress.ScaleWidth \ Screen.TwipsPerPixelX
    tR.Bottom = picProgress.ScaleHeight \ Screen.TwipsPerPixelY
    OleTranslateColor vbButtonFace, 0, lColor
    hBrush = CreateSolidBrush(lColor)
    FillRect picProgress.hdc, tR, hBrush
    DeleteObject hBrush
    If (lValue > 0) Then
        If (lValue > m_lMax) Then
            lValue = m_lMax
        End If
        m_lValue = lValue
        lRIght = ((tR.Right - tR.Left) * lValue) \ m_lMax
        tR.Right = lRIght
        OleTranslateColor RGB(0, 0, &HC0), 0, lColor
        hBrush = CreateSolidBrush(lColor)
        FillRect picProgress.hdc, tR, hBrush
        DeleteObject hBrush
    End If
    picProgress.Refresh
    
End Property

Private Sub pParseCommand()
Dim sRem As String
Dim iPos As Long
Dim bPrint As Boolean

    ' do we have a print?
    iPos = InStr(m_sCmdLine, "/p")
    If (iPos <> 0) Then
        bPrint = True
        If (iPos > 1) Then
            sRem = Left$(m_sCmdLine, (iPos - 1))
        End If
        If (iPos < Len(m_sCmdLine) - 2) Then
            sRem = sRem & Mid$(m_sCmdLine, (iPos + 3))
        End If
        m_sCmdLine = sRem
    End If
    
    ' The remainder should be interpreted as a file:
    
    
End Sub

Public Property Let CommandLine(ByVal sCmd As String)
    m_sCmdLine = sCmd
End Property

Private Function pbGetTypeLibInfo( _
        ByVal sFIle As String _
    ) As Boolean
On Error GoTo pGetTypeLibInfoError
    
    ' Clear up info we're holding about previous TypeLib, if any:
    m_iCount = 0
    Erase m_sInterfaces
    Erase m_iTypeInfoIndex
    Erase m_iBelongsToInterface
    cboClass.Clear
    cboClass.AddItem "<No Interfaces>"
    cboClass.ListIndex = 0
    cboClass.Enabled = False
    lstGeneral.Clear
    lstMembers.Clear
    
    ' Generate a TypeLibInfo object for the specified file.
    Status "Linking to Type Library..."
    Set m_cTLI = TLI.TypeLibInfoFromFile(sFIle)
    
    Me.Caption = App.Title & " (" & sFIle & ")"
    ' If we succeed, then organize the TypeInfo members.
    ' VB classes have a number of components which are normally hidden from you:
        ' -the CoClass, which has the correct name but is empty because all its functions
        '   are performed by the members with _ before the name,
        ' -one or two DispInterface items, which underscores first.  The first has one underscore
        '   and contains the non-event interfaces.  The second has two and contains the events.
    Dim iTypeInfo As Long
    Dim sName As String
    Dim sBelongsTo As String
    Dim iCheckOwner As Long
    
    ' Populate general information:
    With m_cTLI
        lstGeneral.AddItem "Library:" & vbTab & .Name & " (" & .HelpString & ")"
        lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        lstGeneral.AddItem "File:" & vbTab & sFIle
        lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        lstGeneral.AddItem "GUID:" & vbTab & .Guid
        lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        lstGeneral.AddItem "Version:" & vbTab & .MajorVersion & "." & .MinorVersion
        lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
    End With
    
    Status "Counting Type Library Members..."
    With m_cTLI
        ' Items with an attribute mask = 16 are old interfaces:
        m_iCount = .TypeInfoCount
        ReDim Preserve m_sInterfaces(1 To m_iCount) As String
        ReDim Preserve m_iTypeInfoIndex(1 To m_iCount) As Long
        ReDim Preserve m_iBelongsToInterface(1 To m_iCount) As Long
        ReDim Preserve m_bHidden(1 To m_iCount) As Boolean
        For iTypeInfo = 1 To m_iCount
            Debug.Print .TypeInfos(iTypeInfo).Name, .TypeInfos(iTypeInfo).AttributeMask
            m_sInterfaces(iTypeInfo) = .TypeInfos(iTypeInfo).Name
            m_iTypeInfoIndex(iTypeInfo) = iTypeInfo
            m_bHidden(iTypeInfo) = (.TypeInfos(iTypeInfo).AttributeMask = 16)
        Next iTypeInfo
    End With
    
    Status "Checking for Related VB Type Libraries and Parsing..."
    For iTypeInfo = 1 To m_iCount
        If Not (m_bHidden(iTypeInfo)) Then
            If (Left$(m_sInterfaces(iTypeInfo), 1) = "_") Then
                sBelongsTo = Mid$(m_sInterfaces(iTypeInfo), 2)
                If (Left$(sBelongsTo, 1) = "_") Then
                    sBelongsTo = Mid$(sBelongsTo, 2)
                End If
                For iCheckOwner = 1 To m_iCount
                    If (iCheckOwner <> iTypeInfo) Then
                        If Not (m_bHidden(iCheckOwner)) Then
                            If (m_sInterfaces(iCheckOwner) = sBelongsTo) Then
                                m_iBelongsToInterface(iTypeInfo) = iCheckOwner
                                Exit For
                            End If
                        End If
                    End If
                Next iCheckOwner
            End If
        End If
    Next iTypeInfo
    
    ' Add to the combo box:
    If (iTypeInfo > 0) Then
        Status "Adding Type Library Members..."
        cboClass.Clear
        cboClass.AddItem "<All Interfaces>"
        cboClass.ItemData(cboClass.NewIndex) = &HFFFFFFF
        For iTypeInfo = 1 To m_iCount
            If Not (m_bHidden(iTypeInfo)) Then
                If (m_iBelongsToInterface(iTypeInfo) = 0) Then
                    If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString = "enum") Then
                        cboClass.AddItem "Enum " & m_sInterfaces(iTypeInfo)
                    Else
                        cboClass.AddItem m_sInterfaces(iTypeInfo)
                    End If
                    cboClass.ItemData(cboClass.NewIndex) = iTypeInfo
                End If
            End If
        Next iTypeInfo
        cboClass.Enabled = True
        cboClass.ListIndex = 0
    End If
            
    Status "Ready."
    
    pbGetTypeLibInfo = True
    Exit Function
pGetTypeLibInfoError:
    MsgBox "Failed to get type lib info for file: '" & sFIle & "'" & vbCrLf & vbCrLf & Err.Description, vbExclamation
    Set m_cTLI = Nothing
    Exit Function
End Function
Private Function psGetGeneralInfoRtf(ByVal lIndex As Long) As String
Dim sPre As String
    sPre = lstGeneral.List(lIndex)
    ' parse backslashes into Rtf version:
    ReplaceSection sPre, "\", "\\"
    ' parse tabs into Rtf version:
    ReplaceSection sPre, Chr$(9), "\tab "
    ' parse { brackets:
    pParseGUID sPre

    psGetGeneralInfoRtf = sPre
End Function
Private Sub pParseGUID(ByRef sThis As String)
    ReplaceSection sThis, "{", "\{"
    ReplaceSection sThis, "}", "\}"
End Sub

Private Sub ReplaceSection( _
        ByRef sToModify As String, _
        ByVal sToReplace As String, _
        ByVal sReplaceWith As String _
    )
' ==================================================================
' Replaces all occurrences of sToReplace with
' sReplaceWidth in sToModify.
' ==================================================================
' ==================================================================
' Replaces all occurrences of sToReplace with
' sReplaceWidth in sToModify.
' ==================================================================
Dim iLastPos As Long
Dim iNextPos As Long
Dim iReplaceLen As Long
Dim sOut As String
    iReplaceLen = Len(sToReplace)
    iLastPos = 1
    iNextPos = InStr(iLastPos, sToModify, sToReplace)
    sOut = ""
    Do While (iNextPos > 0)
        If (iNextPos > 1) Then
            sOut = sOut & Mid$(sToModify, iLastPos, (iNextPos - iLastPos))
        End If
        sOut = sOut & sReplaceWith
        iLastPos = iNextPos + iReplaceLen
        iNextPos = InStr(iLastPos, sToModify, sToReplace)
    Loop
    If (iLastPos <= Len(sToModify)) Then
        sOut = sOut & Mid$(sToModify, (iLastPos))
    End If
    sToModify = sOut
End Sub
Private Sub pCreateHelpFile( _
        ByVal sFIle As String _
    )
Dim iTypeInfo As Long
Dim sName As String
Dim sGUID As String
Dim sHelpString As String
Dim sType As String
Dim sMembers() As String
Dim sHelp() As String
Dim iMemberCount As Long
Dim bDone() As Boolean
Dim sJunk As String
Dim sEvents() As String
Dim sEventHelp() As String
Dim iEventCount As Long
Dim iBelongsTo As Long


    With m_cTLI
        Debug.Print "Library:" & vbTab & .Name
        Debug.Print "Filename:" & vbTab & sFIle
        Debug.Print "GUID:" & vbTab & .Guid
        Debug.Print "Version:" & vbTab & .MajorVersion & "." & .MinorVersion
        Debug.Print ""
        Debug.Print "Members:"
        Debug.Print
                
        ' Now document:
        For iTypeInfo = 1 To .TypeInfoCount
            If (m_iBelongsToInterface(iTypeInfo) = 0) Then
            
                Erase sMembers: Erase sHelp: iMemberCount = 0
                Erase sEvents: Erase sEventHelp: iEventCount = 0
                
                pEvaluateMember .TypeInfos(iTypeInfo), sName, sGUID, sHelpString, sType, sMembers(), sHelp(), iMemberCount
                
                For iBelongsTo = 1 To .TypeInfoCount
                    If (m_iBelongsToInterface(iBelongsTo) = iTypeInfo) Then
                        If (Left$(m_sInterfaces(iBelongsTo), 2) = "__") Then
                            ' events:
                            pEvaluateMember .TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sEvents(), sEventHelp(), iEventCount
                        Else
                            ' methods/properties:
                            pEvaluateMember .TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sMembers(), sHelp(), iMemberCount
                        End If
                    End If
                Next iBelongsTo
                
                pOutputMember iTypeInfo, sName, sGUID, sHelpString, sType, sMembers(), sHelp(), iMemberCount, sEvents(), sEventHelp(), iEventCount
            End If
        Next iTypeInfo
    End With
    
End Sub

Private Sub pOutputMember( _
        ByVal iIndex As Long, _
        ByVal sName As String, _
        ByVal sGUID As String, _
        ByVal sHelpString As String, _
        ByVal sType As String, _
        ByRef sMembers() As String, _
        ByRef sHelp() As String, _
        ByVal iMemberCount As Long, _
        ByRef sEvents() As String, _
        ByRef sEventHelp() As String, _
        ByVal iEventCount As Long _
    )
Dim iM As Long
    If (sType = "enum") Then
        Debug.Print "Public Enum " & sName
        For iM = 1 To iMemberCount
            Debug.Print vbTab & sMembers(iM)
        Next iM
        Debug.Print "End Enum"
    Else
        Debug.Print "Class:" & sName & " (" & sGUID & ")" & vbCrLf & "Methods:" 'vbCrLf & vbCrLf & "Type:" & sType
        For iM = 1 To iMemberCount
            Debug.Print sMembers(iM), sHelp(iM)
        Next iM
        If (iEventCount > 0) Then
            Debug.Print "Events:"
            For iM = 1 To iEventCount
                Debug.Print sEvents(iM), sEventHelp(iM)
            Next iM
        End If
    End If
    Debug.Print

End Sub
    
Private Sub pEvaluateEnum( _
        ByRef ti As TypeInfo, _
        ByRef sMembers() As String, _
        ByRef iMemberCount As Long _
    )
Dim iMember As Long

    iMemberCount = 0
    Erase sMembers
    
    With ti
        On Error Resume Next
        iMemberCount = .Members.Count
        If (Err.Number <> 0) Then
            iMemberCount = 0
        End If
        Err.Clear
        
        On Error GoTo 0
        If (iMemberCount > 0) Then
            ReDim sMembers(1 To iMemberCount) As String
            For iMember = 1 To iMemberCount
                With .Members(iMember)
                    sMembers(iMember) = .Name & "=" & .Value
                End With
            Next iMember
        End If
    End With
    
End Sub
Private Sub pEvaluateClass( _
        ByRef ti As TypeInfo, _
        ByRef sMembers() As String, _
        ByRef sHelp() As String, _
        ByRef iMemberCount As Long _
    )
Dim iMember As Long
Dim iMemCount As Long
        
    ' Initialise:
    iMemberCount = 0
    Erase sMembers
    Erase sHelp
    
    ' Find out the contents of the TypeInfo:
    With ti
        
        ' Get number of members in this class:
        On Error Resume Next
        iMemCount = .Members.Count
        If (Err.Number <> 0) Then
            iMemCount = 0
        End If
        Err.Clear
        
        On Error GoTo 0
        If (iMemCount > 0) Then
        
            For iMember = 1 To iMemCount
                If (.Members(iMember).AttributeMask = 0) Then ' Not hidden
                    iMemberCount = iMemberCount + 1
                    ReDim Preserve sMembers(1 To iMemberCount) As String
                    ReDim Preserve sHelp(1 To iMemberCount) As String
                    pEvaluateClassMember .Members(iMember), sMembers(iMemberCount), sHelp(iMemberCount), iMemberCount
                End If
            Next iMember
        End If
        
    End With
End Sub

Private Function psGetMemberType( _
        ByRef tM As MemberInfo, _
        ByRef bIsLet As Boolean _
    ) As String
    
    bIsLet = False
    
    Select Case tM.InvokeKind
    Case INVOKE_EVENTFUNC
        psGetMemberType = "Event"
    Case INVOKE_FUNC
        If (tM.ReturnType.VarType = VT_VOID) Then
            psGetMemberType = "Sub"
        Else
            psGetMemberType = "Function"
        End If
    Case INVOKE_PROPERTYGET
        psGetMemberType = "Property Get"
    Case INVOKE_PROPERTYPUT
        psGetMemberType = "Property Let"
        bIsLet = True
    Case INVOKE_PROPERTYPUTREF
        psGetMemberType = "Property Set"
    Case INVOKE_UNKNOWN
        psGetMemberType = "Const"
    Case Else
        Debug.Assert 1 = 0
    End Select
    
End Function

Private Sub pEvaluateClassMember( _
        ByRef tM As MemberInfo, _
        ByRef sMember As String, _
        ByRef sHelp As String, _
        ByRef iMember As Long _
    )
Dim iParam As Long
Dim iParamCount As Long
Dim lType As TliVarType
Dim bOptional As Boolean
Dim bIsLet As Boolean
Dim sDefault As String
Dim sName As String
Dim sPrefix As String

'On Error Resume Next

    With tM
        
        ' Type of member (sub, function, property..):
        sMember = psGetMemberType(tM, bIsLet)
        sName = .Name
        If (Left$(sName, 1) = "_") Then
            ' check for standard prefixes:
            sPrefix = Left$(sName, 7)
            If (sPrefix = "_B_var_") Then
                sName = Mid$(sName, 8)
            ElseIf (sPrefix = "_B_str_") Then
                sName = Mid$(sName, 8) & "$"
            End If
        End If
        
        sMember = sMember & " " & sName
        
        ' Any parameters?
        iParamCount = .Parameters.Count
        If (Err.Number <> 0) Then
            iParamCount = 0
        End If
        Err.Clear
        
        ' If we have parameters then add the function description:
        For iParam = 1 To iParamCount
            
            bOptional = False
            
            With .Parameters(iParam)
                ' Add open bracket first time:
                If (iParam = 1) Then
                    sMember = sMember & "("
                End If
                
                ' .HasCustomData or .Optional implies the parameter is optional:
                If (.HasCustomData() = True) Then
                   sMember = sMember & "Optional "
                   bOptional = True
                Else
                    If .Optional Then
                        sMember = sMember & "Optional "
                    End If
                End If
                
                ' Check Byref/Byval status of member:
                If ((lType And VT_BYREF) = VT_BYREF) Then
                Else
                    sMember = sMember & "ByVal "
                End If
                
                ' Name of parameter:
                sMember = sMember & .Name
                
                ' Evaluate the parameter type:
                If (.VarTypeInfo.VarType = 0) Then
                    ' Custom type:
                    sMember = sMember & " As " & .VarTypeInfo.TypeInfo.Name
                Else
                    lType = .VarTypeInfo.VarType
                    sMember = sMember & psTranslateType(lType)
                End If
                
                ' Add default value if there is one:
                If (bOptional) Then
                   If (.Default) Then
                    On Error Resume Next
                    sDefault = CStr(.DefaultValue)
                    If (Err.Number = 0) Then
                        sMember = sMember & "=" & sDefault
                    Else
                        sMember = sMember & "=Nothing"
                    End If
                    Err.Clear
                    On Error GoTo 0
                   End If
                End If
                
                ' If this is the last parameter then close the declaration,
                ' otherwise put a comma in front of the next one:
                If (iParam < iParamCount) Then
                    sMember = sMember & ", "
                Else
                    If Not (bIsLet) Then
                        sMember = sMember & ")"
                    End If
                End If
                
            End With
        Next iParam
                    
        ' Now add the return type and fix up Property Lets as required:
        If (.ReturnType.VarType <> 0) Then
            ' Returns a Standard type:
            If (.ReturnType.VarType = VT_VOID) Then
                ' sub
            Else
                ' If a constant, we want to get the constant value:
                If (Left$(sMember, 5) = "Const") Then
                    Debug.Print sMember, .ReturnType.VarType
                    If (.ReturnType.VarType = VT_BSTR) Or (.ReturnType.VarType = VT_LPSTR) Then
                        sMember = sMember & " = " & psParseForNonPrintable(.Value) & ""
                    Else
                        sMember = sMember & " = " & .Value
                    End If
                Else
                    ' If property let, must put in the RHS argument:
                    If (bIsLet) Then
                        If (iParamCount = 0) Then
                            ' property let has only one var:
                            sMember = sMember & "(RHS "
                        Else
                            ' more than on property let var:
                            sMember = sMember & ", RHS"
                        End If
                    End If
                    
                    ' No paramters, put the open close in:
                    If (iParamCount = 0) Then
                        If Not (bIsLet) Then
                            sMember = sMember & "()"
                        End If
                    End If
                    
                    ' Add the return type:
                    sMember = sMember & psTranslateType(.ReturnType.VarType)
                    
                    ' Close the property let statement:
                    If (bIsLet) Then
                        sMember = sMember & ")"
                    End If
                End If
            End If
        Else
            ' If property let, must put in the RHS argument:
            If (bIsLet) Then
                If (iParamCount = 0) Then
                    ' property let has only one var:
                    sMember = sMember & "(RHS "
                Else
                    ' more than on property let var:
                    sMember = sMember & ", RHS"
                End If
            End If
            
            ' No paramters, put the open close in:
            If (iParamCount = 0) Then
                If Not (bIsLet) Then
                    sMember = sMember & "()"
                End If
            End If
            
            ' Returns a custom type:
            sMember = sMember & " As " & .ReturnType.TypeInfo.Name
            
            ' Close the property let statement:
            If (bIsLet) Then
                sMember = sMember & ")"
            End If
        End If
        
        sHelp = .HelpString
    
    End With
    
End Sub
Private Function psParseForNonPrintable(ByVal vThis As Variant) As String
Dim iPos As Long
Dim sRet As String
Dim iLen As Long
Dim iChar As Integer
Dim sChar As String
Dim bLastNonPrintable As Boolean

    iLen = Len(vThis)
    For iPos = 1 To iLen
        sChar = Mid$(vThis, iPos, 1)
        iChar = Asc(sChar)
        If (iChar < 32) Then
            If (iPos <> 1) Then
                sRet = sRet & "& "
            End If
            If (bLastNonPrintable) Then
                sChar = "Chr$(" & iChar & ") "
            Else
                If (iPos = 1) Then
                    sChar = "Chr$(" & iChar & ") "
                Else
                    sChar = """ & Chr$(" & iChar & ") "
                End If
            End If
            bLastNonPrintable = True
        Else
            If (bLastNonPrintable) Or (iPos = 1) Then
                If (iPos <> 1) Then
                    sRet = sRet & "& "
                End If
                sChar = """" & sChar
            End If
            bLastNonPrintable = False
        End If
        sRet = sRet & sChar
    Next iPos
    If Not (bLastNonPrintable) Then
        sRet = sRet & """"
    End If
    psParseForNonPrintable = sRet

End Function

Private Sub pEvaluateMember( _
        ByRef ti As TypeInfo, _
        ByRef sName As String, _
        ByRef sGUID As String, _
        ByRef sHelpString As String, _
        ByRef sType As String, _
        ByRef sMembers() As String, _
        ByRef sHelp() As String, _
        ByRef iMemberCount As Long _
    )
Dim iTypeInfo As Long

    With ti
        sName = .Name
        sGUID = .Guid
        sHelpString = .HelpString
        sType = .TypeKindString
        
        If (.TypeKind = TKIND_ENUM) Then
            ' do enum:
            pEvaluateEnum ti, sMembers(), iMemberCount
        Else
            ' do class:
            pEvaluateClass ti, sMembers(), sHelp(), iMemberCount
        End If
    End With
    
End Sub
Private Function psTranslateType(ByVal lType As Long)
Dim sType As String
    Select Case (lType And &HFF&)
    Case VT_BOOL
        sType = "Boolean"
    Case VT_BSTR, VT_LPSTR
        sType = "String"
    Case VT_DATE
        sType = "Date"
    Case VT_INT
        sType = "Integer"
    Case VT_VARIANT
        sType = "Variant"
    Case VT_DECIMAL
        sType = "Decimal"
    Case VT_I4
        sType = "Long"
    Case VT_I2
        sType = "Integer"
    Case VT_I8
        sType = "Unknown"
    Case VT_SAFEARRAY
        sType = "SafeArray"
    Case VT_CLSID
        sType = "CLSID"
    Case VT_UINT
        sType = "UInt"
    Case VT_UI4
        sType = "ULong"
    Case VT_UNKNOWN
        sType = "Unknown"
    Case VT_VECTOR
        sType = "Vector"
    Case VT_R4
        sType = "Single"
    Case VT_R8
        sType = "Double"
    Case VT_DISPATCH
        sType = "Object"
    Case VT_UI1
        sType = "Byte"
    Case VT_CY
        sType = "Currency"
    Case Else
        sType = "???"
        Debug.Assert 1 = 0
    End Select
    If (lType And VT_ARRAY) = VT_ARRAY Then
        sType = "() As " & sType
    Else
        sType = " As " & sType
    End If
    psTranslateType = sType

End Function


Private Sub cboClass_Click()
Dim iTypeInfo As Long
Dim i As Long
Dim sRtf As String
Dim sTypeLibName As String
Dim sDateString As String
Dim sTypeLibString As String

    ' Clear list
    Status "Getting Type Library Information..."
    lstMembers.Clear

    If (cboClass.ListIndex > -1) Then
        ' Evaluate the contents:
        iTypeInfo = cboClass.ItemData(cboClass.ListIndex)
        
        If (iTypeInfo > 0) Then
                
            Screen.MousePointer = vbHourglass
        
            ' Prepare the RTF header:
            sDateString = "yr" & Year(Now) & "\mo" & Month(Now) & "\dy" & Day(Now) & "\hr" & Hour(Now) & "\min" & Minute(Now)
            sTypeLibString = m_cTLI.Name
            
            sRtf = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}" & vbCrLf
            sRtf = sRtf & "{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}{\f15\fswiss\fcharset0\fprq2{\*\panose 020b0604030504040204}Verdana;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;" & vbCrLf
            sRtf = sRtf & "\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;" & vbCrLf
            sRtf = sRtf & "\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\lang2057\cgrid \snext0 Normal;}{\s1\sb240\sa60\keepn\widctlpar\adjustright \b\f15\fs28\lang2057\kerning28\cgrid \sbasedon0 \snext0 heading 1;}{\s3\sb240\sa60\keepn\widctlpar\adjustright" & vbCrLf
            sRtf = sRtf & "\b\f15\lang2057\cgrid \sbasedon0 \snext0 heading 3;}{\*\cs10 \additive Default Paragraph Font;}{\s15\qc\widctlpar\adjustright \b\f15\fs16\lang2057\cgrid \sbasedon0 \snext0 caption;}{\s16\li720\widctlpar\adjustright \f2\fs16\lang2057\cgrid" & vbCrLf
            sRtf = sRtf & "\sbasedon0 \snext16 Code;}{\*\cs17 \additive \ul\cf12 \sbasedon10 FollowedHyperlink;}{\*\cs18 \additive \ul\cf2 \sbasedon10 Hyperlink;}{\s19\widctlpar\adjustright \f15\fs20\lang2057\cgrid \sbasedon0 \snext19 Paragraph;}}{\info" & vbCrLf
            sRtf = sRtf & "{\title " & sTypeLibName & "Interface Definition}{\author ActiveX Documenter}{\operator ActiveX Documenter}{\creatim\" & sDateString & "}{\revtim\" & sDateString & "}{\printim\" & sDateString & "}{\version1}{\edmins8}" & vbCrLf
            sRtf = sRtf & "{\nofchars1789}{\*\company vbaccelerator}{\nofcharsws2197}{\vern89}}\paperw11906\paperh16838 \widowctrl\ftnbj\aenddoc\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\headery709\footery709\colsx709\endnhere\sectdefaultcl {\*\pnseclvl1" & vbCrLf
            sRtf = sRtf & "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & vbCrLf
            sRtf = sRtf & "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & vbCrLf
            sRtf = sRtf & "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain \widctlpar\adjustright \fs20\lang2057\cgrid {\b\f1\fs24" & vbCrLf
            
            sRtf = sRtf & sTypeLibString & " Interface Definition \par " & vbCrLf & "\par }"
            
            sRtf = sRtf & "{\b\f1 General Information" & vbCrLf & "\par }"
            sRtf = sRtf & "\pard \widctlpar\tx993\adjustright {\f1 " & psGetGeneralInfoRtf(0) & vbCrLf
            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(1) & vbCrLf
            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(2) & vbCrLf
            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(3) & vbCrLf
            sRtf = sRtf & "\par }\pard \widctlpar\adjustright {" & vbCrLf
            sRtf = sRtf & "\par }"
        
            If (iTypeInfo = &HFFFFFFF) Then
                ' Do all the enums:
                sRtf = sRtf & "{\b\f1 Enumerations" & vbCrLf
                sRtf = sRtf & "\par }{\f1 This section lists enumerations exposed by " & sTypeLibString & "." & vbCrLf
                sRtf = sRtf & "\par }{\f1" & vbCrLf
                
                Status "Reading enums..."
                ProgressMax = (cboClass.ListCount - 1) * 2
                For i = 0 To cboClass.ListCount - 1
                    If (i <> cboClass.ListIndex) Then
                        ProgressValue = i
                        iTypeInfo = cboClass.ItemData(i)
                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString = "enum") Then
                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
                            pDisplayInterfaces iTypeInfo, sRtf
                        End If
                    End If
                Next i
                
                sRtf = sRtf & "}{" & vbCrLf
                sRtf = sRtf & "\par" & vbCrLf
                sRtf = sRtf & "\par }{\b\f1 Interfaces}{\b\f1" & vbCrLf
                sRtf = sRtf & "\par }{\f1 This section lists }{\f1 the Classes exposed by " & sTypeLibString & ".  For each class, the methods and events are listed.}{\f1" & vbCrLf
                sRtf = sRtf & "\par }{" & vbCrLf
                sRtf = sRtf & "\par }" & vbCrLf
                
                ' Do all the interfaces:
                For i = 0 To cboClass.ListCount - 1
                    If (i <> cboClass.ListIndex) Then
                        iTypeInfo = cboClass.ItemData(i)
                        ProgressValue = i + cboClass.ListCount - 1
                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString <> "enum") Then
                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
                            pDisplayInterfaces iTypeInfo, sRtf
                            sRtf = sRtf & "{\par}" & vbCrLf
                        End If
                    End If
                Next i
                
            Else
                Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
                sRtf = sRtf & " {\f1 "
                pDisplayInterfaces iTypeInfo, sRtf
                sRtf = sRtf & "\par }" & vbCrLf
            End If
            
            ' Complete the RTF:
            sRtf = sRtf & "\par }}"
                
            Status "Displaying the TypeLibrary Document..."
            ' DIsplay the Rtf:
            rtfDocument.Contents(SF_RTF) = sRtf

            Screen.MousePointer = vbDefault
            Status "Ready."
            ProgressValue = 0
        Else
            Status "No Type Library Information."
        End If
    Else
        Status "No Type Library Information."
    End If

End Sub
Private Sub pDisplayInterfaces(iTypeInfo As Long, ByRef sRtf As String)
Dim sJunk As String
Dim sType As String
Dim sGUID As String
Dim sMembers() As String
Dim sHelp() As String
Dim iMemberCount As Long
Dim sEvents() As String
Dim sEventHelp() As String
Dim iEventCount As Long
Dim iBelongsTo As Long
Dim sParseItem As String
Dim iMember As Long

    pEvaluateMember m_cTLI.TypeInfos(iTypeInfo), sJunk, sGUID, sJunk, sType, sMembers(), sHelp(), iMemberCount
    
    For iBelongsTo = 1 To m_iCount
        If (m_iBelongsToInterface(iBelongsTo) = iTypeInfo) Then
            If (Left$(m_sInterfaces(iBelongsTo), 2) = "__") Then
                ' events:
                pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sEvents(), sEventHelp(), iEventCount
            Else
                ' methods/properties:
                pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sMembers(), sHelp(), iMemberCount
            End If
        End If
    Next iBelongsTo

    ' Add the information to the class list:
    If (sType = "enum") Then
        lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
        lstMembers.ItemData(lstMembers.NewIndex) = &HFFFFFFF
        sRtf = sRtf & "\par " & "Public Enum " & m_sInterfaces(iTypeInfo) & vbCrLf
        For iMember = 1 To iMemberCount
            lstMembers.AddItem vbTab & sMembers(iMember)
            lstMembers.ItemData(lstMembers.NewIndex) = iMember
            sRtf = sRtf & "\par \tab " & sMembers(iMember) & vbCrLf
        Next iMember
        lstMembers.AddItem "End Enum"
        lstMembers.ItemData(lstMembers.NewIndex) = &HFFFFFFF
        sRtf = sRtf & "\par End Enum" & vbCrLf
    Else
        
        lstMembers.AddItem "Class:" & vbTab & m_sInterfaces(iTypeInfo) & " " & sGUID
        lstMembers.ItemData(lstMembers.NewIndex) = -1
        pParseGUID sGUID
        sRtf = sRtf & "{\b\f1 " & m_sInterfaces(iTypeInfo) & " " & sGUID & vbCrLf
        sRtf = sRtf & "\par }{" & vbCrLf
        
        sRtf = sRtf & "\par }{\f1\ul Methods" & vbCrLf & "}"
        If (iMemberCount > 0) Then
            lstMembers.AddItem "Methods:"
            lstMembers.ItemData(lstMembers.NewIndex) = -1
            For iMember = 1 To iMemberCount
                sParseItem = "Public " & sMembers(iMember)
                sRtf = sRtf & "{\b\f1 " & vbCrLf & "\par " & sMembers(iMember) & vbCrLf & "\par }"
                If Trim$(Len(sHelp(iMember))) > 0 Then
                    sParseItem = sParseItem & " '" & sHelp(iMember)
                    sRtf = sRtf & "{\f1 " & sHelp(iMember) & "}"
                End If
                lstMembers.AddItem sParseItem
                lstMembers.ItemData(lstMembers.NewIndex) = iMember
                
            Next iMember
        Else
            lstMembers.AddItem "No Methods."
            lstMembers.ItemData(lstMembers.NewIndex) = -1
            sRtf = sRtf & "{\f1" & "None " & vbCrLf & "\par}"
        End If
        
        sRtf = sRtf & "{\f1" & vbCrLf & "\par}{\f1\ul Events" & vbCrLf & "\par}"
        If (iEventCount > 0) Then
            lstMembers.AddItem "Events:"
            lstMembers.ItemData(lstMembers.NewIndex) = -1

            For iMember = 1 To iEventCount
                sParseItem = "Public Event " & Mid$(sEvents(iMember), 5)
                sRtf = sRtf & "{\b\f1 " & vbCrLf & "\par " & sParseItem & vbCrLf & "\par }"
                If Trim$(Len(sEventHelp(iMember))) > 0 Then
                    sParseItem = sParseItem & " '" & sEventHelp(iMember)
                    sRtf = sRtf & "{\f1 " & sEventHelp(iMember) & "}"
                End If
                lstMembers.AddItem sParseItem
                lstMembers.ItemData(lstMembers.NewIndex) = iMember
            Next iMember
            
        Else
            lstMembers.AddItem "No Events."
            lstMembers.ItemData(lstMembers.NewIndex) = -1
            sRtf = sRtf & "{\f1" & "None " & vbCrLf & "\par}"
            
        End If
        
    End If

End Sub
Private Sub pShowTab(ByVal lTab As Long)
Dim i As Long
    picTab(lTab - 1).Visible = True
    picTab(lTab - 1).ZOrder
    For i = 0 To 2
        If (i <> lTab - 1) Then
            picTab(i).Visible = False
        End If
    Next i
    mnuEdit(3).Enabled = (lTab = 1)
    
End Sub
Private Sub pResizeTabs()
Dim i As Long
Dim lH As Long
Dim lMinH As Long
Dim lW As Long
Dim lCentreH As Long
Dim lbT As Long

    If (tabMain.ClientWidth > 0) And (tabMain.ClientHeight > 0) Then

        For i = 0 To 2
            With tabMain
                picTab(i).BorderStyle = 0
                picTab(i).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
            End With
        Next i
        
        ' resize tab 1 contents:
        lstGeneral.Width = picTab(0).ScaleWidth
        lH = picTab(0).ScaleHeight - lstMembers.Top
        If (lH > 0) Then
            lstMembers.Move 0, lstMembers.Top, picTab(0).ScaleWidth, lH
        End If
        
        ' resize tab 2 contents:
        If Not (lstAvailable.Visible) Then
            lblAvail.Width = Me.ScaleWidth
        Else
            lH = picTab(1).ScaleHeight - lstAvailable.Top
            lW = (picTab(1).ScaleWidth - cmdAddInterface - 4 * Screen.TwipsPerPixelX) \ 2
            If (lH > 0) And (lW > 0) Then
                ' Reposition available list:
                lstAvailable.Move 0, lstAvailable.Top, lW, lH
                ' Attempt to reposition the two add lists within bounds:
                lMinH = cmdAddInterface.Height * 2 + 2 * Screen.TwipsPerPixelY
                lCentreH = lH \ 2
                If (lCentreH - lblCreate.Height < lMinH) Then lCentreH = lMinH + lblCreate.Height
                lH = lCentreH - lblCreate.Height
                lblCreate.Left = lW + cmdAddInterface.Width + 4 * Screen.TwipsPerPixelX
                lstInterface.Move lblCreate.Left, lstInterface.Top, lW, lH
                lbT = lstInterface.Top + (lstInterface.Height) \ 2
                cmdAddInterface.Move lW + Screen.TwipsPerPixelX, lbT - cmdAddInterface.Height - Screen.TwipsPerPixelY
                cmdRemoveInterface.Move cmdAddInterface.Left, cmdAddInterface.Top + cmdAddInterface.Height + Screen.TwipsPerPixelY
                lbT = lstInterface.Top + lstInterface.Height + 2 * Screen.TwipsPerPixelY
                lblSuper.Move lstInterface.Left, lbT
                lH = picTab(2).ScaleHeight - lblSuper.Top - lblSuper.Height
                If (lH < lMinH) Then lH = lMinH
                lstSuperClass.Move lstInterface.Left, lblSuper.Top + lblSuper.Height, lW, lH
                lbT = lstSuperClass.Top + (lstSuperClass.Height) \ 2
                cmdAddSuperClass.Move cmdAddInterface.Left, lbT - cmdAddSuperClass.Height - Screen.TwipsPerPixelY
                cmdRemoveSuperClass.Move cmdAddInterface.Left, cmdAddSuperClass.Top + cmdAddSuperClass.Height + Screen.TwipsPerPixelY
            End If
        End If
        
        ' resize tab 3 contents:
        rtfDocument.Move 0, 0, picTab(2).ScaleWidth, picTab(2).ScaleHeight
    End If
    
End Sub

Private Sub Form_Load()

    ' set up class combo:
    cboClass.Clear
    cboClass.AddItem "<No Interfaces>"
    cboClass.ListIndex = 0
    cboClass.Enabled = False
    Me.Caption = App.Title & " (No Open TypeLib)"
    
    ' set up tabs:
    tabMain.AddTab "Browse", , , "Browse"
    tabMain.AddTab "SuperClass", , , "SuperClass"
    tabMain.AddTab "Document", , , "Document"
    
    pResizeTabs
    pShowTab 1
    
    ' set up toolbar:
    With tbrMain
        .ButtonImageSource = CTBResourceBitmap
        .ButtonImageResourceID = 24
        .CreateToolbar 16
        .AddButton 1, "Open [Ctrl-O]", 1, , , , CTBNormal, "open"
        .AddButton 2, "", -1, , , , CTBSeparator
        .AddButton 3, "Save [Ctrl-S]", 2, , , , , "save"
        .AddButton 4, "", -1, , , , CTBSeparator
        .AddButton 5, "Copy [Ctrl-C]", 4, , , , , "copy"
        .AddButton 6, "", -1, , , , CTBSeparator
        .AddButton 7, "Print [Ctrl-P]", 6, , , , , "print"
        .AddButton 8, "", -1, , , , CTBSeparator
        .AddButton 9, "vbAccelerator on the Web", 7, , , , , "vba"
    End With
    
    ' set up rebar:
    rbrMain.AddBandByHwnd tbrMain.hwnd, , False, , "ToolbarBand"
    rbrMain.AddBandByHwnd cboClass.hwnd, "Class", False, , "ClassBand"
    
    ' set up 'statusbar':
    ThinBorder lblStatus.hwnd
    ThinBorder picProgress.hwnd
       
    ' Load options
    Set m_cMRU = New cMRUFileList
    pLoadOptions
    
    ' parse command line if any:
    ' options are /p /pt "filename"
    pParseCommand
    
    ' The superclasser isn't ready yet:
    lblAvail.Caption = "Sorry, super-class functionality is not available in this release."
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Save options:
    pSaveOptions
    ' Clear up objects
    Set m_cTLI = Nothing
End Sub

Private Sub Form_Resize()
Dim lT As Long
Dim lST As Long
Dim lSW As Long, lSH As Long
Dim lTH As Long, lTW As Long

    ' Do Rebar:
    rbrMain.Resize Me
    lT = rbrMain.RebarHeight * Screen.TwipsPerPixelY
    ' Do Status Bar:
    lST = Me.ScaleHeight - lblStatus.Height
    lSW = Me.ScaleWidth - picProgress.Width - 4 * Screen.TwipsPerPixelX
    If (lSW < picProgress.Width) Then lSW = picProgress.Width
    lblStatus.Move Screen.TwipsPerPixelX, lST, lSW, lblStatus.Height
    picProgress.Move lblStatus.Left + lblStatus.Width + 2 * Screen.TwipsPerPixelX, lST
    ' Resize Tabs:
    lTH = Me.ScaleHeight - lT - 8 * Screen.TwipsPerPixelY - lblStatus.Height
    lTW = Me.ScaleWidth - 2 * Screen.TwipsPerPixelX
    If (lTH > 0) And (lTW > 0) Then
        tabMain.Move Screen.TwipsPerPixelX, lT + 4 * Screen.TwipsPerPixelY, lTW, lTH
        pResizeTabs
    End If
End Sub

Private Sub pLoadOptions()
    Dim cR As New cRegistry
    If Not (pbLoadOptionsFromKey(HKEY_CURRENT_USER)) Then
        pbLoadOptionsFromKey (HKEY_LOCAL_MACHINE)
    End If
    m_cMRU.MaxFileCount = 4
    pDisplayMRU
End Sub
Private Function pbLoadOptionsFromKey(ByVal hKey As ERegistryClassConstants) As Boolean
Dim cR As New cRegistry
    cR.ClassKey = hKey
    cR.SectionKey = "Software\vbaccelerator\TLibHelp\MRUFiles"
    If (cR.KeyExists) Then
        m_cMRU.Load cR
        cR.SectionKey = "Software\vbaccelerator\TLibHelp"
        cR.ValueKey = "InitOpenDir"
        cR.ValueType = REG_SZ
        m_sInitOpenDir = cR.Value
        cR.ValueKey = "InitSaveDir"
        cR.ValueType = REG_SZ
        m_sInitSaveDir = cR.Value
        If (Trim$(m_sInitSaveDir) = "") Then
            ' Get MyDocuments directory:
            On Error Resume Next
            m_sInitSaveDir = ShellFolder(CSIDL_PERSONAL)
        End If
        pbLoadOptionsFromKey = True
    End If
End Function
Private Sub pSaveOptionsToKey(ByVal hKey As ERegistryClassConstants)
Dim cR As New cRegistry
    cR.ClassKey = hKey
    cR.SectionKey = "Software\vbaccelerator\TLibHelp\MRUFiles"
    m_cMRU.Save cR
    cR.SectionKey = "Software\vbaccelerator\TLibHelp"
    cR.ValueKey = "InitOpenDir"
    cR.ValueType = REG_SZ
    cR.Value = m_sInitOpenDir
    cR.ValueKey = "InitSaveDir"
    cR.ValueType = REG_SZ
    cR.Value = m_sInitSaveDir
End Sub
Private Sub pSaveOptions()
    pSaveOptionsToKey HKEY_CURRENT_USER
    pSaveOptionsToKey HKEY_LOCAL_MACHINE
End Sub
Private Sub pDisplayMRU()
Dim iFile As Long
    For iFile = 1 To m_cMRU.FileCount
        If (m_cMRU.FileExists(iFile)) Then
            If (iFile = 1) And cboClass.Enabled Then mnuFile(iFile + 6).Checked = True
            mnuFile(iFile + 6).Caption = m_cMRU.MenuCaption(iFile)
            mnuFile(iFile + 6).Tag = CStr(iFile)
            mnuFile(iFile + 6).Visible = True
        Else
            mnuFile(iFile + 6).Visible = False
        End If
    Next iFile
    If (m_cMRU.FileCount > 0) Then
        mnuFile(11).Visible = True
    Else
        mnuFile(11).Visible = False
    End If
    
End Sub


Private Sub mnuEdit_Click(Index As Integer)
Dim lI As Long
    Select Case Index
    Case 0
        ' copy
        tbrMain_ButtonClick tbrMain.ButtonIndex("copy")
    Case 2
        ' select all
        Select Case tabMain.SelectedTab
        Case 1
            pSelect lstMembers, True
        Case 2
        Case 3
            rtfDocument.SelectAll
        End Select
    Case 3
        ' invert selection (tab 1 only)
        pSelect lstMembers, , True
    Case 4
        ' clear selection:
        Select Case tabMain.SelectedTab
        Case 1
            pSelect lstGeneral, False, , True
            pSelect lstMembers, False, , True
        Case 2
        Case 3
            rtfDocument.SelectNone
        End Select
    End Select
End Sub
Private Sub pSelect(ByRef lstThis As ListBox, Optional bSelectState As Variant, Optional bInvert As Variant, Optional bEverything As Boolean = False)
Dim lI As Long
    If Not (IsMissing(bSelectState)) Then
        For lI = 0 To lstThis.ListCount - 1
            If (lstThis.ItemData(lI) > 0) Or (bEverything) Then
                lstThis.Selected(lI) = bSelectState
            End If
        Next lI
    Else
        If (IsMissing(bInvert)) Then
            bInvert = True
        End If
        For lI = 0 To lstThis.ListCount - 1
            If (lstThis.ItemData(lI) > 0) Or (bEverything) Then
                lstThis.Selected(lI) = Not (lstThis.Selected(lI))
            End If
        Next lI
    End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim sFIle As String
    Select Case Index
    Case 0
        ' open
        tbrMain_ButtonClick tbrMain.ButtonIndex("open")
    Case 2
        ' save
        tbrMain_ButtonClick tbrMain.ButtonIndex("save")
    Case 4
        ' print
        tbrMain_ButtonClick tbrMain.ButtonIndex("print")
    Case 5
        ' print preview
        ' ...
        
        ' ...
    Case 7 To 10
        ' open MRU file
        sFIle = m_cMRU.file(CLng(mnuFile(Index).Tag))
        If (pbGetTypeLibInfo(sFIle)) Then
            ' success:
            m_cMRU.AddFile sFIle
            pDisplayMRU
        End If
    Case 12
        ' exit
        Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
    Case 0
        ' help
        MsgBox "Sorry, help contents are not available yet.", vbInformation
    Case 1
        ' Shell vbAccelerator:
        On Error Resume Next
        ShellEx "http://www.dogma.demon.co.uk", , , , , Me.hwnd
        If (Err.Number <> 0) Then
            MsgBox "Sorry, I failed to open the web site: www.dogma.demon.co.uk due to an error." & vbCrLf & vbCrLf & "[" & Err.Description & "]", vbExclamation
        End If
    Case 3
        frmAbout.Show vbModal, Me
    End Select
End Sub

Private Sub rbrMain_HeightChanged(lNewHeight As Long)
    Form_Resize
End Sub

Private Sub tabMain_TabClick(ByVal lTab As Long)
    pShowTab lTab
End Sub

Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
Dim sFIle As String
    Select Case lButton
    Case 1
        ' open
        Dim cDlg As New cCommonDialog
        If (cDlg.VBGetOpenFileName( _
            sFIle, , True, , , , _
            "ActiveX Files (*.OCX;*.DLL;*.TLB;*.OLB;*.EXE)|*.OCX;*.DLL;*.TLB;*.OLB;*.EXE|ActiveX Controls (*.OCX)|*.OCX|ActiveX DLLs (*.DLL)|*.DLL|Type Libraries (*.TLB;*.OLB)|*.TLB;*.OLB|ActiveX Executables (*.EXE)|*.EXE|All Files (*.*)|*.", _
            1, m_sInitOpenDir, "Choose Type Library", "*.OCX", Me.hwnd)) Then
            If (pbGetTypeLibInfo(sFIle)) Then
                ' success
                m_cMRU.AddFile sFIle
                pDisplayMRU
            End If
        End If
    Case 3
        ' save:
        If Not (m_cTLI Is Nothing) And cboClass.Enabled Then
            If (cboClass.ItemData(cboClass.ListIndex) = &HFFFFFFF) Then
                pSave m_cTLI.Name & " Interface Definition.RTF"
            Else
                pSave cboClass.List(cboClass.ListIndex) & ".RTF"
            End If
        Else
            MsgBox "No Type Library is Loaded to Save a Document for.", vbInformation
        End If
    Case 5
        ' copy:
        pCopy
    Case 7
        If Not (m_cTLI Is Nothing) And cboClass.Enabled Then
            ' print:
            rtfDocument.PrintDoc m_cTLI.Name & " Interface Definition"
        Else
            MsgBox "No Type Library is Loaded to Save a Document for.", vbInformation
        End If
    Case 9
        ' vba
        mnuHelp_Click 1
    End Select
        
End Sub
Private Sub pSave(ByVal sName As String)
Dim sTitle As String
Dim eType As ERECFileTypes
Dim iPos As Integer
Dim iFIlterIndex As Long
Dim sExt As String

    Dim cc As New cCommonDialog
    
    sName = m_sInitSaveDir & sName
    iFIlterIndex = 1
    Debug.Print cc.VBGetSaveFileName(sName, sTitle, True, "Rich Text Document (*.RTF)|*.RTF|Text Document (*.TXT)|*.TXT|All FIles (*.*)|*.*", iFIlterIndex, m_sInitSaveDir, "Choose Location to Save Document to", "RTF", Me.hwnd, OFN_PATHMUSTEXIST Or OFN_NOREADONLYRETURN)
    If (sName <> "") Then
        iPos = InStr(sName, sTitle)
        If (iPos <> 0) Then
            m_sInitSaveDir = Left$(sName, (iPos - 1))
        End If
        If iFIlterIndex = 1 Then
            eType = SF_RTF
        ElseIf iFIlterIndex = 2 Then
            eType = SF_TEXT
        Else
            ' check extension:
            For iPos = Len(sTitle) To 1 Step -1
                If (Mid$(sTitle, iPos, 1) = ".") Then
                    sExt = UCase$(Mid$(sTitle, iPos + 1))
                End If
            Next iPos
            Select Case sExt
            Case "TXT"
                eType = SF_TEXT
            Case Else
                eType = SF_RTF
            End Select
        End If
        
        Dim iFile As Integer, sRtf As String
        
        On Error Resume Next
        
        sRtf = rtfDocument.Contents(eType)
        If (Err.Number <> 0) Then
            MsgBox "Failed to get the document to save to disk:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
            Exit Sub
        End If
        Kill sName
        Err.Clear
        
        iFile = FreeFile
        Open sName For Binary Access Write Lock Read As #iFile
        If (Err.Number <> 0) Then
            MsgBox "Failed to open the file '" & sName & "' for writing:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
            Close #iFile
            Exit Sub
        End If
        Put #iFile, , sRtf
        If (Err.Number <> 0) Then
            MsgBox "Failed to write the document to the file '" & sName & "'." & vbCrLf & vbCrLf & Err.Description, vbExclamation
        End If
        Close #iFile
    End If
End Sub
Private Sub pCopy()
Dim sOut As String
Dim bIsGeneralSel As Boolean
Dim lI As Long

    ' depends on what tab we are on:
    Select Case tabMain.SelectedTab
    Case 1
        ' copy the selected items in the general list:
        For lI = 0 To lstGeneral.ListCount - 1
            If (lstGeneral.Selected(lI)) Then
                sOut = sOut & lstGeneral.List(lI) & vbCrLf
                bIsGeneralSel = True
            End If
        Next lI
        If (bIsGeneralSel) Then
            sOut = sOut & vbCrLf
        End If
        ' copy the selected members:
        For lI = 0 To lstMembers.ListCount - 1
            If (lstMembers.Selected(lI)) Then
                sOut = sOut & lstMembers.List(lI) & vbCrLf
            End If
        Next lI
        Clipboard.Clear
        Clipboard.SetText sOut
    Case 2
        ' todo..
    Case 3
        ' Call the copy method on the richedit box:
        rtfDocument.Copy
    End Select
End Sub
