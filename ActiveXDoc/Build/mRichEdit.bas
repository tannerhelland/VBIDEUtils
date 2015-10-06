Attribute VB_Name = "mRichEdit"
Option Explicit

' General
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const LF_FACESIZE = 32


'' /*
' *  RICHEDIT.H
' *
' *  Purpose:
' *      RICHEDIT v2.0 public definitions.  Note that there is additional
' *      functionality available for v2.0 that is not in the original
' *      Windows 95 release.
' *
' *  Copyright (c) 1985-1996, Microsoft Corporation
' */

' #ifndef _RICHEDIT_
'public const _RICHEDIT_

' #ifdef _WIN32
' #include <pshpack4.h>
' #elif !defined(RC_INVOKED)
' #pragma pack(4)
' #End If

' #ifdef __cplusplus
'extern "C" {
' #endif ' /* __cplusplus */

' /* To mimic older RichEdit behavior, simply set _RICHEDIT_VER to the appropriate value */
' /*      Version 1.0     =&H0100  */
' /*      Version 2.0     =&H0200  */
' #ifndef _RICHEDIT_VER
Public Const RICHEDIT_VER = &H210
' #End If

' /*
' *  To make some structures which can be passed between 16 and 32 bit windows
' *  almost compatible, padding is introduced to the 16 bit versions of the
' *  structure.
' */
' #ifdef _WIN32
' #   define  _WPAD   /' #' #/
' #Else
' #   define  _WPAD   WORD
' #End If

Public Const cchTextLimitDefault = 32767

' /* Richedit2.0 Window Class. */

Public Const RICHEDIT_CLASSA = "RichEdit20A"
Public Const RICHEDIT_CLASS10A = "RICHEDIT"           '// Richedit 1.0

' #ifndef MACPORT
'public Const RICHEDIT_CLASSW = "RichEdit20W"
' #else   ' /*----------------------MACPORT */
'public const RICHEDIT_CLASSW     =TEXT("RichEdit20W") ' /* MACPORT change */
' #endif ' /* MACPORT  */

' #if (_RICHEDIT_VER >= =&H0200 )
' #ifdef UNICODE
'public const RICHEDIT_CLASS      RICHEDIT_CLASSW
' #Else
Public Const RICHEDIT_CLASS = RICHEDIT_CLASSA
' #endif ' /* UNICODE */
' #Else
'public const RICHEDIT_CLASS      RICHEDIT_CLASS10A
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

' /* RichEdit messages */

' #ifndef WM_CONTEXTMENU
Public Const WM_CONTEXTMENU = &H7B
' #End If

' #ifndef WM_PRINTCLIENT
Public Const WM_PRINTCLIENT = &H318
' #End If

' #ifndef EM_GETLIMITTEXT
'public Const EM_GETLIMITTEXT = (WM_USER + 37)
' #End If

' #ifndef EM_POSFROMCHAR
'public Const EM_POSFROMCHAR = (WM_USER + 38)
'public Const EM_CHARFROMPOS = (WM_USER + 39)
' #End If

' #ifndef EM_SCROLLCARET
'public Const EM_SCROLLCARET = (WM_USER + 49)
' #End If
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_DISPLAYBAND = (WM_USER + 51)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXLIMITTEXT = (WM_USER + 53)
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_FORMATRANGE = (WM_USER + 57)
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_GETEVENTMASK = (WM_USER + 59)
Public Const EM_GETOLEINTERFACE = (WM_USER + 60)
Public Const EM_GETPARAFORMAT = (WM_USER + 61)
Public Const EM_GETSELTEXT = (WM_USER + 62)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_PASTESPECIAL = (WM_USER + 64)
Public Const EM_REQUESTRESIZE = (WM_USER + 65)
Public Const EM_SELECTIONTYPE = (WM_USER + 66)
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const EM_SETOLECALLBACK = (WM_USER + 70)
Public Const EM_SETPARAFORMAT = (WM_USER + 71)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_STREAMOUT = (WM_USER + 74)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_FINDWORDBREAK = (WM_USER + 76)
Public Const EM_SETOPTIONS = (WM_USER + 77)
Public Const EM_GETOPTIONS = (WM_USER + 78)
Public Const EM_FINDTEXTEX = (WM_USER + 79)
' #ifdef _WIN32
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80)
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81)
' #End If

' /* Richedit v2.0 messages */
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)

Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)

' /* enum for use with EM_GET/SETTEXTMODE */
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8          ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32           ' /* default behavior */
End Enum

Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const EM_GETAUTOURLDETECT = (WM_USER + 92)
Public Const EM_SETPALETTE = (WM_USER + 93)
Public Const EM_GETTEXTEX = (WM_USER + 94)
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)

' /* Far East specific messages */
Public Const EM_SETPUNCTUATION = (WM_USER + 100)
Public Const EM_GETPUNCTUATION = (WM_USER + 101)
Public Const EM_SETWORDWRAPMODE = (WM_USER + 102)
Public Const EM_GETWORDWRAPMODE = (WM_USER + 103)
Public Const EM_SETIMECOLOR = (WM_USER + 104)
Public Const EM_GETIMECOLOR = (WM_USER + 105)
Public Const EM_SETIMEOPTIONS = (WM_USER + 106)
Public Const EM_GETIMEOPTIONS = (WM_USER + 107)
Public Const EM_CONVPOSITION = (WM_USER + 108)

Public Const EM_SETLANGOPTIONS = (WM_USER + 120)
Public Const EM_GETLANGOPTIONS = (WM_USER + 121)
Public Const EM_GETIMECOMPMODE = (WM_USER + 122)

Public Const EM_FINDTEXTW = (WM_USER + 123)
Public Const EM_FINDTEXTEXW = (WM_USER + 124)

' /* BiDi specific messages */
Public Const EM_SETBIDIOPTIONS = (WM_USER + 200)
Public Const EM_GETBIDIOPTIONS = (WM_USER + 201)

' /* Options for EM_SETLANGOPTIONS and EM_GETLANGOPTIONS */
Public Const IMF_AUTOKEYBOARD = &H1
Public Const IMF_AUTOFONT = &H2
Public Const IMF_IMECANCELCOMPLETE = &H4      '// high completes the comp string when aborting, low cancels.
Public Const IMF_IMEALWAYSSENDNOTIFY = &H8

' /* Values for EM_GETIMECOMPMODE */
Public Const ICM_NOTOPEN = &H0
Public Const ICM_LEVEL3 = &H1
Public Const ICM_LEVEL2 = &H2
Public Const ICM_LEVEL2_5 = &H3
Public Const ICM_LEVEL2_SUI = &H4

' /* New notifications */

Public Const EN_MSGFILTER = &H700
Public Const EN_REQUESTRESIZE = &H701
Public Const EN_SELCHANGE = &H702
Public Const EN_DROPFILES = &H703
Public Const EN_PROTECTED = &H704
Public Const EN_CORRECTTEXT = &H705                   ' /* PenWin specific */
Public Const EN_STOPNOUNDO = &H706
Public Const EN_IMECHANGE = &H707                     ' /* Far East specific */
Public Const EN_SAVECLIPBOARD = &H708
Public Const EN_OLEOPFAILED = &H709
Public Const EN_OBJECTPOSITIONS = &H70A
Public Const EN_LINK = &H70B
Public Const EN_DRAGDROPDONE = &H70C

' /* BiDi specific notifications */

Public Const EN_ALIGN_LTR = &H710
Public Const EN_ALIGN_RTL = &H711

' /* Event notification masks */

Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
Public Const ENM_CORRECTTEXT = &H400000               ' /* PenWin specific */
Public Const ENM_SCROLLEVENTS = &H8
Public Const ENM_DRAGDROPDONE = &H10

' /* Far East specific notification mask */
Public Const ENM_IMECHANGE = &H800000                 ' /* unused by RE2.0 */
Public Const ENM_LANGCHANGE = &H1000000
Public Const ENM_OBJECTPOSITIONS = &H2000000
Public Const ENM_LINK = &H4000000

' /* New edit control styles */

Public Const ES_SAVESEL = &H8000
Public Const ES_SUNKEN = &H4000
Public Const ES_DISABLENOSCROLL = &H2000
' /* same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value */
Public Const ES_SELECTIONBAR = &H1000000
' /* same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop */
Public Const ES_NOOLEDRAGDROP = &H8

' /* New edit control extended style */
' #ifdef  _WIN32
Public Const ES_EX_NOCALLOLEINIT = &H1000000
' #End If

' /* These flags are used in FE Windows */
Public Const ES_VERTICAL = &H400000
Public Const ES_NOIME = &H80000
Public Const ES_SELFIME = &H40000

' /* Edit control options */
Public Const ECO_AUTOWORDSELECTION = &H1
Public Const ECO_AUTOVSCROLL = &H40
Public Const ECO_AUTOHSCROLL = &H80
Public Const ECO_NOHIDESEL = &H100
Public Const ECO_READONLY = &H800
Public Const ECO_WANTRETURN = &H1000
Public Const ECO_SAVESEL = &H8000
Public Const ECO_SELECTIONBAR = &H1000000
Public Const ECO_VERTICAL = &H400000                  ' /* FE specific */


' /* ECO operations */
Public Const ECOOP_SET = &H1
Public Const ECOOP_OR = &H2
Public Const ECOOP_AND = &H3
Public Const ECOOP_XOR = &H4

' /* new word break function actions */
Public Const WB_CLASSIFY = 3
Public Const WB_MOVEWORDLEFT = 4
Public Const WB_MOVEWORDRIGHT = 5
Public Const WB_LEFTBREAK = 6
Public Const WB_RIGHTBREAK = 7

' /* Far East specific flags */
Public Const WB_MOVEWORDPREV = 4
Public Const WB_MOVEWORDNEXT = 5
Public Const WB_PREVBREAK = 6
Public Const WB_NEXTBREAK = 7

Public Const PC_FOLLOWING = 1
Public Const PC_LEADING = 2
Public Const PC_OVERFLOW = 3
Public Const PC_DELIMITER = 4
Public Const WBF_WORDWRAP = &H10
Public Const WBF_WORDBREAK = &H20
Public Const WBF_OVERFLOW = &H40
Public Const WBF_LEVEL1 = &H80
Public Const WBF_LEVEL2 = &H100
Public Const WBF_CUSTOM = &H200

' /* Far East specific flags */
Public Const IMF_FORCENONE = &H1
Public Const IMF_FORCEENABLE = &H2
Public Const IMF_FORCEDISABLE = &H4
Public Const IMF_CLOSESTATUSWINDOW = &H8
Public Const IMF_VERTICAL = &H20
Public Const IMF_FORCEACTIVE = &H40
Public Const IMF_FORCEINACTIVE = &H80
Public Const IMF_FORCEREMEMBER = &H100
Public Const IMF_MULTIPLEEDIT = &H400

' /* Word break flags (used with WB_CLASSIFY) */
Public Const WBF_CLASS = &HF          '((BYTE) =&H0F)
Public Const WBF_ISWHITE = &H10       '((BYTE) =&H10)
Public Const WBF_BREAKLINE = &H20     '((BYTE) =&H20)
Public Const WBF_BREAKAFTER = &H40    '((BYTE) =&H40)


' /* new data types */

' #ifdef _WIN32
' /* extended edit word break proc (character set aware) */
'typedef LONG (*EDITWORDBREAKPROCEX)(char *pchText, LONG cchText, BYTE bCharSet, INT action);
' #End If

' /* all character format measurements are in twips */
Public Type CHARFORMAT
    cbSize As Long
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    ' VB hostile...
    szFaceName(0 To LF_FACESIZE - 1) As String
    wPad2 As Integer
End Type

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const CHARFORMAT CHARFORMATW
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* UNICODE */
' #Else
'public const CHARFORMAT CHARFORMATA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

' /* CHARFORMAT masks */
Public Const CFM_BOLD = &H1
Public Const CFM_ITALIC = &H2
Public Const CFM_UNDERLINE = &H4
Public Const CFM_STRIKEOUT = &H8
Public Const CFM_PROTECTED = &H10
Public Const CFM_LINK = &H20                  ' /* Exchange hyperlink extension */
Public Const CFM_SIZE = &H80000000
Public Const CFM_COLOR = &H40000000
Public Const CFM_FACE = &H20000000
Public Const CFM_OFFSET = &H10000000
Public Const CFM_CHARSET = &H8000000

' /* CHARFORMAT effects */
Public Const CFE_BOLD = &H1
Public Const CFE_ITALIC = &H2
Public Const CFE_UNDERLINE = &H4
Public Const CFE_STRIKEOUT = &H8
Public Const CFE_PROTECTED = &H10
Public Const CFE_LINK = &H20
Public Const CFE_AUTOCOLOR = &H40000000       ' /* NOTE: this corresponds to */
                                        ' /* CFM_COLOR, which controls it */
Public Const yHeightCharPtsMost = 1638

' /* EM_SETCHARFORMAT wParam masks */
Public Const SCF_SELECTION = &H1
Public Const SCF_WORD = &H2
Public Const SCF_DEFAULT = &H0            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8         '// modifier for SCF_SELECTION; says that
                                   ' // the format came from a toolbar, etc. and
                                   ' // therefore UI formatting rules should be
                                   ' // used instead of strictly formatting the
                                   ' // selection.


Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Public Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
End Type

'typedef struct _textrangew
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;   ' /* allocated by caller, zero terminated by RichEdit */
'} TEXTRANGEW;

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const TEXTRANGE   TEXTRANGEW
' #Else
'public const TEXTRANGE   TEXTRANGEA
' #endif ' /* UNICODE */
' #Else
'public const TEXTRANGE   TEXTRANGEA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */


'typedef DWORD (CALLBACK *EDITSTREAMCALLBACK)(DWORD dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb);

Public Type EDITSTREAM
    dwCookie As Long     ' /* user value passed to callback as first parameter */
    dwError As Long      ' /* last error */
    pfnCallback As Long  'EDITSTREAMCALLBACK
End Type

' /* stream formats */

'Public Const SF_TEXT = &H1
'Public Const SF_RTF = &H2
Public Const SF_RTFNOOBJS = &H3           ' /* outbound only */
Public Const SF_TEXTIZED = &H4            ' /* outbound only */
Public Const SF_UNICODE = &H10            ' /* Unicode file of some kind */

' /* Flag telling stream operations to operate on the selection only */
' /* EM_STREAMIN will replace the current selection */
' /* EM_STREAMOUT will stream out the current selection */
Public Const SFF_SELECTION = &H8000

' /* Flag telling stream operations to operate on the common RTF keyword only */
' /* EM_STREAMIN will accept the only common RTF keyword */
' /* EM_STREAMOUT will stream out the only common RTF keyword */
Public Const SFF_PLAINRTF = &H4000

Public Type FINDTEXT
    chrg As CHARRANGE
    lpstrText As Long
End Type

'typedef struct _findtextw
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;
'} FINDTEXTW;'

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const FINDTEXT    FINDTEXTW
' #Else
'public const FINDTEXT    FINDTEXTA
' #endif ' /* UNICODE */
' #Else
'public const FINDTEXT    FINDTEXTA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */

Public Type FINDTEXTEX
    chrg As CHARRANGE
    lpstrText As Long
    chrgText As CHARRANGE
End Type

'typedef struct _findtextexw
'{
'    CHARRANGE chrg;
'    LPWSTR lpstrText;
'    CHARRANGE chrgText;
'} FINDTEXTEXW;'

' #if (_RICHEDIT_VER >= =&H0200)
' #ifdef UNICODE
'public const FINDTEXTEX  FINDTEXTEXW
' #Else
'public const FINDTEXTEX  FINDTEXTEXA
' #endif ' /* UNICODE */
' #Else
'public const FINDTEXTEX  FINDTEXTEXA
' #endif ' /* _RICHEDIT_VER >= =&H0200 */


Public Type FORMATRANGE
    hdc As Long
    hdcTarget As Long
    rc As RECT
    rcPage As RECT
    chrg As CHARRANGE
End Type

' /* all paragraph measurements are in twips */

Public Const MAX_TAB_STOPS = 32
Public Const lDefaultTab = 720

Public Type PARAFORMAT
    cbSize As Long
    wPad1 As Integer
    dwMask As Long
    wNumbering(0 To 1) As Byte
' #if (_RICHEDIT_VER >= =&H0210)
    wEffects(0 To 1) As Byte
' #Else
    wReserved(0 To 1) As Byte
' #endif ' /* _RICHEDIT_VER >= =&H0210 */
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment(0 To 1) As Byte
    cTabCount(0 To 1) As Byte
    rgxTabs(0 To MAX_TAB_STOPS) As Long
End Type

' /* PARAFORMAT mask values */
Public Const PFM_STARTINDENT = &H1
Public Const PFM_RIGHTINDENT = &H2
Public Const PFM_OFFSET = &H4
Public Const PFM_ALIGNMENT = &H8
Public Const PFM_TABSTOPS = &H10
Public Const PFM_NUMBERING = &H20
Public Const PFM_OFFSETINDENT = &H80000000

' /* PARAFORMAT numbering options */
Public Const PFN_BULLET = &H1

' /* PARAFORMAT alignment options */
Public Const PFA_LEFT = &H1
Public Const PFA_RIGHT = &H2
Public Const PFA_CENTER = &H3

' /* CHARFORMAT2 and PARAFORMAT2 structures */

' #ifdef __cplusplus

'struct CHARFORMAT2W : _charformatw
'{
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'};

'struct CHARFORMAT2A : _charformat
'{
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'};

' #else   ' /* regular C-style  */

'type C
'{
'    UINT        cbSize;
''    _WPAD       _wPad1;
 '   DWORD       dwMask;
 '   DWORD       dwEffects;
 '   LONG        yHeight;
 ''   LONG        yOffset;            ' /* > 0 for superscript, < 0 for subscript */
'    COLORREF    crTextColor;
'    BYTE        bCharSet;
'    BYTE        bPitchAndFamily;
'    WCHAR       szFaceName[LF_FACESIZE];
'    _WPAD       _wPad2;
'    WORD        wWeight;            ' /* Font weight (LOGFONT value)      */
'    SHORT       sSpacing;           ' /* Amount to space between letters  */
'    COLORREF    crBackColor;        ' /* Background color                 */
'    LCID        lcid;               ' /* Locale ID                        */
'    DWORD       dwReserved;         ' /* Reserved. Must be 0              */
'    SHORT       sStyle;             ' /* Style handle                     */
'    WORD        wKerning;           ' /* Twip size above which to kern char pair*/
'    BYTE        bUnderlineType;     ' /* Underline type                   */
'    BYTE        bAnimation;         ' /* Animated text like marching ants */
'    BYTE        bRevAuthor;         ' /* Revision author index            */
'    BYTE        bReserved1;
'} CHARFORMAT2W;

Public Type CHARFORMAT2
    cbSize As Long
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long            ' /* > 0 for superscript, < 0 for subscript */
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight(0 To 1) As Byte            ' /* Font weight (LOGFONT value)      */
    sSpacing(0 To 1) As Byte           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle(0 To 1) As Byte             ' /* Style handle                     */
    wKerning(0 To 1) As Byte           ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
End Type

' #endif ' /* C++ */

' #ifdef UNICODE
'public const CHARFORMAT2 CHARFORMAT2W
' #Else
'public const CHARFORMAT2 CHARFORMAT2A
' #End If

'public Const CHARFORMATDELTA = (Len(CHARFORMAT2) - Len(CHARFORMAT))


' /* CHARFORMAT and PARAFORMAT "ALL" masks
'   CFM_COLOR mirrors CFE_AUTOCOLOR, a little hack to easily deal with autocolor*/

Public Const CFM_EFFECTS = (CFM_BOLD Or CFM_ITALIC Or CFM_UNDERLINE Or CFM_COLOR Or _
                     CFM_STRIKEOUT Or CFE_PROTECTED Or CFM_LINK)
Public Const CFM_ALL = (CFM_EFFECTS Or CFM_SIZE Or CFM_FACE Or CFM_OFFSET Or CFM_CHARSET)

' /* New masks and effects -- a parenthesized asterisk indicates that
'   the data is stored by RichEdit2.0, but not displayed */

Public Const CFM_SMALLCAPS = &H40                 ' /* (*)  */
Public Const CFM_ALLCAPS = &H80                   ' /* (*)  */
Public Const CFM_HIDDEN = &H100                   ' /* (*)  */
Public Const CFM_OUTLINE = &H200                  ' /* (*)  */
Public Const CFM_SHADOW = &H400                   ' /* (*)  */
Public Const CFM_EMBOSS = &H800                   ' /* (*)  */
Public Const CFM_IMPRINT = &H1000                 ' /* (*)  */
Public Const CFM_DISABLED = &H2000
Public Const CFM_REVISED = &H4000

Public Const CFM_BACKCOLOR = &H4000000
Public Const CFM_LCID = &H2000000
Public Const CFM_UNDERLINETYPE = &H800000         ' /* (*)  */
Public Const CFM_WEIGHT = &H400000
Public Const CFM_SPACING = &H200000               ' /* (*)  */
Public Const CFM_KERNING = &H100000               ' /* (*)  */
Public Const CFM_STYLE = &H80000                  ' /* (*)  */
Public Const CFM_ANIMATION = &H40000              ' /* (*)  */
Public Const CFM_REVAUTHOR = &H8000

Public Const CFE_SUBSCRIPT = &H10000              ' /* Superscript and subscript are */
Public Const CFE_SUPERSCRIPT = &H20000            ' /*  mutually exclusive           */

Public Const CFM_SUBSCRIPT = CFE_SUBSCRIPT Or CFE_SUPERSCRIPT
Public Const CFM_SUPERSCRIPT = CFM_SUBSCRIPT

Public Const CFM_EFFECTS2 = (CFM_EFFECTS Or CFM_DISABLED Or CFM_SMALLCAPS Or CFM_ALLCAPS _
                    Or CFM_HIDDEN Or CFM_OUTLINE Or CFM_SHADOW Or CFM_EMBOSS _
                    Or CFM_IMPRINT Or CFM_DISABLED Or CFM_REVISED _
                    Or CFM_SUBSCRIPT Or CFM_SUPERSCRIPT Or CFM_BACKCOLOR)

Public Const CFM_ALL2 = (CFM_ALL Or CFM_EFFECTS2 Or CFM_BACKCOLOR Or CFM_LCID _
                    Or CFM_UNDERLINETYPE Or CFM_WEIGHT Or CFM_REVAUTHOR _
                    Or CFM_SPACING Or CFM_KERNING Or CFM_STYLE Or CFM_ANIMATION)

Public Const CFE_SMALLCAPS = CFM_SMALLCAPS
Public Const CFE_ALLCAPS = CFM_ALLCAPS
Public Const CFE_HIDDEN = CFM_HIDDEN
Public Const CFE_OUTLINE = CFM_OUTLINE
Public Const CFE_SHADOW = CFM_SHADOW
Public Const CFE_EMBOSS = CFM_EMBOSS
Public Const CFE_IMPRINT = CFM_IMPRINT
Public Const CFE_DISABLED = CFM_DISABLED
Public Const CFE_REVISED = CFM_REVISED

' /* NOTE: CFE_AUTOCOLOR and CFE_AUTOBACKCOLOR correspond to CFM_COLOR and
'   CFM_BACKCOLOR, respectively, which control them */
Public Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR

' /* Underline types */
Public Const CFU_CF1UNDERLINE = &HFF      ' /* map charformat's bit underline to CF2.*/
Public Const CFU_INVERT = &HFE            ' /* For IME composition fake a selection.*/
Public Const CFU_UNDERLINEDOTTED = &H4    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEDOUBLE = &H3    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEWORD = &H2      ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINE = &H1
Public Const CFU_UNDERLINENONE = 0

' #ifdef __cplusplus
'struct PARAFORMAT2 : _paraformat
'{
'    LONG    dySpaceBefore;          ' /* Vertical spacing before para         */
'    LONG    dySpaceAfter;           ' /* Vertical spacing after para          */
'    LONG    dyLineSpacing;          ' /* Line spacing depending on Rule       */
'    SHORT   sStyle;                 ' /* Style handle                         */
'    BYTE    bLineSpacingRule;       ' /* Rule for line spacing (see tom.doc)  */
'    BYTE    bCRC;                   ' /* Reserved for CRC for rapid searching */
'    WORD    wShadingWeight;         ' /* Shading in hundredths of a per cent  */
'    WORD    wShadingStyle;          ' /* Nibble 0: style, 1: cfpat, 2: cbpat  */
'    WORD    wNumberingStart;        ' /* Starting value for numbering         */
'    WORD    wNumberingStyle;        ' /* Alignment, roman/arabic, (), ), ., etc.*/
'    WORD    wNumberingTab;          ' /* Space bet FirstIndent and 1st-line text*/
'    WORD    wBorderSpace;           ' /* Space between border and text (twips)*/
'    WORD    wBorderWidth;           ' /* Border pen width (twips)             */
'    WORD    wBorders;               ' /* Byte 0: bits specify which borders   */
'                                    ' /* Nibble 2: border style, 3: color index*/
'};

' #else   ' /* regular C-style  */

Public Type PARAMFORMAT2
    cbSize As Long
    wPad1 As Integer
    dwMask As Long
    wNumbering(0 To 1) As Byte
' #if (_RICHEDIT_VER >= =&H0210)
    wEffects(0 To 1) As Long
' #Else
'    WORD    wReserved;
' #endif ' /* _RICHEDIT_VER >= =&H0210 */
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    rgxTabs(0 To MAX_TAB_STOPS - 1) As Byte
    dySpaceBefore As Long          ' /* Vertical spacing before para         */
    dySpaceAfter As Long           ' /* Vertical spacing after para          */
    dyLineSpacing As Long          ' /* Line spacing depending on Rule       */
    sStyle As Integer                  ' /* Style handle                         */
    bLineSpacingRule As Byte       ' /* Rule for line spacing (see tom.doc)  */
    bCRC As Byte                   ' /* Reserved for CRC for rapid searching *
    wShadingWeight As Integer          ' /* Shading in hundredths of a per cent  */
    wShadingStyle As Integer           ' /* Nibble 0: style, 1: cfpat, 2: cbpat  */
    wNumberingStart As Integer         ' /* Starting value for numbering         */
    wNumberingStyle As Integer        ' /* Alignment, roman/arabic, (), ), ., etc.*/
    wNumberingTab As Integer           ' /* Space bet 1st indent and 1st-line text*/
    wBorderSpace As Integer            ' /* Space between border and text (twips)*/
    wBorderWidth As Integer           ' /* Border pen width (twips)             */
    wBorders As Integer                ' /* Byte 0: bits specify which borders   */
                                    ' /* Nibble 2: border style, 3: color index*/
End Type

' #endif ' /* C++   */

' /* PARAFORMAT 2.0 masks and effects */

Public Const PFM_SPACEBEFORE = &H40
Public Const PFM_SPACEAFTER = &H80
Public Const PFM_LINESPACING = &H100
Public Const PFM_STYLE = &H400
Public Const PFM_BORDER = &H800                   ' /* (*)  */
Public Const PFM_SHADING = &H1000                 ' /* (*)  */
Public Const PFM_NUMBERINGSTYLE = &H2000          ' /* (*)  */
Public Const PFM_NUMBERINGTAB = &H4000            ' /* (*)  */
Public Const PFM_NUMBERINGSTART = &H8000          ' /* (*)  */

Public Const PFM_DIR = &H10000
Public Const PFM_RTLPARA = &H10000                ' /* (Version 1.0 flag) */
Public Const PFM_KEEP = &H20000                   ' /* (*)  */
Public Const PFM_KEEPNEXT = &H40000               ' /* (*)  */
Public Const PFM_PAGEBREAKBEFORE = &H80000        ' /* (*)  */
Public Const PFM_NOLINENUMBER = &H100000          ' /* (*)  */
Public Const PFM_NOWIDOWCONTROL = &H200000        ' /* (*)  */
Public Const PFM_DONOTHYPHEN = &H400000           ' /* (*)  */
Public Const PFM_SIDEBYSIDE = &H800000            ' /* (*)  */

Public Const PFM_TABLE = &HC0000000               ' /* (*)  */

' /* Note: PARAFORMAT has no effects */
Public Const PFM_EFFECTS = (PFM_DIR Or PFM_KEEP Or PFM_KEEPNEXT Or PFM_TABLE _
                    Or PFM_PAGEBREAKBEFORE Or PFM_NOLINENUMBER _
                    Or PFM_NOWIDOWCONTROL Or PFM_DONOTHYPHEN Or PFM_SIDEBYSIDE _
                    Or PFM_TABLE)

Public Const PFM_ALL = (PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET Or _
                 PFM_ALIGNMENT Or PFM_TABSTOPS Or PFM_NUMBERING Or _
                 PFM_OFFSETINDENT Or PFM_DIR)

Public Const PFM_ALL2 = (PFM_ALL Or PFM_EFFECTS Or PFM_SPACEBEFORE Or PFM_SPACEAFTER _
                    Or PFM_LINESPACING Or PFM_STYLE Or PFM_SHADING Or PFM_BORDER _
                    Or PFM_NUMBERINGTAB Or PFM_NUMBERINGSTART Or PFM_NUMBERINGSTYLE)

'public const PFE_RTLPARA  =           (PFM_DIR             >> 16)
'public const PFE_RTLPAR              (PFM_RTLPARA         >> 16) ' /* (Version 1.0 flag) */
'public const PFE_KEEP                (PFM_KEEP            >> 16) ' /* (*)  */
'public const PFE_KEEPNEXT            (PFM_KEEPNEXT        >> 16) ' /* (*)  */
'public const PFE_PAGEBREAKBEFORE     (PFM_PAGEBREAKBEFORE >> 16) ' /* (*)  */
'public const PFE_NOLINENUMBER        (PFM_NOLINENUMBER    >> 16) ' /* (*)  */
'public const PFE_NOWIDOWCONTROL      (PFM_NOWIDOWCONTROL  >> 16) ' /* (*)  */
'public const PFE_DONOTHYPHEN         (PFM_DONOTHYPHEN     >> 16) ' /* (*)  */
'public const PFE_SIDEBYSIDE          (PFM_SIDEBYSIDE      >> 16) ' /* (*)  */'

Public Const PFE_TABLEROW = &HC000                ' /* These 3 options are mutually */
Public Const PFE_TABLECELLEND = &H8000            ' /*  exclusive and each imply    */
Public Const PFE_TABLECELL = &H4000               ' /*  that para is part of a table*/

' /*
' *  PARAFORMAT numbering options (values for wNumbering):
' *
' *      Numbering Type      Value   Meaning
' *      tomNoNumbering        0     Turn off paragraph numbering
' *      tomNumberAsLCLetter   1     a, b, c, ...
' *      tomNumberAsUCLetter   2     A, B, C, ...
' *      tomNumberAsLCRoman    3     i, ii, iii, ...
' *      tomNumberAsUCRoman    4     I, II, III, ...
' *      tomNumberAsSymbols    5     default is bullet
' *      tomNumberAsNumber     6     0, 1, 2, ...
' *      tomNumberAsSequence   7     tomNumberingStart is first Unicode to use
' *
' *  Other valid Unicode chars are Unicodes for bullets.
' */


Public Const PFA_JUSTIFY = 4          ' /* New paragraph-alignment option 2.0 (*)


' /* notification structures */

Public Type NMHDR_RICHEDIT
    hwndFrom As Long
    wPad1 As Integer
    idfrom As Long
    wPad2 As Integer
    code As Long
    wPad3 As Long
End Type
' #endif  ' /* !WM_NOTIFY */

Public Type MSGFILTER
    NMHDR As NMHDR_RICHEDIT
    msg As Long
    wPad1 As Integer
    wParam As Long
    wPad2 As Integer
    lParam As Long
End Type

Public Type REQRESIZE
    NMHDR As NMHDR_RICHEDIT
    rc As RECT
End Type

Public Type SELCHANGE
    NMHDR As NMHDR_RICHEDIT
    chrg As CHARRANGE
    seltyp As Long
End Type

Public Const SEL_EMPTY = &H0
Public Const SEL_TEXT = &H1
Public Const SEL_OBJECT = &H2
Public Const SEL_MULTICHAR = &H4
Public Const SEL_MULTIOBJECT = &H8

' /* used with IRichEditOleCallback::GetContextMenu, this flag will be
'   passed as a "selection type".  It indicates that a context menu for
'   a right-mouse drag drop should be generated.  The IOleObject parameter
'   will really be the IDataObject for the drop
' */
Public Const GCM_RIGHTMOUSEDROP = &H8000

Public Type ENDROPFILES
    NMHDR As NMHDR_RICHEDIT
    hDrop As Long
    cp As Long
    fProtected As Long
End Type

Public Type ENPROTECTED
    NMHDR As NMHDR_RICHEDIT
    msg As Long
    wPad1 As Integer
    wParam As Long
    wPad2 As Integer
    lParam As Long
    chrg As CHARRANGE
End Type

Public Type ENSAVECLIPBOARD
    NMHDR As NMHDR_RICHEDIT
    cObjectCount As Long
    cch As Long
End Type

' #ifndef MACPORT
Public Type ENOLEOPFAILED
    NMHDR As NMHDR_RICHEDIT
    iob As Long
    lOper As Long
    hr As Long
End Type
' #End If

Public Const OLEOP_DOVERB = 1

Public Type OBJECTPOSITIONS
    NMHDR As NMHDR_RICHEDIT
    cObjectCount As Long
        ' !!!POINTER to long value!!!
    pcpPositions As Long
End Type

Public Type ENLINK
    NMHDR As NMHDR_RICHEDIT
    msg As Long
    wPad1 As Integer
    wParam As Long
    wPad2 As Integer
    lParam As Long
    chrg As CHARRANGE
End Type

' /* PenWin specific */
Public Type ENCORRECTTEXT
    NMHDR As NMHDR_RICHEDIT
    chrg As CHARRANGE
    seltyp As Integer
End Type

' /* Far East specific */
'typedef struct _punctuation
'{
'    UINT    iSize;
'    LPSTR   szPunctuation;
'} PUNCTUATION;

' /* Far East specific */
'typedef struct _compcolor
'{
'    COLORREF crText;
'    COLORREF crBackground;
'    DWORD dwEffects;
'}COMPCOLOR;


' /* clipboard formats - use as parameter to RegisterClipboardFormat() */
Public Const CF_RTF = "Rich Text Format"
Public Const CF_RTFNOOBJS = "Rich Text Format Without Objects"
Public Const CF_RETEXTOBJ = "RichEdit Text and Objects"

' /* Paste Special */
Public Type REPASTESPECIAL
    dwAspect As Long
    dwParam As Long
End Type

' /*  UndoName info */
Public Enum UNDONAMEID
    UID_UNKNOWN = 0
    UID_TYPING = 1
    UID_DELETE = 2
    UID_DRAGDROP = 3
    UID_CUT = 4
    UID_PASTE = 5
End Enum

' /* flags for the GETEXTEX data structure */
Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1

' /* EM_GETTEXTEX info; this struct is passed in the wparam of the message */
Public Type GETTEXTEX
    cb As Long             ' /* count of bytes in the string             */
    flags As Long          ' /* flags (see the GT_XXX defines            */
    codepage As Long       ' /* code page for translation (CP_ACP for default,
                           '    1200 for Unicode                         */
    lpDefaultChar As Long ';  ' /* replacement for unmappable chars         */
    lpUsedDefChar As Long ';  ' /* pointer to flag set when def char used   */
End Type

' /* flags for the GETTEXTLENGTHEX data structure                         */
Public Const GTL_DEFAULT = 0      ' /* do the default (return ' # of chars)       */
Public Const GTL_USECRLF = 1      ' /* compute answer using CRLFs for paragraphs*/
Public Const GTL_PRECISE = 2      ' /* compute a precise answer                 */
Public Const GTL_CLOSE = 4        ' /* fast computation of a "close" answer     */
Public Const GTL_NUMCHARS = 8     ' /* return the number of characters          */
Public Const GTL_NUMBYTES = 16    ' /* return the number of _bytes_             */

' /* EM_GETTEXTLENGTHEX info; this struct is passed in the wparam of the msg */
Public Type GETTEXTLENGTHEX
    flags As Long          ' /* flags (see GTL_XXX defines)              */
    codepage As Long       ' /* code page for translation (CP_ACP for default,
                              '1200 for Unicode                         */
End Type
    
' /* BiDi specific features */
Public Type BIDIOPTIONS
    cbSize As Long
    wPad1 As Integer
    wMask As Integer
    wEffects As Integer
End Type

' /* BIDIOPTIONS masks */
' #if (_RICHEDIT_VER == =&H0100)
Public Const BOM_DEFPARADIR = &H1             ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOM_PLAINTEXT = &H2              ' /* Use plain text layout (obsolete) */
Public Const BOM_NEUTRALOVERRIDE = &H4        ' /* Override neutral layout (obsolete) */
' #endif ' /* _RICHEDIT_VER == =&H0100 */
Public Const BOM_CONTEXTREADING = &H8         ' /* Context reading order */
Public Const BOM_CONTEXTALIGNMENT = &H10      ' /* Context alignment */

' /* BIDIOPTIONS effects */
' #if (_RICHEDIT_VER == =&H0100)
Public Const BOE_RTLDIR = &H1                 ' /* Default paragraph direction (implies alignment) (obsolete) */
Public Const BOE_PLAINTEXT = &H2              ' /* Use plain text layout (obsolete) */
Public Const BOE_NEUTRALOVERRIDE = &H4        ' /* Override neutral layout (obsolete) */
' #endif ' /* _RICHEDIT_VER == =&H0100 */
Public Const BOE_CONTEXTREADING = &H8         ' /* Context reading order */
Public Const BOE_CONTEXTALIGNMENT = &H10      ' /* Context alignment */

' /* Additional EM_FINDTEXT[EX] flags */
Public Const FR_MATCHDIAC = &H20000000
Public Const FR_MATCHKASHIDA = &H40000000
Public Const FR_MATCHALEFHAMZA = &H80000000

' /* UNICODE embedding character */
' #ifndef WCH_EMBEDDING
Public Const WCH_EMBEDDING = &HFFFC
' #endif ' /* WCH_EMBEDDING */
        

' #undef _WPAD

' #ifdef _WIN32
' #include <poppack.h>
' #elif !defined(RC_INVOKED)
' #pragma pack()
' #End If

' #ifdef __cplusplus
'}
' #endif  ' /* __cplusplus */

' #endif ' /* !_RICHEDIT_ */


' /*
 '* Edit Control Messages
 '*/
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETRECT = &HB2
Public Const EM_SETRECT = &HB3
Public Const EM_SETRECTNP = &HB4
Public Const EM_SCROLL = &HB5
Public Const EM_LINESCROLL = &HB6
Public Const EM_SCROLLCARET = &HB7
Public Const EM_GETMODIFY = &HB8
Public Const EM_SETMODIFY = &HB9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETHANDLE = &HBC
Public Const EM_GETHANDLE = &HBD
Public Const EM_GETTHUMB = &HBE
Public Const EM_LINELENGTH = &HC1
Public Const EM_REPLACESEL = &HC2
Public Const EM_GETLINE = &HC4
Public Const EM_LIMITTEXT = &HC5
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_FMTLINES = &HC8
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_GETPASSWORDCHAR = &HD2
'#if(WINVER >= =&H0400)
Public Const EM_SETMARGINS = &HD3
Public Const EM_GETMARGINS = &HD4
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT          ' /* ;win40 Name change */
Public Const EM_GETLIMITTEXT = &HD5
Public Const EM_POSFROMCHAR = &HD6
Public Const EM_CHARFROMPOS = &HD7
'#End If ' /* WINVER >= =&H0400 */
'/*
' * Edit Control Styles
' */
Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&
'#if(WINVER >= =&H0400)
Public Const ES_NUMBER = &H2000&
'#endif /* WINVER >= =&H0400 */

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Public Const OFS_MAXPATHNAME = 128
Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

' Printing support:

' VB API VIEWER VERSION OF DOCINFO STRUCTURE IS WRONG!
Type DOCINFO
    cbSize As Long
    lpszDocName As Long
    lpszOutput As Long
End Type

Public Type PRINTDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Public Declare Function PRINTDLG Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As PRINTDLG) As Integer

Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Const MM_TEXT = 1

' Streaming support:
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private m_sText As String
Private m_lPos As Long
Private m_lLen As Long
Private m_bFileMode As Boolean

Public Property Let FileMode(ByVal bMode As Boolean)
    m_bFileMode = bMode
End Property
Public Property Get FileMode() As Boolean
    FileMode = m_bFileMode
End Property
Public Sub ClearStreamText()
    m_sText = ""
End Sub
Public Property Get StreamText() As String
    StreamText = m_sText
End Property
Public Property Let StreamText(ByRef sText As String)
    m_sText = sText
    m_lPos = 1
    m_lLen = Len(m_sText)
End Property
Public Function LoadCallBack( _
        ByVal dwCookie As Long, _
        ByVal lPtrPbBuff As Long, _
        ByVal cb As Long, _
        ByVal pcb As Long _
    ) As Long
Dim sBuf As String
Dim b() As Byte
Dim lLen As Long
Dim lRead As Long
    
    If (m_bFileMode) Then
        ReadFile dwCookie, ByVal lPtrPbBuff, cb, pcb, 0
        CopyMemory lRead, ByVal pcb, 4
        If (lRead < cb) Then
            ' Complete:
            LoadCallBack = 0
        Else
            ' More to read:
            LoadCallBack = lRead
        End If
    Else
        CopyMemory lRead, ByVal pcb, 4
        Debug.Print lRead, cb
        ' Place cb bytes if possible, or place in the whole string:
        If (m_lLen - m_lPos > 0) Then
            If (m_lLen - m_lPos < cb) Then
                ReDim b(0 To (m_lLen - m_lPos)) As Byte
                b = StrConv(Mid$(m_sText, m_lPos), vbFromUnicode)
                lRead = m_lLen - m_lPos + 1
                CopyMemory ByVal lPtrPbBuff, b(0), lRead
                m_lPos = m_lLen + 1
            Else
                ReDim b(0 To cb - 1) As Byte
                b = StrConv(Mid$(m_sText, m_lPos, cb), vbFromUnicode)
                CopyMemory ByVal lPtrPbBuff, b(0), cb
                m_lPos = m_lPos + cb
                lRead = cb
            End If
                        
            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        Else
            lRead = 0
            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        End If
    End If
    
End Function
Public Function SaveCallBack( _
        ByVal dwCookie As Long, _
        ByVal lPtrPbBuff As Long, _
        ByVal cb As Long, _
        ByVal pcb As Long _
    ) As Long
Dim sBuf As String
Dim b() As Byte
Dim lLen As Long

    lLen = cb
    
    If (lLen > 0) Then
        If (m_bFileMode) Then
            WriteFile dwCookie, ByVal lPtrPbBuff, cb, pcb, 0
        Else
            ReDim b(0 To lLen - 1) As Byte
            CopyMemory b(0), ByVal lPtrPbBuff, lLen
            sBuf = StrConv(b, vbUnicode)
            m_sText = m_sText & sBuf
            m_lPos = 1
            m_lLen = Len(m_sText)
        End If
    End If
    SaveCallBack = 0
    
    
End Function

