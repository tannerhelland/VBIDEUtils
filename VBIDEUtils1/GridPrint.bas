Attribute VB_Name = "Grid_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 27/04/2000
' * Time             : 13:10
' * Module Name      : Module1
' * Module Filename  :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************
Option Explicit

Public Sub GridPrint(Grid As vbalGrid, sTitle As String, bLines As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 27/04/2000
   ' * Time             : 13:10
   ' * Module Name      : Grid_Module
   ' * Module Filename  : Grid.bas
   ' * Procedure Name   : GridPrint
   ' * Parameters       :
   ' *                    Grid As Control
   ' *                    sTitle As String
   ' *                    bLines As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   ' *** Print a grid control ***

   Dim nRow          As Long
   Dim nCol          As Long
   Dim sCurrentX     As Single
   Dim sCurrentY     As Single
   Dim sMaxLine      As Single
   Dim nPage         As Integer
   Dim sPage         As String
   Dim sOldY         As Single
   Dim sPageWidth    As Single
   Dim sBeginLeft    As Single
   Dim sBeginGrid    As Single

   If (Grid.Rows = 0) Then Exit Sub

   ' *** We look for the maximum size for the Font.Size ***
   sMaxLine = 99999999

   ' *** By default, we set the Font.Size of the grid ***
   Printer.Font.Size = Grid.Font.Size

   ' *** Here is the maximum width of the page
   sPageWidth = Printer.Width * 0.94 - 200

   ' *** We calculate the maximum possible Font.Size ***
   Do While (sMaxLine > sPageWidth) And (Printer.Font.Size > 0)
      sMaxLine = 0
      For nCol = 2 To Grid.Columns
         sMaxLine = sMaxLine + (Printer.Font.Size / Grid.Font.Size) * Grid.ColumnWidth(nCol) * 15
      Next

      ' *** We change the fontsiz if needed ***
      If (sMaxLine > sPageWidth) Then Printer.Font.Size = Printer.Font.Size - 1
   Loop

   ' *** We begin on page 1 ***
   nPage = 1

   ' *** We put The title ***
   ' *** and the headers of each column ***
   GoSub PRINT_HEADERS

   For nRow = 1 To Grid.Rows - 1
      If (bLines = True) Then Printer.Line (sBeginLeft, sCurrentY)-(sMaxLine, sCurrentY)

      ' *** We print on a new page if needed ***
      If (sCurrentY >= Printer.Height * 0.93 - Printer.TextHeight("A")) Then
         If (bLines = True) Then
            ' *** Bottom line
            Printer.Line (sBeginLeft, sCurrentY - 4 * Printer.TwipsPerPixelY)-(sMaxLine, sCurrentY - 4 * Printer.TwipsPerPixelY)
            ' *** Left line ***
            Printer.Line (sBeginLeft, sBeginGrid)-(sBeginLeft, sCurrentY - 4 * Printer.TwipsPerPixelY)
            ' *** Right line ***
            Printer.Line (sMaxLine, sBeginGrid)-(sMaxLine, sCurrentY - 4 * Printer.TwipsPerPixelY)
         End If

         Printer.NewPage
         nPage = nPage + 1

         ' *** We put The title ***
         ' *** and the headers of each column ***
         GoSub PRINT_HEADERS

      End If

      sCurrentX = 4 * Printer.TwipsPerPixelX

      For nCol = 2 To Grid.Columns
         If (nCol > 1) Then
            sCurrentX = sCurrentX + (Printer.Font.Size / Grid.Font.Size) * Grid.ColumnWidth(nCol - 1) * 15

            If bLines = True Then Printer.Line (sCurrentX - 4 * Printer.TwipsPerPixelX, sBeginGrid)-(sCurrentX - 4 * Printer.TwipsPerPixelX, Printer.CurrentY + (Printer.TextHeight("A") / 2) - 4 * Printer.TwipsPerPixelY)
         End If

         ' *** Print cell text ***
         Printer.CurrentX = sCurrentX
         Printer.CurrentY = sCurrentY + (Printer.TextHeight("A") / 2)
         If Not IsMissing(Grid.CellText(nRow, nCol)) Then Printer.Print Grid.CellText(nRow, nCol)
      Next

      sCurrentY = sCurrentY + (Printer.TextHeight("A") * 2)
   Next
   If (bLines = True) Then
      ' *** Bottom line
      Printer.Line (sBeginLeft, sCurrentY - 4 * Printer.TwipsPerPixelY)-(sMaxLine, sCurrentY - 4 * Printer.TwipsPerPixelY)
      ' *** Left line ***
      Printer.Line (sBeginLeft, sBeginGrid)-(sBeginLeft, sCurrentY - 4 * Printer.TwipsPerPixelY)
      ' *** Right line ***
      Printer.Line (sMaxLine, sBeginGrid)-(sMaxLine, sCurrentY - 4 * Printer.TwipsPerPixelY)
   End If

   Printer.EndDoc

   Exit Sub

PRINT_HEADERS:
   ' *** We print the title ***
   sCurrentY = Printer.CurrentY
   Printer.FontBold = True
   Printer.Print sTitle
   Printer.FontBold = False
   Printer.Print ""

   ' *** We print the page number on the first line ***
   sOldY = Printer.CurrentY
   sPage = "Page " & CStr(nPage)
   Printer.FontItalic = True
   Printer.CurrentX = sPageWidth - Printer.TextWidth(sPage)
   Printer.CurrentY = sCurrentY
   Printer.Print sPage
   Printer.FontItalic = False
   Printer.CurrentY = sOldY

   ' *** We print the grid ***
   Printer.CurrentY = Printer.CurrentY + (Printer.TextHeight("A"))
   sCurrentY = Printer.CurrentY
   sBeginGrid = sCurrentY
   sBeginLeft = 0
   sCurrentX = 4 * Printer.TwipsPerPixelX

   ' *** We print the header of each column ***
   If (bLines = True) Then Printer.Line (sBeginLeft, sCurrentY)-(sMaxLine, sCurrentY)

   Printer.Print
   sCurrentY = Printer.CurrentY + (Printer.TextHeight("A") / 2)
   For nCol = 2 To Grid.Columns
      If (nCol > 1) Then sCurrentX = sCurrentX + (Printer.Font.Size / Grid.Font.Size) * Grid.ColumnWidth(nCol - 1) * 15

      ' *** Print cell text ***
      Printer.CurrentX = sCurrentX
      Printer.CurrentY = sCurrentY
      Printer.Print Grid.ColumnHeader(nCol)
   Next

   sCurrentY = sCurrentY + (Printer.TextHeight("A") * 1.5)
   Printer.Print

   If (bLines = True) Then Printer.Line (sBeginLeft, sCurrentY)-(sMaxLine, sCurrentY)

   Return

End Sub

