Attribute VB_Name = "Database_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 25/04/1999
' * Time             : 16:06
' * Module Name      : Database_Module
' * Module Filename  : Database.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Function ReadRecordSet(record As Recordset, sField As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:56
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : ReadRecordSet
   ' * Parameters       :
   ' *                    record As Recordset
   ' *                    sField As String
   ' **********************************************************************
   ' * Comments         :
   ' *  Read the recordset
   ' *
   ' **********************************************************************

   If IsNull(record(sField)) Then
      Select Case VarType(record(sField))
         Case vbLong: ReadRecordSet = 0
         Case vbString: ReadRecordSet = ""
         Case vbDate: ReadRecordSet = "01/01/1000"
      End Select
   Else
      ReadRecordSet = record(sField)
   End If

End Function

Function ReadRecordSetField(record As Recordset, nField As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:56
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : ReadRecordSetField
   ' * Parameters       :
   ' *                    record As Recordset
   ' *                    nField As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Teste si le paramètre est vide ou non ***

   If IsNull(record.Fields(nField)) Then
      Select Case VarType(record.Fields(nField))
         Case vbLong: ReadRecordSetField = 0
         Case vbString: ReadRecordSetField = ""
         Case vbDate: ReadRecordSetField = "01/01/1000"
      End Select
   Else
      ReadRecordSetField = record.Fields(nField)
   End If

End Function

Sub SaveOLEField(record As Recordset, sFieldName As String, sFileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:56
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : SaveOLEField
   ' * Parameters       :
   ' *                    record As Recordset
   ' *                    sFieldName As String
   ' *                    sFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Save an OLE field to a file

   Const nChunkSize = 32767

   Dim nTotalSize       As Long
   Dim nNumChunks       As Long
   Dim nToRead          As Long
   Dim nI               As Integer
   Dim nFile            As Integer

   On Error Resume Next

   ' *** Takes the total size
   nTotalSize = record(sFieldName).FieldSize

   ' *** Count the number of parts
   nNumChunks = nTotalSize \ nChunkSize - (nTotalSize Mod nChunkSize <> 0)

   If (nTotalSize < nChunkSize) Then
      nToRead = nTotalSize
   Else
      nToRead = nChunkSize
   End If

   ' *** Allocates memory
   ReDim szNoteArray(nNumChunks) As String * nChunkSize

   ' *** Initialization
   For nI = 1 To nNumChunks
      szNoteArray(nI) = record(sFieldName).GetChunk((nI - 1) * nChunkSize, nChunkSize)
   Next

   ' *** Delete the file if it exists
   Kill sFileName

   ' *** Open the file
   nFile = FreeFile
   Open sFileName For Output Access Write As #nFile

   ' *** Save all the file
   For nI = 1 To nNumChunks
      If ((nI = nNumChunks) And nI > 1) Then nToRead = nTotalSize - (nChunkSize * (nI - 1))
      Print #nFile, left$(szNoteArray(nI), nToRead)
   Next

   ' *** Close the file
   Close #nFile

End Sub

Sub GetFileFromDB(BinaryField As Field, sFileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:56
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : GetFileFromDB
   ' * Parameters       :
   ' *                    BinaryField As Field
   ' *                    sFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *  Will retrieve an entire Binary field and write it to disk
   ' *
   ' **********************************************************************

   Dim NumBlocks        As Long
   Dim TotalSize        As Long
   Dim RemBlocks        As Integer
   Dim CurSize          As Integer
   Dim nBlockSize       As Long
   Dim nI               As Integer
   Dim nFile            As Integer
   Dim CurChunk         As String

   If IsNull(BinaryField) Then Exit Sub

   nBlockSize = 32000    ' Set size of chunk.

   ' *** Get field size.
   TotalSize = BinaryField.FieldSize()
   NumBlocks = TotalSize \ nBlockSize   ' Set number of chunks.

   ' *** Set number of remaining bytes.
   RemBlocks = TotalSize Mod nBlockSize

   ' *** Set starting size of chunk.
   CurSize = nBlockSize
   nFile = FreeFile ' Get free file number.

   Open sFileName For Binary As #nFile  ' Open the file.
   For nI = 0 To NumBlocks
      If nI = NumBlocks Then CurSize = RemBlocks
      CurChunk = BinaryField.GetChunk(nI * nBlockSize, CurSize)
      Put #nFile, , CurChunk   ' Write chunk to file.
   Next
   Close nFile

End Sub

Function SaveFileToDB(sSource As String, BynaryField As Field) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:57
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : SaveFileToDB
   ' * Parameters       :
   ' *                    sSource As String
   ' *                    BynaryField As Field
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Will save a file to a Binary field

   Dim nNumBlocks       As Integer
   Dim nFile            As Integer
   Dim nI               As Integer
   Dim nFileLen         As Long
   Dim nLeftOver        As Long
   Dim sFileData        As String
   Dim nBlockSize       As Long

   On Error GoTo Error_SaveFileToDB

   nBlockSize = 32000    ' Set size of chunk.

   ' *** Open the sSource file.
   nFile = FreeFile
   Open sSource For Binary Access Read As nFile

   ' *** Get the length of the file.
   nFileLen = LOF(nFile)
   If nFileLen = 0 Then
      SaveFileToDB = 0
      Exit Function
   End If

   ' *** Calculate the number of blocks to read and nLeftOver bytes.
   nNumBlocks = nFileLen \ nBlockSize
   nLeftOver = nFileLen Mod nBlockSize

   ' *** Read the nLeftOver data, writing it to the table.
   sFileData = String$(nLeftOver, 32)
   Get nFile, , sFileData
   BynaryField.AppendChunk (sFileData)

   ' *** Read the remaining blocks of data, writing them to the table.
   sFileData = String$(nBlockSize, Chr$(32))
   For nI = 1 To nNumBlocks
      Get nFile, , sFileData
      BynaryField.AppendChunk (sFileData)
   Next
   Close #nFile
   SaveFileToDB = nFileLen

   SaveFileToDB = 0

   Exit Function

Error_SaveFileToDB:
   SaveFileToDB = -err
   Exit Function

End Function

Function CopyDBStruct(sDBTo As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/03/1999
   ' * Time             : 23:22
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : CopyDBStruct
   ' * Parameters       :
   ' *                    sDBTo As String
   ' **********************************************************************
   ' * Comments         : Copy the entire structure of a table
   ' *  from a database to another one
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_CopyDBStruct

   Dim DBTo             As Database
   Dim nI               As Integer
   Dim nJ               As Integer
   Dim tblTableDefObj   As TableDef
   Dim fldFieldObj      As Field
   Dim indIndexObj      As index
   Dim tdf              As TableDef
   Dim fld              As Field
   Dim idx              As index

   ' *** Delete the eventual existing output DB
   On Error Resume Next
   Kill sDBTo
   On Error GoTo ERROR_CopyDBStruct

   ' *** Open the output database
   Set DBTo = CreateDatabase(sDBTo, dbLangGeneral & ";pwd=anthony")

   ' *** For Each tdf In DB.Tabledefs
   For nI = 0 To DB.TableDefs.Count - 1
      Set tdf = DB.TableDefs(nI)

      ' *** Create a new table
      Set tblTableDefObj = DB.CreateTableDef(nI)

      ' *** Strip off owner if needed
      tblTableDefObj.Name = StripOwner(tdf.Name)

      ' *** Create the fields
      For nJ = 0 To tdf.Fields.Count - 1
         Set fld = tdf.Fields(nJ)

         ' *** Create this new field
         Set fldFieldObj = tdf.CreateField(fld.Name)
         fldFieldObj.Type = fld.Type
         fldFieldObj.Size = fld.Size
         fldFieldObj.DefaultValue = fld.DefaultValue
         On Error Resume Next
         fldFieldObj.AllowZeroLength = fld.AllowZeroLength
         On Error GoTo ERROR_CopyDBStruct
         err.Clear
         fldFieldObj.Required = fld.Required

         ' *** Add this field
         tblTableDefObj.Fields.Append fldFieldObj
      Next

      ' *** Create the indexes
      For nJ = 0 To tdf.Indexes.Count - 1
         Set idx = tdf.Indexes(nJ)

         ' *** Create the index
         Set indIndexObj = tdf.CreateIndex(idx.Name)
         indIndexObj.Fields = idx.Fields
         indIndexObj.Unique = idx.Unique
         indIndexObj.Primary = idx.Primary

         ' *** Add this index
         tblTableDefObj.Indexes.Append indIndexObj
      Next

      ' *** Append this new table
      DBTo.TableDefs.Append tblTableDefObj
Next_Table:
   Next

   DBTo.Close
   Set DBTo = Nothing

   CopyDBStruct = True
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_CopyDBStruct:
   If err = 3110 Then
      Resume Next_Table
   End If
   MsgBox "Can not create the output database " & Error, vbCritical
   CopyDBStruct = False
   Exit Function

End Function

Function StripOwner(sTableName As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/03/1999
   ' * Time             : 23:28
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : StripOwner
   ' * Parameters       :
   ' *                    sTableName As String
   ' **********************************************************************
   ' * Comments         : Strips the owner off of ODBC table names
   ' *
   ' *
   ' **********************************************************************

   If InStr(sTableName, ".") > 0 Then
      sTableName = Mid$(sTableName, InStr(sTableName, ".") + 1, Len(sTableName))
   End If
   StripOwner = sTableName

End Function

Public Sub ImportData(sDBFrom As String, DBTo As Database)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 27/03/1999
   ' * Time             : 01:18
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : ImportData
   ' * Parameters       :
   ' *                    sDBFrom As String
   ' *                    DBTo As Database
   ' **********************************************************************
   ' * Comments         : Import items From the database
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ImportData

   Dim DBFrom           As Database

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Load frmProgress
   frmProgress.MessageText = Translation("Importing databaseµ186")
   frmProgress.bCancel = True
   frmProgress.Show
   frmProgress.ZOrder

   ' *** Open the import DB
   Set DBFrom = OpenDatabase(sDBFrom, False, False, ";PWD=anthony")

   Call ImportTable("Articles", DBFrom, DBTo)
   Call ImportTable("Code", DBFrom, DBTo)
   Call ImportTable("HTML", DBFrom, DBTo)
   Call ImportTable("Files", DBFrom, DBTo)
   Call ImportTable("Samples", DBFrom, DBTo)

   DBFrom.Close
   Set DBFrom = Nothing

   Unload frmProgress
   Set frmProgress = Nothing

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ImportData:
   MsgBox "Error " & Error & Chr$(13) & Translation("When Importing the databaseµ329")
   If Not (DBFrom Is Nothing) Then
      DBFrom.Close
      Set DBFrom = Nothing
   End If

   Unload frmProgress
   Set frmProgress = Nothing

   Exit Sub

End Sub

Private Function VerifyCategory(DBFrom As Database, recordItem As Recordset, DBTo As Database) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 15:57
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : VerifyCategory
   ' * Parameters       :
   ' *                    DBFrom As Database
   ' *                    recordItem As Recordset
   ' *                    DBTo As Database
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim recordCategoryFrom As Recordset
   Dim recordCategoryTo As Recordset

   Dim sSQL             As String

   Dim nCategory        As Long
   Dim nParent          As Long

   nCategory = 0

   ' *** Verify for the category
   sSQL = "Select * From Categories "
   sSQL = sSQL & "Where (ID = " & recordItem("Category") & ") "
   Set recordCategoryFrom = DBFrom.OpenRecordset(sSQL)

   If recordCategoryFrom.EOF = False Then
      ' *** Get the parent
      nParent = recordCategoryFrom("Parent")

      ' *** Verify if it is the same category
      sSQL = "Select * From Categories "
      sSQL = sSQL & "Where (LCase$(Category) = '" & Replace(LCase$(recordCategoryFrom("Category")), "'", "''") & "') "
      Set recordCategoryTo = DBTo.OpenRecordset(sSQL)

      If recordCategoryTo.EOF = True Then
         ' *** Add the category
         recordCategoryTo.AddNew
         recordCategoryTo(Translation("Categoryµ83")) = recordCategoryFrom(Translation("Categoryµ83"))
         recordCategoryTo("Parent") = 0
         nCategory = recordCategoryTo("ID")
         recordCategoryTo.Update
      Else
         nCategory = recordCategoryTo("ID")
      End If

      recordCategoryTo.Close
      Set recordCategoryTo = Nothing

      recordCategoryFrom.Close
      Set recordCategoryFrom = Nothing

      If nParent = 0 Then
         ' *** Set to no parent
         sSQL = "Update Categories "
         sSQL = sSQL & "Set Parent = 0 "
         sSQL = sSQL & "Where ID = " & nCategory
         DBTo.Execute sSQL

      Else
         ' *** Get the parent in the import database
         sSQL = "Select * From Categories "
         sSQL = sSQL & "Where (ID = " & nParent & ") "
         Set recordCategoryFrom = DBFrom.OpenRecordset(sSQL)

         If recordCategoryFrom.EOF = False Then
            ' *** Find the parent
            sSQL = "Select * From Categories "
            sSQL = sSQL & "Where (LCase$(Category) = '" & Replace(LCase$(recordCategoryFrom("Category")), "'", "''") & "') "
            Set recordCategoryTo = DBTo.OpenRecordset(sSQL)
            If recordCategoryTo.EOF = False Then
               ' *** Update the parent
               sSQL = "Update Categories "
               sSQL = sSQL & "Set Parent = " & recordCategoryTo("ID") & " "
               sSQL = sSQL & "Where ID = " & nCategory
               DBTo.Execute sSQL

            Else
               ' *** Parent not found

            End If

            recordCategoryTo.Close
            Set recordCategoryTo = Nothing

         End If

      End If
   End If

   On Error Resume Next
   recordCategoryFrom.Close
   Set recordCategoryFrom = Nothing

   VerifyCategory = nCategory

End Function

Private Sub ImportTable(sTable As String, DBFrom As Database, DBTo As Database)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/04/1999
   ' * Time             : 17:51
   ' * Module Name      : Database_Module
   ' * Module Filename  : Database.bas
   ' * Procedure Name   : ImportTable
   ' * Parameters       :
   ' *                    sTable As String
   ' *                    DBFrom As Database
   ' *                    DBTo As Database
   ' **********************************************************************
   ' * Comments         : Import a specific table
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ImportTable

   Dim sSQL             As String
   Dim record           As Recordset
   Dim recordTo         As Recordset
   Dim recordItem       As Recordset
   Dim recordTmp        As Recordset

   Dim nItemID          As Long
   Dim nCategory        As Long
   Dim bNewItem         As Boolean

   Dim nI               As Long
   Dim nJ               As Long

   BeginTrans

   ' *** Import all table
   sSQL = "Select * From " & sTable & ", Items "
   sSQL = sSQL & "Where (" & sTable & ".Item = Items.ID) "
   Set record = DBFrom.OpenRecordset(sSQL)

   If record.RecordCount > 0 Then
      frmProgress.MessageText = "Importing " & sTable

      record.MoveLast
      record.MoveFirst
      frmProgress.Maximum = record.RecordCount
   End If

   For nJ = 1 To record.RecordCount
      frmProgress.Progress = nJ

      ' *** Get item ID, or create it
      GoSub GetItem

      If nItemID <> 0 Then
         sSQL = "Select * From " & sTable & " "
         sSQL = sSQL & "Where (Item = " & nItemID & ") "

         Set recordTo = DBTo.OpenRecordset(sSQL)

         If recordTo.EOF = True Then
            ' *** Add this record
            Set recordTmp = DBFrom.OpenRecordset("Select * From " & sTable & "")
            recordTo.AddNew
            For nI = 1 To recordTmp.Fields.Count - 1
               recordTo(nI) = record(nI)
            Next
            recordTo("Item") = nItemID
            recordTo.Update
            recordTmp.Close
            Set recordTmp = Nothing

         Else
            ' *** Verify the description of the item
            recordTo.Close
            Set recordTo = Nothing

            sSQL = "Select * From Items "
            sSQL = sSQL & "Where (Items.ID = " & nItemID & ") "

            Set recordTo = DBTo.OpenRecordset(sSQL)

            If LCase$(recordTo("Title")) = LCase$(record("Title")) Then
               ' *** Update it
               recordTo.Close
               Set recordTo = Nothing
               sSQL = "Select * From " & sTable & " "
               sSQL = sSQL & "Where (Item = " & nItemID & ") "
               Set recordTo = DBTo.OpenRecordset(sSQL)

               sSQL = "Select * From " & sTable & " "
               sSQL = sSQL & "Where (ID = " & record("" & sTable & ".ID") & ") "
               Set recordTmp = DBFrom.OpenRecordset(sSQL)
               recordTo.Edit
               For nI = 1 To recordTmp.Fields.Count - 1
                  recordTo(nI) = recordTmp(recordTmp.Fields(nI).Name)
               Next
               recordTo("Item") = nItemID
               recordTo.Update
               recordTmp.Close
               Set recordTmp = Nothing

            Else
               ' *** Add this record
               recordTo.Close
               Set recordTo = Nothing
               sSQL = "Select * From " & sTable & " "
               sSQL = sSQL & "Where (Item = " & nItemID & ") "
               Set recordTo = DBTo.OpenRecordset(sSQL)

               sSQL = "Select * From " & sTable & " "
               sSQL = sSQL & "Where (ID = " & record("" & sTable & ".ID") & ") "
               Set recordTmp = DBFrom.OpenRecordset(sSQL)
               recordTo.AddNew
               For nI = 1 To recordTmp.Fields.Count - 1
                  recordTo(nI) = record(nI)
               Next
               recordTo("Item") = nItemID
               recordTo.Update
               recordTmp.Close
               Set recordTmp = Nothing

            End If
         End If
         recordTo.Close
         Set recordTo = Nothing

      End If

      record.MoveNext
   Next

   record.Close
   Set record = Nothing

   CommitTrans

   Exit Sub

GetItem:
   bNewItem = False
   ' *** Get the item in the export DB
   sSQL = "Select * From Items "
   sSQL = sSQL & "Where (ID = " & record("Item") & ") "
   Set recordItem = DBFrom.OpenRecordset(sSQL)

   If recordItem.EOF = False Then
      ' *** See if the items exists in the import database
      sSQL = "Select * From Items "
      sSQL = sSQL & "Where (LCase$(Title) = '" & Replace(LCase$(recordItem("Title")), "'", "''") & "') "
      sSQL = sSQL & "Or (GUID = '" & Replace(LCase$(recordItem("GUID")), "'", "''") & "') "
      Set recordTo = DBTo.OpenRecordset(sSQL)

      If recordTo.EOF = True Then
         ' *** Add the Item
         recordTo.AddNew
         For nI = 1 To recordTo.Fields.Count - 1
            recordTo(nI) = recordItem(nI)
         Next
         nItemID = recordTo("ID")
         recordTo.Update
         bNewItem = True

      Else
         ' *** Verify if it has not being changed
         If recordItem("Created") <> recordTo("Created") Then
            ' *** Modify it except the first field
            ' *** wich contains the primary key
            recordTo.Edit
            For nI = 1 To recordTo.Fields.Count - 1
               recordTo(nI) = recordItem(nI)
            Next
            nItemID = recordTo("ID")
            recordTo.Update
            bNewItem = True

         Else
            ' *** The item already exists, get the ID
            nItemID = recordTo("ID")

         End If

      End If

      ' *** Get the right category
      nCategory = VerifyCategory(DBFrom, recordItem, DBTo)

      ' **** Update the category of this item
      sSQL = "Update Items "
      sSQL = sSQL & "Set Category = " & nCategory & " "
      sSQL = sSQL & "Where ID = " & nItemID
      DBTo.Execute sSQL

   Else
      nItemID = 0
   End If

   recordItem.Close
   Set recordItem = Nothing
   Return

EXIT_ImportTable:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ImportTable:
   If err = 3022 Then Resume EXIT_ImportTable

   Select Case MsgBox("Error " & err.number & ": " & err.Description & vbCrLf & "in ImportTable" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ImportTable
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select

   Resume EXIT_ImportTable

End Sub
