Attribute VB_Name = "VCS_ImportExport"
Option Compare Database
Option Explicit

' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
' Set to "*" to export the contents of all tables
' Only used in ExportSource
Private Const INCLUDE_TABLES As String = "*"
' If set to false, it will not export linked tables (only local)
Private Const EXPORT_LINKED_TABLES As Boolean = False

' This is used in ImportSource
Private Const DEBUG_OUTPUT As Boolean = False

' This is used in ExportAllSource
' Causes the VCS_ code to be exported
Private Const ARCHIVE_MYSELF As Boolean = False

' Used to defined subfolder for all import/export objects
' This is handy when you have a FE and/or BE to deal with
Private Const SOURCE_SUB As String = "source.ui\"
'
' ExportProject - export complete project
'
Public Sub ExportProject()
    On Error GoTo errorHandler

    Debug.Print "Started at: " & Time
    ExportSource ("*")
    Debug.Print "Done at: " & Time
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.ExportProject: Error #" & Err.Number & vbCrLf & _
        Err.Description
End Sub
'
' ImportProject - imports complete project
'
Public Sub ImportProject()
    On Error GoTo errorHandler

    Debug.Print "Started at: " & Time
    DeleteSource ("*")
    ImportSource ("*")
    Debug.Print "Done at: " & Time
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
        Err.Description
End Sub
'
' ImportObjectType - remove and import a single object type (e.g. queries, tbldef, forms, modules,...)
'
Public Sub ImportObjectType(sType As String)
    On Error GoTo errorHandler

    Debug.Print "Started at: " & Time
    DeleteSource (sType)
    ImportSource (sType)
    Debug.Print "Done at: " & Time
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.ImportObjectType: Error #" & Err.Number & vbCrLf & _
        Err.Description
End Sub
'
' ExportObjectType - remove and import a single object type (e.g. queries, tbldef, forms, modules,...)
'
Public Sub ExportObjectType(sType As String)
    On Error GoTo errorHandler

    Debug.Print "Started at: " & Time
    ExportSource (sType)
    Debug.Print "Done at: " & Time
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.ExportObjectType: Error #" & Err.Number & vbCrLf & _
        Err.Description
End Sub
'
' Check if object is NOT part of the VCS code
'
Private Function IsNotVCS(ByVal name As String) As Boolean
    IsNotVCS = (name <> "VCS_ImportExport" And _
      name <> "VCS_IE_Functions" And _
      name <> "VCS_File" And _
      name <> "VCS_Dir" And _
      name <> "VCS_String" And _
      name <> "VCS_Loader" And _
      name <> "VCS_Table" And _
      name <> "VCS_Reference" And _
      name <> "VCS_DataMacro" And _
      name <> "VCS_Report" And _
      name <> "VCS_Relation")
End Function

'
' ExportSource - export database objects as specified.
'
Private Sub ExportSource(Optional sObjects As String = "*")
    Dim db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim Def, FSO, Stream
    
    CloseFormsReports
    
    sObjects = LCase(sObjects)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb

    source_path = VCS_Dir.ProjectPath() & SOURCE_SUB
    VCS_Dir.MkDirIfNotExist source_path
    
    '
    ' Export Access references (only on full export)
    '
    If (sObjects = "*") Then
        VCS_Reference.ExportReferences source_path
    End If
    
    '
    ' Export queries
    '
    If (sObjects = "*" Or sObjects = "queries") Then
        obj_path = source_path & "queries\"
        VCS_Dir.ClearTextFilesFromDir obj_path, "bas"
        Debug.Print VCS_String.PadRight("Exporting queries...", 24);
        obj_count = 0
        For Each qry In db.QueryDefs
            DoEvents
            If Left$(qry.name, 1) <> "~" Then
                VCS_IE_Functions.ExportObject acQuery, qry.name, obj_path & qry.name & ".bas"
                '
                ' SQL export
                '
                Set Stream = FSO.CreateTextFile(obj_path & qry.name & ".sql")
                Stream.Write (qry.sql)
                Stream.Close
                
                obj_count = obj_count + 1
            End If
        Next
        VCS_IE_Functions.SanitizeTextFiles obj_path, "bas"
        Debug.Print "[" & obj_count & "]"
    End If
    
    '
    ' Export forms, reports, macros and modules
    '
    For Each obj_type In Split( _
        "forms|Forms|" & acForm & "," & _
        "reports|Reports|" & acReport & "," & _
        "macros|Scripts|" & acMacro & "," & _
        "modules|Modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_name = obj_type_split(1)
        obj_type_num = val(obj_type_split(2))
        
        '
        ' All objects or just one range?
        '
        If (sObjects = "*" Or sObjects = obj_type_label) Then
            obj_path = source_path & obj_type_label & "\"
            obj_count = 0
            VCS_Dir.ClearTextFilesFromDir obj_path, "bas"
            Debug.Print VCS_String.PadRight("Exporting " & obj_type_label & "...", 24);
            For Each doc In db.Containers(obj_type_name).Documents
                DoEvents
                If (Left$(doc.name, 1) <> "~") And _
                   (IsNotVCS(doc.name) Or ARCHIVE_MYSELF) Then
                    VCS_IE_Functions.ExportObject obj_type_num, doc.name, obj_path & doc.name & ".bas"
                    
                    If obj_type_label = "reports" Then
                        VCS_Report.ExportPrintVars doc.name, obj_path & doc.name & ".pv"
                    End If
                    
                    obj_count = obj_count + 1
                End If
            Next
            Debug.Print "[" & obj_count & "]"
    
            If obj_type_label <> "modules" Then
                VCS_IE_Functions.SanitizeTextFiles obj_path, "bas"
            End If
        End If
    Next


    '
    ' Export tables and table data as selected
    '
    If (sObjects = "*" Or sObjects = "tabdef" Or sObjects = "tables") Then
        obj_path = source_path & "tables\"
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        VCS_Dir.ClearTextFilesFromDir obj_path, "txt"
        
        Dim td As DAO.TableDef
        Dim tds As DAO.TableDefs
        Set tds = db.TableDefs
    
        obj_type_label = "tbldef"
        obj_type_name = "Table_Def"
        obj_type_num = acTable
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        obj_data_count = 0
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        
        'move these into Table and DataMacro modules?
        ' - We don't want to determine file extentions here - or obj_path either!
        VCS_Dir.ClearTextFilesFromDir obj_path, "sql"
        VCS_Dir.ClearTextFilesFromDir obj_path, "xml"
        VCS_Dir.ClearTextFilesFromDir obj_path, "LNKD"
        
        Dim IncludeTablesCol As Collection
        Set IncludeTablesCol = StrSetToCol(INCLUDE_TABLES, ",")
        
        Debug.Print VCS_String.PadRight("Exporting " & obj_type_label & "...", 24);
        
        For Each td In tds
            ' This is not a system table
            ' this is not a temporary table
            If Left$(td.name, 4) <> "MSys" And _
            Left$(td.name, 1) <> "~" Then
                If Len(td.connect) = 0 Then
                    ' this is not an external table
                    VCS_Table.ExportTableDef db, td, td.name, obj_path
                    ' export table data
                    If INCLUDE_TABLES = "*" Then
                        DoEvents
                        VCS_Table.ExportTableData CStr(td.name), source_path & "tables\"
                        If Len(Dir$(source_path & "tables\" & td.name & ".txt")) > 0 Then
                            obj_data_count = obj_data_count + 1
                        End If
                    ElseIf (Len(Replace(INCLUDE_TABLES, " ", vbNullString)) > 0) And INCLUDE_TABLES <> "*" Then
                        DoEvents
                        On Error GoTo Err_TableNotFound
                        If IncludeTablesCol(td.name) = td.name Then
                            VCS_Table.ExportTableData CStr(td.name), source_path & "tables\"
                            obj_data_count = obj_data_count + 1
                        End If
Err_TableNotFound:
                    'else don't export table data
                    End If
                Else
                    If EXPORT_LINKED_TABLES Then
                        VCS_Table.ExportLinkedTable td.name, obj_path
                    End If
                End If
                
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"
        If obj_data_count > 0 Then
            Debug.Print VCS_String.PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
        End If
    End If
    
    '
    ' Export table relationships
    '
    If (sObjects = "*" Or sObjects = "relations") Then
        Debug.Print VCS_String.PadRight("Exporting relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        VCS_Dir.ClearTextFilesFromDir obj_path, "txt"
    
        Dim aRelation As DAO.Relation
        
        For Each aRelation In CurrentDb.Relations
            ' Exclude relations from system tables and inherited (linked) relations
            If Not (aRelation.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                    Or aRelation.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                    Or (aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                    DAO.RelationAttributeEnum.dbRelationInherited) Then
                VCS_Relation.ExportRelation aRelation, obj_path & aRelation.name & ".txt"
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"
    End If
End Sub
'
' Import database objects as specified.
'
Private Sub ImportSource(Optional sObjects As String = "*")
    Dim FSO As Object
    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    CloseFormsReports
    Set FSO = CreateObject("Scripting.FileSystemObject")
    sObjects = LCase(sObjects)
    
    source_path = VCS_Dir.ProjectPath() & SOURCE_SUB
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If
    
    '
    ' Import Access References only on full import
    '
    If (sObjects = "*") Then
        If Not VCS_Reference.ImportReferences(source_path) Then
            Debug.Print "Info: no references file in " & source_path
            Debug.Print
        End If
    End If

    '
    ' Import Queries
    '
    If (sObjects = "*" Or sObjects = "queries") Then
        obj_path = source_path & "queries\"
        fileName = Dir$(obj_path & "*.bas")
        
        Dim tempFilePath As String
        tempFilePath = VCS_File.TempFile()
        
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing queries...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                DoEvents
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                VCS_IE_Functions.ImportObject acQuery, obj_name, obj_path & fileName, VCS_File.UsingUcs2
                VCS_IE_Functions.ExportObject acQuery, obj_name, tempFilePath, VCS_File.UsingUcs2
                VCS_IE_Functions.ImportObject acQuery, obj_name, tempFilePath, VCS_File.UsingUcs2
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
        VCS_Dir.DelIfExist tempFilePath
    End If
    
    '
    ' Import Table Definitions (ie schema)
    '
    If (sObjects = "*" Or sObjects = "tbldef") Then
        obj_path = source_path & "tbldef\"
        fileName = Dir$(obj_path & "*.sql")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing tabledefs...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                If DEBUG_OUTPUT Then
                    If obj_count = 0 Then
                        Debug.Print
                    End If
                    Debug.Print "  [debug] table " & obj_name;
                    Debug.Print
                End If
                VCS_Table.ImportTableDef CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    
    '
    ' Import Table Definitions from linked tables
    '
    If (sObjects = "*" Or sObjects = "lnkd") Then
        fileName = Dir$(obj_path & "*.LNKD")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing Linked tabledefs...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                If DEBUG_OUTPUT Then
                    If obj_count = 0 Then
                        Debug.Print
                    End If
                    Debug.Print "  [debug] table " & obj_name;
                    Debug.Print
                End If
                VCS_Table.ImportLinkedTable CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    
    '
    ' Import Table Data
    '
    If (sObjects = "*" Or sObjects = "tables") Then
        obj_path = source_path & "tables\"
        fileName = Dir$(obj_path & "*.txt")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing tables...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                DoEvents
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                VCS_Table.ImportTableData CStr(obj_name), obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    
    '
    ' Import Data Macros
    '
    If (sObjects = "*" Or sObjects = "dm") Then
        obj_path = source_path & "tbldef\"
        fileName = Dir$(obj_path & "*.xml")
        If Len(fileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing Data Macros...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                DoEvents
                obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                'VCS_Table.ImportTableData CStr(obj_name), obj_path
                VCS_DataMacro.ImportDataMacros obj_name, obj_path
                obj_count = obj_count + 1
                fileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    
    '
    ' Import forms, reports, macros and modules
    '
    For Each obj_type In Split( _
        "forms|" & acForm & "," & _
        "reports|" & acReport & "," & _
        "macros|" & acMacro & "," & _
        "modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_num = val(obj_type_split(1))
        obj_path = source_path & obj_type_label & "\"
        
        '
        ' All objects or just one range?
        '
        If (sObjects = "*" Or sObjects = obj_type_label) Then
            fileName = Dir$(obj_path & "*.bas")
            If Len(fileName) > 0 Then
                Debug.Print VCS_String.PadRight("Importing " & obj_type_label & "...", 24);
                obj_count = 0
                Do Until Len(fileName) = 0
                    ' DoEvents no good idea!
                    obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
                    If obj_type_label = "modules" Then
                        ucs2 = False
                    Else
                        ucs2 = VCS_File.UsingUcs2
                    End If
                    If IsNotVCS(obj_name) Then
                        VCS_IE_Functions.ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
                        obj_count = obj_count + 1
                    Else
                        If ARCHIVE_MYSELF Then
                            MsgBox "Module " & obj_name & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
                        End If
                    End If
                    fileName = Dir$()
                Loop
                Debug.Print "[" & obj_count & "]"
            
            End If
        End If
    Next
    
    '
    ' Import Print Variables
    '
    If (sObjects = "*" Or sObjects = "reports") Then
        Debug.Print VCS_String.PadRight("Importing Print Vars...", 24);
        obj_count = 0
        
        obj_path = source_path & "reports\"
        fileName = Dir$(obj_path & "*.pv")
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            VCS_Report.ImportPrintVars obj_name, obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    '
    ' Import Relations
    '
    If (sObjects = "*" Or sObjects = "relations") Then
        Debug.Print VCS_String.PadRight("Importing Relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        fileName = Dir$(obj_path & "*.txt")
        Do Until Len(fileName) = 0
            DoEvents
            VCS_Relation.ImportRelation obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
End Sub
'
' Drop database objects as specified
'
Private Sub DeleteSource(Optional sObjects As String = "*")
On Error GoTo errorHandler
    Dim msg As String
    
    sObjects = LCase(sObjects)
    msg = "This action will delete all existing: " & vbCrLf & vbCrLf
    If (sObjects = "*") Then
        msg = msg & _
              Chr$(149) & " Tables" & vbCrLf & _
              Chr$(149) & " Forms" & vbCrLf & _
              Chr$(149) & " Macros" & vbCrLf & _
              Chr$(149) & " Modules" & vbCrLf & _
              Chr$(149) & " Queries" & vbCrLf & _
              Chr$(149) & " Reports" & vbCrLf
    Else
        msg = msg & Chr$(149) & " " & UCase(sObjects) & " only !" & vbCrLf
    End If
    msg = msg & vbCrLf & "Are you sure you want to proceed?"
              
    If MsgBox(msg, vbCritical + vbYesNo, "Import") <> vbYes Then
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    CloseFormsReports

    Debug.Print "Deleting Object(s): " & sObjects
    If (sObjects = "*" Or sObjects = "relations") Then
        Dim rel As DAO.Relation
        For Each rel In CurrentDb.Relations
            If Not (rel.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                    rel.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                CurrentDb.Relations.Delete (rel.name)
            End If
        Next
    End If

    If (sObjects = "*" Or sObjects = "queries") Then
        Dim dbObject As Object
        For Each dbObject In db.QueryDefs
            DoEvents
            If Left$(dbObject.name, 1) <> "~" Then
                'Debug.Print dbObject.Name
                db.QueryDefs.Delete dbObject.name
            End If
        Next
    End If
    
    If (sObjects = "*" Or sObjects = "tbldef") Then
        Dim td As DAO.TableDef
        For Each td In CurrentDb.TableDefs
            If Left$(td.name, 4) <> "MSys" And _
                Left$(td.name, 1) <> "~" Then
                CurrentDb.TableDefs.Delete (td.name)
            End If
        Next
    End If
    
    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME As Byte = 0
    Const OTID As Byte = 1

    For Each objType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(objType, "|")
        DoEvents
        '
        ' All objects or just one range?
        '
        If (sObjects = "*" Or sObjects = LCase(objTypeArray(OTNAME))) Then
            For Each doc In db.Containers(objTypeArray(OTNAME)).Documents
                DoEvents
                If (Left$(doc.name, 1) <> "~") And _
                   (IsNotVCS(doc.name)) Then
                    ' Debug.Print doc.Name
                    DoCmd.DeleteObject objTypeArray(OTID), doc.name
                End If
            Next
        End If
    Next
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.DeleteSource: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub
' Expose for use as function, can be called by query
Public Sub make()
    ImportProject
End Sub
'===================================================================================================================================
'-----------------------------------------------------------'
' Helper Functions - these should be put in their own files '
'-----------------------------------------------------------'
' Close all open forms.
Private Sub CloseFormsReports()
    On Error GoTo errorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).name
        DoEvents
    Loop
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub


'errno 457 - duplicate key (& item)
Public Function StrSetToCol(ByVal strSet As String, ByVal Delimiter As String) As Collection 'throws errors
    Dim strSetArray() As String
    Dim col As Collection
    
    Set col = New Collection
    strSetArray = Split(strSet, Delimiter)
    
    Dim item As Variant
    For Each item In strSetArray
        col.Add item, item
    Next
    
    Set StrSetToCol = col
End Function



