VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Structure Check"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "DBChecker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox QuickDbCheck 
      Caption         =   "Quick Db Check"
      Height          =   315
      Left            =   5070
      TabIndex        =   29
      Top             =   360
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin VB.TextBox Text4 
      Height          =   1500
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   3360
      Width           =   6405
   End
   Begin VB.CheckBox DefaultStructure 
      Caption         =   "Add Defaults To Structure"
      Height          =   345
      Left            =   0
      TabIndex        =   25
      Top             =   1530
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compare"
      Enabled         =   0   'False
      Height          =   1035
      Left            =   5070
      TabIndex        =   24
      Top             =   780
      Width           =   1755
      Begin VB.CheckBox Defaults 
         Caption         =   "Defaults"
         Enabled         =   0   'False
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   1005
      End
      Begin VB.CheckBox StoredProc 
         Caption         =   "Stored Procedures"
         Enabled         =   0   'False
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   420
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox TriggerCheck 
         Caption         =   "Triggers"
         Enabled         =   0   'False
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   660
         Width           =   1605
      End
   End
   Begin VB.CheckBox AddDefaults 
      Caption         =   "Add Default Values To Nulls"
      Height          =   345
      Left            =   5100
      TabIndex        =   9
      Top             =   1860
      Width           =   1605
   End
   Begin VB.TextBox DbDrive 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "C"
      Top             =   390
      Width           =   435
   End
   Begin VB.CommandButton AddDbs 
      Caption         =   ">>"
      Height          =   285
      Left            =   2910
      TabIndex        =   22
      Top             =   1320
      Width           =   465
   End
   Begin VB.CommandButton DeleteDbs 
      Caption         =   "<<"
      Height          =   285
      Left            =   2910
      TabIndex        =   21
      Top             =   1620
      Width           =   465
   End
   Begin VB.ListBox DbList1 
      Height          =   1620
      Left            =   1260
      TabIndex        =   4
      Top             =   720
      Width           =   1605
   End
   Begin VB.CommandButton DelDb 
      Caption         =   "<"
      Height          =   285
      Left            =   2910
      TabIndex        =   20
      Top             =   1020
      Width           =   465
   End
   Begin VB.ListBox DbList 
      Height          =   1620
      Left            =   3450
      TabIndex        =   5
      Top             =   720
      Width           =   1605
   End
   Begin VB.CommandButton AddDb 
      Caption         =   ">"
      Height          =   285
      Left            =   2910
      TabIndex        =   19
      Top             =   720
      Width           =   465
   End
   Begin VB.TextBox Text3 
      Height          =   2580
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   5130
      Width           =   6405
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Prepare Target Database for Replication(this option takes time)"
      Height          =   435
      Left            =   60
      TabIndex        =   17
      Top             =   990
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Database &Check"
      Height          =   300
      Left            =   5070
      TabIndex        =   10
      Top             =   2250
      Width           =   1740
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   390
      Width           =   1605
   End
   Begin VB.CommandButton BrowseBtn 
      Caption         =   "&Browse"
      Height          =   300
      Left            =   5070
      TabIndex        =   1
      Top             =   60
      Width           =   1740
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1230
      TabIndex        =   0
      Top             =   60
      Width           =   3780
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   4410
      Left            =   6585
      TabIndex        =   11
      Top             =   3330
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   7779
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Summerized Info"
      Height          =   195
      Left            =   75
      TabIndex        =   28
      Top             =   3165
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Detailed Info"
      Height          =   195
      Left            =   90
      TabIndex        =   27
      Top             =   4935
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Temporary Db Drive"
      Height          =   195
      Left            =   3000
      TabIndex        =   23
      Top             =   435
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Database Name"
      Height          =   225
      Left            =   15
      TabIndex        =   16
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label Label3 
      Caption         =   "Server Name"
      Height          =   225
      Left            =   15
      TabIndex        =   15
      Top             =   420
      Width           =   1110
   End
   Begin VB.Label ProgressTitle 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Script Location"
      Height          =   225
      Left            =   15
      TabIndex        =   13
      Top             =   90
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   90
      TabIndex        =   12
      Top             =   2520
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' "GoTo" may not be compatible with structured

 "GoTo" may not be compatible with structured
 programming concepts.
Option Explicit

Private PassToPass As String
Private passedTesting As Boolean

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias _
                          "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                          lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, _
                          lpTotalNumberOfFreeBytes As Currency) As Long

Private Sub AddDb_Click()

    If DbList1.ListIndex = -1 Then
        Exit Sub '>---> Bottom
    End If
    If Trim$(DbList1.List(DbList1.ListIndex)) = "" Then
        Exit Sub '>---> Bottom
    End If
    DbList.AddItem Trim$(DbList1.List(DbList1.ListIndex))
    DbList1.RemoveItem DbList1.ListIndex

End Sub

Private Sub AddDbs_Click()

  Dim i As Integer

    For i = DbList1.ListCount - 1 To 0 Step -1
        DbList.AddItem DbList1.List(i)
        DbList1.RemoveItem i
    Next i

End Sub

Private Sub AddDefaults_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub AddingColumns(cnSource As rdoConnection, cnDestination As rdoConnection, Sourcetb As rdoTable, DestinationTb As rdoTable, StartingRow As Integer, errorstring As String)

  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim i As Integer
  Dim TbName As String
  Dim colType As String
  Dim AcceptNulls As String
  Dim ColDef As String
  Dim j As Integer
  Dim DefaultCol As String

    On Error GoTo AddingColumns_Error

    '---------------------

    TbName = DestinationTb.Name
    Set cpw = cnSource.CreateQuery("", "exec sp_columns @table_name='" & TbName & "'")
    Set tb = cpw.OpenResultset(2)
    For i = (Sourcetb.rdoColumns.Count - StartingRow + 1) To Sourcetb.rdoColumns.Count
        tb.MoveFirst
        Do While Not tb.EOF
            If tb!ordinal_position = i Then
                If Val(tb!nullable) = 1 Then
                    AcceptNulls = "Null"
                  Else 'NOT VAL(TB!NULLABLE)...
                    AcceptNulls = "Not Null"
                End If
                If Not IsNull(tb!COLUMN_DEF) Then
                    DefaultCol = " Default " & tb!COLUMN_DEF
                  Else 'NOT NOT...
                    DefaultCol = ""
                End If
                Select Case LCase$(Trim$(tb!type_name))
                  Case "numeric"
                    cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & " " & DefaultCol
                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & "  Added In " & TbName
                  Case "datetime", "bit", "int", "smallint", "real"
                    cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & " " & AcceptNulls & " " & DefaultCol
                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & " " & AcceptNulls & " Added In " & TbName
                  Case Else
                    cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & ") " & AcceptNulls & " " & DefaultCol
                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type : " & tb!type_name & "(" & tb!Precision & ") " & AcceptNulls & " Added In " & TbName
                End Select
                Exit Do '>---> Loop
            End If
            tb.MoveNext
        Loop
    Next i

    On Error GoTo 0

Exit Sub

AddingColumns_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddingColumns of Form Form1"

End Sub

Private Sub AddingOneColumn(cnSource As rdoConnection, cnDestination As rdoConnection, Sourcetb As rdoTable, DestinationTb As rdoTable, errorstring As String, ColName As String)

  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim TbName As String
  Dim colType As String
  Dim AcceptNulls As String
  Dim j As Integer
  Dim DefaultCol As String
  Dim TSQL As String

    On Error GoTo AddingOneColumn_Error

    '---------------------

    TbName = DestinationTb.Name
    Set cpw = cnSource.CreateQuery("", "exec sp_columns @table_name='" & TbName & "'")
    Set tb = cpw.OpenResultset(2)
    Do While Not tb.EOF
        If UCase$(tb!COLUMN_NAME) = UCase$(ColName) Then
            If Val(tb!nullable) = 1 Then
                AcceptNulls = "Null"
              Else 'NOT VAL(TB!NULLABLE)...
                AcceptNulls = "Not Null"
            End If
            If Not IsNull(tb!COLUMN_DEF) Then
                DefaultCol = " Default " & tb!COLUMN_DEF
              Else 'NOT NOT...
                DefaultCol = ""
            End If
            Select Case LCase$(Trim$(tb!type_name))
              Case "numeric"
                cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & " " & DefaultCol
                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & "  Added In " & TbName
              Case "datetime", "bit", "int", "smallint"
                cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & " " & AcceptNulls & " " & DefaultCol
                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & " " & AcceptNulls & " Added In " & TbName
              Case "numeric() identity"
                If Text3.ForeColor <> vbRed Then
                    Text3.ForeColor = vbBlue
                End If
                Text3.Refresh
                Text3 = Chr$(9) & "this might take some time " & vbCrLf & Text3
                Text3 = Chr$(9) & "Now Adding A PrimaryKey to '" & UCase$(TbName) & "'" & vbCrLf & Text3
                TSQL = "ALTER TABLE " & TbName & " ADD PrimaryKey " & Left$(tb!type_name, InStr(tb!type_name, "(") - 1) & "(" & tb!Precision & "," & tb!Scale & ") " & " identity not for replication " & AcceptNulls & Chr$(13)
                TSQL = TSQL & " ALTER TABLE [dbo].[" & TbName & "] WITH NOCHECK ADD CONSTRAINT [PK_" & TbName & "] PRIMARY KEY  NONCLUSTERED ([PrimaryKey]) WITH  FILLFACTOR = 90  ON [PRIMARY]"
                cnDestination.Execute TSQL
                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & "  Added In " & TbName
                If Text3.ForeColor <> vbRed Then
                    Text3.ForeColor = vbBlack
                End If
              Case "real"
                cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & " " & AcceptNulls & " " & DefaultCol
                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type: " & tb!type_name & " " & AcceptNulls & " Added In " & TbName
              Case Else
                cnDestination.Execute "ALTER TABLE " & TbName & " ADD " & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & ") " & AcceptNulls & " " & DefaultCol
                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Type : " & tb!type_name & "(" & tb!Precision & ") " & AcceptNulls & " Added In " & TbName
            End Select
            Exit Do '>---> Loop
        End If
        tb.MoveNext
    Loop

    On Error GoTo 0

Exit Sub

AddingOneColumn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddingOneColumn of Form Form1"

End Sub

Private Sub AddTable(cnSource As rdoConnection, cnDestination As rdoConnection, Sourcetb As rdoTable, DestinationTb As rdoTable)

  Dim i As Integer
  Dim tmp As String
  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim AcceptNulls As String
  Dim DefaultCol As String

    On Error GoTo AddTable_Error

    '------------------
    Set cpw = cnSource.CreateQuery("", "exec sp_columns @table_name='" & Sourcetb.Name & "'")
    Set tb = cpw.OpenResultset(2)
    tmp = "Create Table " & Sourcetb.Name & " ( "
    Do While Not tb.EOF
        If Val(tb!nullable) = 1 Then
            AcceptNulls = "Null"
          Else 'NOT VAL(TB!NULLABLE)...
            AcceptNulls = "Not Null"
        End If
        If Not IsNull(tb!COLUMN_DEF) Then
            DefaultCol = " Default " & tb!COLUMN_DEF
          Else 'NOT NOT...
            DefaultCol = ""
        End If
        Select Case LCase$(Trim$(tb!type_name))
          Case "numeric() identity"
            tmp = tmp & tb!COLUMN_NAME & " " & Left$(tb!type_name, InStr(tb!type_name, "(") - 1) & "(" & tb!Precision & "," & tb!Scale & ") " & " identity " & AcceptNulls & " " & DefaultCol
          Case "decimal() identity"
            tmp = tmp & tb!COLUMN_NAME & " " & Left$(tb!type_name, InStr(tb!type_name, "(") - 1) & "(" & tb!Precision & "," & tb!Scale & ") " & " identity " & AcceptNulls & " " & DefaultCol
          Case "numeric"
            tmp = tmp & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & "," & tb!Scale & ") " & AcceptNulls & " " & DefaultCol
          Case "datetime", "bit", "int", "smallint"
            tmp = tmp & tb!COLUMN_NAME & " " & tb!type_name & " " & AcceptNulls & " " & DefaultCol
          Case "real", "text"
            tmp = tmp & tb!COLUMN_NAME & " " & tb!type_name & " " & AcceptNulls & " " & DefaultCol
          Case Else
            tmp = tmp & tb!COLUMN_NAME & " " & tb!type_name & "(" & tb!Precision & ") " & AcceptNulls & " " & DefaultCol
        End Select
        tmp = tmp & ","
        tb.MoveNext
    Loop
    tmp = Left$(tmp, Len(tmp) - 1) & ")"
    cnDestination.Execute tmp

    On Error GoTo 0

Exit Sub

AddTable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddTable of Form Form1"

End Sub

Private Sub BrowseBtn_Click()

    CD1.InitDir = "C:\"
    CD1.Filter = "Script Files (*.sql)|*.sql"
    CD1.FilterIndex = 1
    CD1.ShowOpen
    If CD1.FileTitle <> "" Then
        Text1 = CD1.FileName
    End If

End Sub

Private Sub CheckDefaults(cnSource As rdoConnection, cnDest As rdoConnection, TbName As String, ColumnName As String, errorstring As String)

  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim cpw1 As rdoQuery
  Dim tb1 As rdoResultset

    On Error GoTo CheckDefaults_Error

    Set cpw = cnSource.CreateQuery("", "exec sp_columns @table_name='" & TbName & "',@column_name='" & ColumnName & "'")
    Set tb = cpw.OpenResultset(2)

    Set cpw1 = cnDest.CreateQuery("", "exec sp_columns @table_name='" & TbName & "',@column_name='" & ColumnName & "'")
    Set tb1 = cpw1.OpenResultset(2)

    If Not (tb.EOF And tb.BOF) Then
        If Not (tb1.EOF And tb1.BOF) Then
            If Not IsNull(tb!COLUMN_DEF) Then
                If Not IsNull(tb1!COLUMN_DEF) Then
                    If tb!COLUMN_DEF <> tb1!COLUMN_DEF Then
                        errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & ColumnName & " : Different Defaults"
                        'error 'both have defaults but they're different
                    End If
                  Else 'NOT NOT...
                    cnDest.Execute "ALTER TABLE " & TbName & " WITH NOCHECK add CONSTRAINT [DF_" & TbName & "_" & tb!COLUMN_NAME & "] DEFAULT " & tb!COLUMN_DEF & " FOR [" & tb!COLUMN_NAME & "] "
                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & tb!COLUMN_NAME & " Default: " & tb!COLUMN_DEF & " Added In " & TbName
                    'error 'source has default, destination does not
                End If
              Else 'NOT NOT...
                If Not IsNull(tb1!COLUMN_DEF) Then
                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & ColumnName & " : No Default For source"
                    'error  'source Col does not have any defaults, destination has
                End If
            End If
        End If
    End If

    On Error GoTo 0

Exit Sub

CheckDefaults_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckDefaults of Form Form1"

End Sub

Private Sub checkingIndexes(cnSource As rdoConnection, cnDestination As rdoConnection, Sourcetb As rdoTable, DestinationTb As rdoTable, errorstring As String, errOld As String)

  Dim sql As String

  Dim SrcCpw As rdoQuery
  Dim SrcTb As rdoResultset

  Dim DestCpw As rdoQuery
  Dim DestTb As rdoResultset

  Dim TextLine As String
  Dim tb As rdoResultset
  Dim Done As Boolean

    On Error GoTo checkingIndexes_Error

    '-------------------
    On Error Resume Next
        Done = False
        sql = "SELECT TABLE_NAME = sysobjects.name,"
        sql = sql & "INDEX_NAME = sysindexes.name, INDEX_ID = indid "
        sql = sql & "FROM sysindexes INNER JOIN sysobjects ON sysobjects.id = sysindexes.id "
        sql = sql & "Where sysobjects.Name = "

        Set SrcCpw = cnSource.CreateQuery("", sql & "'" & Trim$(Sourcetb.Name) & "'")
        Set SrcTb = SrcCpw.OpenResultset(2)

        If SrcTb.RowCount <> 0 Then
            Do While Not SrcTb.EOF
                Set DestCpw = cnDestination.CreateQuery("", sql & "'" & Trim$(Sourcetb.Name) & "' and sysindexes.name= '" & Trim$(SrcTb!INDEX_NAME) & "'")
                Set DestTb = DestCpw.OpenResultset(2)
                If DestTb.RowCount = 0 Then
                    If errOld = "" Then
                        errorstring = Chr$(13) & Chr$(10) & Chr$(9) & Sourcetb.Name & " : " & SrcTb!INDEX_NAME & " Index Not Found "
                      Else 'NOT ERROLD...
                        errorstring = Chr$(13) & Chr$(10) & SrcTb!INDEX_NAME & " Index Not Found "
                    End If

                    Close #1
                    Open Trim(Text1) For Input As #1
                    sql = ""
                    Do While Not EOF(1)
                        Line Input #1, TextLine
                        If InStr(1, UCase$(TextLine), "INDEX") > 0 And InStr(1, UCase$(TextLine), UCase$(Sourcetb.Name)) > 0 And InStr(1, UCase$(TextLine), UCase$(SrcTb!INDEX_NAME)) > 0 Then
                            Do While Not (InStr(1, UCase$(TextLine), "GO") > 0)
                                sql = sql & " " & TextLine
                                Line Input #1, TextLine
                            Loop
                            Set tb = cnDestination.OpenResultset(sql, 2, rdConcurRowVer)  ', rdAsyncEnable
                            Done = True
                            If InStr(1, UCase$(TextLine), "GO") > 0 Then
                                sql = ""
                                '------------
                            End If
                        End If
                    Loop
                    If Not Done Then
                        Close #1
                        Open Trim(Text1) For Input As #1
                        sql = ""
                        Do While Not EOF(1)
                            Line Input #1, TextLine
                            If InStr(1, UCase$(TextLine), "ALTER TABLE") > 0 And InStr(1, UCase$(TextLine), UCase$(Sourcetb.Name)) > 0 Then
                                sql = TextLine
                                If Not EOF(1) Then
                                    Line Input #1, TextLine
                                    If InStr(1, UCase$(TextLine), UCase$(SrcTb!INDEX_NAME)) > 0 Then
                                        Do While Not (InStr(1, UCase$(TextLine), "GO") > 0)
                                            sql = sql & " " & TextLine
                                            Line Input #1, TextLine
                                        Loop
                                        Set tb = cnDestination.OpenResultset(sql, 2, rdConcurRowVer)  ', rdAsyncEnable
                                        Done = True
                                        If InStr(1, UCase$(TextLine), "GO") > 0 Then
                                            Exit Do '>---> Loop
                                        End If
                                    End If
                                End If
                            End If
                        Loop
                    End If
                End If
                SrcTb.MoveNext
            Loop
        End If

    On Error GoTo 0

Exit Sub

checkingIndexes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkingIndexes of Form Form1"

End Sub

Private Function CheckPro(PrName As String, ConString As String, MyType As String) As String

  Dim Con1 As New rdoConnection
  Dim DestCpw1 As rdoQuery
  Dim DestTb1 As rdoResultset

    On Error GoTo CheckPro_Error

    Con1.Connect = ConString
    Con1.CursorDriver = rdUseOdbc
    Con1.EstablishConnection rdDriverNoPrompt

    If MyType = "S" Then
        Set DestCpw1 = Con1.CreateQuery("", "select * from sysobjects where  name='" & PrName & "' and OBJECTPROPERTY(id, N'IsProcedure') = 1")
        Set DestTb1 = DestCpw1.OpenResultset(2)
        If (DestTb1.EOF And DestTb1.BOF) Then
            CheckPro = Chr$(13) & Chr$(10) & Chr$(9) & " Could Not add Stored Procedure: " & PrName
        End If
      Else 'NOT MYTYPE...
        Set DestCpw1 = Con1.CreateQuery("", "select * from sysobjects where  name='" & PrName & "' and OBJECTPROPERTY(id, N'IsTrigger') = 1")
        Set DestTb1 = DestCpw1.OpenResultset(2)
        If (DestTb1.EOF And DestTb1.BOF) Then
            CheckPro = Chr$(13) & Chr$(10) & Chr$(9) & " Could Not add Trigger: " & PrName
        End If
    End If
    Con1.Close

    On Error GoTo 0

Exit Function

CheckPro_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckPro of Form Form1"

End Function

Private Sub ColDefaults(cn1 As rdoConnection, TbName As String, ColName As String, colType As String, errorstring As String)

  Dim cpw1 As rdoQuery
  Dim tb1 As rdoResultset

    On Error GoTo ColDefaults_Error

    Select Case LCase$(Trim$(colType))
      Case "2", "3", "4", "5", "6", "7", "8", "-2", "-3", "-4", "-5", "-6", "-7"
        'Case "numeric", "bit", "int", "smallint"
        cn1.Execute "update " & TbName & " set " & ColName & " =0 where " & ColName & " is null"
        errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & " Added Default to " & ColName
      Case "9", "10", "11"
        'Case "datetime"
        cn1.Execute "update " & TbName & " set " & ColName & " =getdate() where " & ColName & " is null"
        errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & " Added Default to " & ColName
      Case "12", "1", "-1"
        'Case "varchar", "char"
        cn1.Execute "update " & TbName & " set " & ColName & " =' ' where " & ColName & " is null"
        errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & " Added Default to " & ColName
    End Select

    On Error GoTo 0

Exit Sub

ColDefaults_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ColDefaults of Form Form1"

End Sub

Private Sub Command2_Click()

    passedTesting = False
    If QuickDbCheck.Value = 1 Then
        QuickDbChecker
      Else 'NOT QUICKDBCHECK.VALUE...
        NameCheck
    End If

    MsgBox "Operation is successfully Completed", vbInformation

End Sub

Private Sub DbDrive_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub DbList1_DblClick()

    AddDb_Click

End Sub

Private Sub DbList_DblClick()

    DelDb_Click

End Sub

Private Sub Defaults_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub DelDb_Click()

    If DbList.ListIndex = -1 Then
        Exit Sub '>---> Bottom
    End If
    DbList1.AddItem DbList.List(DbList.ListIndex)
    DbList.RemoveItem DbList.ListIndex

End Sub

Private Sub DeleteDbs_Click()

  Dim i As Integer

    For i = DbList.ListCount - 1 To 0 Step -1
        DbList1.AddItem DbList.List(i)
        DbList.RemoveItem i
    Next i

End Sub

Private Sub DeletePrimaryKeys(cn1 As rdoConnection, tablename As String)

  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim TbName As String
  Dim colType As String
  Dim AcceptNulls As String
  Dim j As Integer, K As Integer

  '---------------------

    On Error GoTo ERRORINDELETEpRIMARYkEY

    Set cpw = cn1.CreateQuery("", "exec sp_columns @table_name='" & tablename & "'")
    Set tb = cpw.OpenResultset(2)
    Do While Not tb.EOF
        If LCase$(Trim$(tb!type_name)) = "numeric() identity" Then
            Text3 = Chr$(9) & "Clearing PrimaryKeys from '" & UCase$(tablename) & "'" & vbCrLf & Text3
            K = 1
            cn1.Execute "Alter table " & tablename & " drop constraint pk_" & tablename
            K = 2
            cn1.Execute "Alter table " & tablename & " drop column " & tb!COLUMN_NAME
            cn1.rdoTables.Refresh
        End If
        tb.MoveNext
    Loop

Exit Sub

ERRORINDELETEpRIMARYkEY:
    If K = 2 Then
        Text3.ForeColor = vbRed
        Text3 = Chr$(9) & "Could not Clean PrimaryKey from " & UCase$(tablename) & vbCrLf & Text3
        Text3 = Chr$(9) & "YOU SHOULD CLEAN IT YOURSELF AND THEN TRY AGAIN CHECK DATABASE" & vbCrLf & Text3
    End If
    Resume Next

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    DoEvents
    Text1 = App.Path
    If Right$(App.Path, 1) <> "\" Then
        Text1 = Text1 & "\"
    End If

    Text1 = Text1 & "Script.sql"
    Text2 = ""
    If GetDiskInfo("c:\") > 3145728 Then
        DbDrive = "C"
      ElseIf GetDiskInfo("D:\") > 3145728 Then 'NOT GETDISKINFO("C:\")...
        DbDrive = "D"
    End If

    Text2_LostFocus
    AddDbs_Click
    Screen.MousePointer = vbNormal

End Sub

Private Function GetDiskInfo(DriveLetter As String) As Currency

  '

  Dim r As Long
  Dim BytesFreeToCalller As Currency
  Dim TotalBytes As Currency
  Dim TotalFreeBytes As Currency
  Dim TotalBytesUsed As Currency
  Dim TNB As Double
  Dim TFB As Double
  Dim FreeBytes As Long

  Dim DLetter As String
  Dim spaceInt As Integer

    On Error GoTo GetDiskInfo_Error

    'If there is a space at the end of DriveLetter, remove it
    spaceInt = InStr(DriveLetter, " ")
    If spaceInt > 0 Then
        DriveLetter = Left$(DriveLetter, spaceInt - 1)
    End If

    'if there is not a "\" at the end of DriveLetter, add it
    If Right$(DriveLetter, 1) <> "\" Then
        DriveLetter = DriveLetter & "\"
    End If
    DLetter = Left$(UCase$(DriveLetter), 1)

    'get the drive's disk parameters
    Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    'show the results, multiplying the returned
    'value by 10000 to adjust for the 4 decimal
    'places that the currency data type returns.

    GetDiskInfo = BytesFreeToCalller * 10000

    On Error GoTo 0

Exit Function

GetDiskInfo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDiskInfo of Form Form1"

End Function

Private Sub NameCheck()

  Dim TextLine As String
  Dim tmp As String
  Dim cn As New rdoConnection
  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim TbSet As rdoTable
  Dim coltb As rdoColumn
  Dim TempCol As rdoColumn
  Dim cn1 As New rdoConnection
  Dim TempTable As rdoTable
  Dim i As Integer
  Dim j As Integer
  Dim a As Integer
  Dim Till As Integer
  Dim found As Boolean
  Dim AnyError As Boolean
  Dim errorstring As String
  Dim ColFound As Boolean
  Dim inderrorstring As String
  Dim dbs As Integer
  Dim SummaryErr As String
  Dim DetailErr As String
  '-------------------------

    Text3 = ""
    ProgressBar1.Value = 0
    ProgressTitle.Caption = ""
    Label1.Caption = ""
    Screen.MousePointer = 11
    If Trim$(Text1) = "" Then
        MsgBox "Enter The Script Location!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Trim$(Text2) = "" Then
        MsgBox "Enter the name of the server!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If DbList.ListCount = 0 Then
        MsgBox "Enter The name(s) Of The Database(s) You Want Checked!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Len(DbDrive) = 0 Then
        MsgBox "Enter The Drive of The Temporary Database!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Dir(Trim$(Text1)) = "" Then
        MsgBox "Script Not Found!", vbCritical
        Exit Sub '>---> Bottom
    End If
    passedTesting = True
    Close #1
    Open Trim(Text1) For Input As #1

    cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & DbList.List(0) & ";DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    tmp = "USE master" & Chr$(13) & Chr$(10)
    tmp = tmp & "select * from sysdatabases where name = 'TempStruc'"
    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)
    If tb.EOF And tb.BOF Then
        If Dir(Trim$(DbDrive) & ":\TempStruc.mdf") <> "" Then
            Kill Trim$(DbDrive) & ":\TempStruc.mdf"
        End If
        If Dir(Trim$(DbDrive) & ":\TempStruclogs.ldf") <> "" Then
            Kill Trim$(DbDrive) & ":\TempStruclogs.ldf"
        End If
      Else 'NOT TB.EOF...
        tmp = "USE master" & Chr$(13) & Chr$(10)
        tmp = "if exists (select * from sysdatabases where name = 'TempStruc') drop DATABASE TempStruc" & Chr$(13) & Chr$(10)
        Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
    End If
    DoEvents
    '------------Create Temporary Database---------------'
    tmp = "USE master" & Chr$(13) & Chr$(10)
    tmp = tmp & "CREATE DATABASE TempStruc" & Chr$(13) & Chr$(10)
    tmp = tmp & "ON ( NAME = TempStruc_dat, FILENAME = '" & Trim$(DbDrive) & ":\TempStruc.mdf',SIZE = 10,MAXSIZE = 50,FILEGROWTH = 5)"
    tmp = tmp & "LOG ON ( NAME = 'TempStruc_log', FILENAME = '" & Trim$(DbDrive) & ":\TempStruclogs.ldf',SIZE = 5MB, MAXSIZE = 25MB,  FILEGROWTH = 5MB )"

    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
    cn.Close
    Text3 = "Temporary Database Created" & Text3 & Chr$(13) & Chr$(10)

    '----------Connect to Temporary Database------------'
    cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=TempStruc;DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    '----------Running Script of Temporary Database------'
    tmp = ""
    ProgressTitle.Caption = "Creating Temporary Structure"
    ProgressTitle.Refresh
    Do While Not EOF(1)
        Label1 = Str$(Val(Label1) + 1) & " lines in Script"
        Label1.Refresh
        Line Input #1, TextLine
        If Len(TextLine) > 1 Then
            If (Left$(UCase$(TextLine), 2)) <> "GO" Then
                tmp = tmp & Chr$(13) & Chr$(10) & TextLine

              Else 'NOT (LEFT$(UCASE$(TEXTLINE),...
                On Error Resume Next
                    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
                    tmp = ""
                End If
              Else 'NOT LEN(TEXTLINE)...
                tmp = tmp & Chr$(13) & Chr$(10) & TextLine
            End If
        Loop
        DoEvents
    On Error GoTo 0
    tb.Close
    '-------------'
    Close #1
    Text3 = "Database Structure Check Result" & Text3
    Text3 = "-----------------------------------------------------" & Chr$(13) & Chr$(10) & Text3
    ProgressTitle.Caption = "Checking Structure"
    ProgressTitle.Refresh

    For dbs = 0 To DbList.ListCount - 1
        ProgressTitle.Caption = "Structure Check For Database: " & DbList.List(dbs) & ""
        ProgressTitle.Refresh
        ProgressBar1.Value = 0
        ProgressBar1.Max = cn.rdoTables.Count
        DoEvents
        cn1.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & DbList.List(dbs) & ";DSN='';"
        cn1.CursorDriver = rdUseOdbc
        cn1.EstablishConnection rdDriverNoPrompt
        cn1.QueryTimeout = 3000
        cn1.rdoTables.Refresh
        found = False

        For Each TbSet In cn.rdoTables
            Set cpw = cn.CreateQuery("", "EXEC sp_tables " & TbSet.Name)
            Set tb = cpw.OpenResultset(2)
            If UCase$(TbSet.Type) = "TABLE" And UCase$(tb!table_type) = "TABLE" Then
                For Each TempTable In cn1.rdoTables
                    errorstring = ""
                    DetailErr = ""
                    SummaryErr = ""
                    If UCase$(TempTable.Type) = "TABLE" Then
                        If UCase$(TbSet.Name) = UCase$(TempTable.Name) Then
                            If Check1.Value = 1 Then
                                DeletePrimaryKeys cn1, TempTable.Name
                            End If
                            found = True
                            For i = 0 To TbSet.rdoColumns.Count - 1
                                ColFound = False
                                For j = 0 To TempTable.rdoColumns.Count - 1
                                    DoEvents
                                    If UCase$(TbSet.rdoColumns(i).Name) = UCase$(TempTable.rdoColumns(j).Name) Then  ''Name Comparison
                                        ColFound = True
                                        DoEvents
                                        If TbSet.rdoColumns(i).Type <> TempTable.rdoColumns(j).Type Then
                                            DetailErr = DetailErr & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Type Mismatch"
                                            DoEvents
                                          Else 'NOT TBSET.RDOCOLUMNS(I).TYPE...
                                            If TbSet.rdoColumns(i).Size <> TempTable.rdoColumns(j).Size Then
                                                DetailErr = DetailErr & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Size Mismatch"
                                                DoEvents
                                              Else 'NOT TBSET.RDOCOLUMNS(I).SIZE...
                                                If TbSet.rdoColumns(i).Required <> TempTable.rdoColumns(j).Required Then
                                                    DetailErr = DetailErr & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Null"
                                                    DoEvents
                                                  Else 'NOT TBSET.RDOCOLUMNS(I).REQUIRED...
                                                    If Defaults = 1 Then
                                                        CheckDefaults cn, cn1, TbSet.Name, TbSet.rdoColumns(i).Name, errorstring
                                                        DoEvents
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If AddDefaults = 1 Then
                                            If TempTable.rdoColumns(j).Required = False Then
                                                ColDefaults cn1, TempTable.Name, TempTable.rdoColumns(j).Name, TempTable.rdoColumns(j).Type, errorstring
                                                DoEvents
                                            End If
                                        End If

                                        Exit For '>---> Next
                                    End If
                                Next j
                                If ColFound = False Then
                                    AddingOneColumn cn, cn1, TbSet, TempTable, errorstring, TbSet.rdoColumns(i).Name
                                    DoEvents
                                End If
                            Next i
                            If errorstring <> "" Then
                                Text4 = TempTable.Name & errorstring & Chr$(13) & Chr$(10) & Text4
                                Text3 = TempTable.Name & errorstring & Chr$(13) & Chr$(10) & Text3
                            End If
                            If DetailErr <> "" Then
                                Text3 = TempTable.Name & DetailErr & Chr$(13) & Chr$(10) & Text3
                            End If
                            Exit For '''the correct table is compared '>---> Next
                        End If
                    End If
                Next TempTable
            End If
            If found = False And UCase$(tb!table_type) = "TABLE" Then
                tmp = "cn.CreateQuery "
                AddTable cn, cn1, TbSet, TempTable
                Text3 = "<<< " & TbSet.Name & " Added To Original Database >>>" & Chr$(13) & Chr$(10) & Text3
                Text4 = "<<< " & TbSet.Name & " Added To Original Database >>>" & Chr$(13) & Chr$(10) & Text4
            End If
            found = False
            Text3.Refresh
            '--------checking Indexes----------'
            inderrorstring = ""
            checkingIndexes cn, cn1, TbSet, TempTable, inderrorstring, errorstring
            If inderrorstring <> "" Then
                Text3 = inderrorstring & Chr$(13) & Chr$(10) & Text3
                Text4 = inderrorstring & Chr$(13) & Chr$(10) & Text4
                Text3.Refresh
            End If
            '----------------------------------'
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            tb.Close
        Next TbSet
        If TriggerCheck = 1 Then
            Triggers cn1, errorstring
            Text3 = errorstring & Chr$(13) & Chr$(10) & Text3
            Text4 = errorstring & Chr$(13) & Chr$(10) & Text4
        End If
        If StoredProc = 1 Then
            StoredProcedures cn1, errorstring
            Text3 = errorstring & Chr$(13) & Chr$(10) & Text3
            Text4 = errorstring & Chr$(13) & Chr$(10) & Text4
        End If
        ProgressTitle.Caption = "Structure Check For Database: " & DbList.List(dbs) & ""
        ProgressTitle.Refresh
        Text3 = "  Results of Structure Check For Database: " & DbList.List(dbs) & Chr$(13) & Chr$(10) & Text3
        Text4 = "  Summerized Results of Structure Check For Database: " & DbList.List(dbs) & Chr$(13) & Chr$(10) & Text4
        DoEvents
        cn1.Close
    Next dbs
    ProgressTitle.Caption = "Structure Check Over"
    ProgressTitle.Refresh

    Screen.MousePointer = 0

End Sub

Private Sub PositionCheck()

  Dim TextLine As String
  Dim tmp As String
  Dim cn As New rdoConnection
  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim cn1 As New rdoConnection
  Dim TbSet As rdoTable
  Dim TempTable As rdoTable
  Dim coltb As rdoColumn
  Dim TempCol As rdoColumn
  Dim i As Integer
  Dim a As Integer
  Dim Till As Integer
  Dim found As Boolean
  Dim AnyError As Boolean
  Dim errorstring As String
  '-------------------------

    Text3 = ""
    ProgressBar1.Value = 0
    ProgressTitle.Caption = ""
    Label1.Caption = ""
    Screen.MousePointer = 11
    If Trim$(Text1) = "" Then
        MsgBox "Enter The Script Location!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Trim$(Text2) = "" Then
        MsgBox "Enter the name of the server!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If
    If DbList.ListCount = 0 Then
        MsgBox "Enter the name of the Database on the Server!", vbInformation, "Structure Check"
        Exit Sub '>---> Bottom
    End If

    Open Trim(Text1) For Input As #1

    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    '------------Create Temporary Database---------------'
    tmp = "USE master" & Chr$(13) & Chr$(10)
    tmp = tmp & "if exists (select * from sysdatabases where name = 'TempStruc') drop DATABASE TempStruc" & Chr$(13) & Chr$(10)
    tmp = tmp & "CREATE DATABASE TempStruc" & Chr$(13) & Chr$(10)
    tmp = tmp & "ON ( NAME = TempStruc_dat, FILENAME = 'd:\TempStruc.mdf',SIZE = 10,MAXSIZE = 50,FILEGROWTH = 5)"
    tmp = tmp & "LOG ON ( NAME = 'TempStruc_log', FILENAME = 'd:\TempStruclogs.ldf',SIZE = 5MB, MAXSIZE = 25MB,  FILEGROWTH = 5MB )"
    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
    cn.Close
    Text3 = "Temporary Database Created" & Text3 & Chr$(13) & Chr$(10)

    '----------Connect to Temporary Database------------'
    cn.Connect = "uid=sa;pwd=;server=" & Text2 & ";driver={SQL Server};database=TempStruc;DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    '----------Running Script of Temporary Database------'
    tmp = ""
    ProgressTitle.Caption = "Creating Temporary Structure"
    ProgressTitle.Refresh
    Do While Not EOF(1)
        Label1 = Str$(Val(Label1) + 1) & " lines in Script"
        Label1.Refresh
        Line Input #1, TextLine
        If Len(TextLine) > 1 Then
            If (Left$(UCase$(TextLine), 2)) <> "GO" Then
                tmp = tmp & Chr$(13) & Chr$(10) & TextLine

              Else 'NOT (LEFT$(UCASE$(TEXTLINE),...
                Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
                tmp = ""
            End If
          Else 'NOT LEN(TEXTLINE)...
            tmp = tmp & Chr$(13) & Chr$(10) & TextLine
        End If
    Loop

    tb.Close
    '-------------'
    Close #1
    cn1.CursorDriver = rdUseOdbc
    cn1.EstablishConnection rdDriverNoPrompt

    Text3 = "Database Structure Check Result" & Text3
    Text3 = "-----------------------------------------------------" & Chr$(13) & Chr$(10) & Text3
    ProgressTitle.Caption = "Checking Structure"
    ProgressTitle.Refresh
    ProgressBar1.Max = cn.rdoTables.Count
    For Each TbSet In cn.rdoTables
        Set cpw = cn.CreateQuery("", "EXEC sp_tables " & TbSet.Name & "")
        Set tb = cpw.OpenResultset(2)
        If UCase$(TbSet.Type) = "TABLE" And UCase$(tb!table_type) = "TABLE" Then
            For Each TempTable In cn1.rdoTables
                errorstring = ""
                If UCase$(TempTable.Type) = "TABLE" Then
                    If UCase$(TbSet.Name) = UCase$(TempTable.Name) Then
                        found = True
                        If TbSet.rdoColumns.Count > TempTable.rdoColumns.Count Then
                            Till = TempTable.rdoColumns.Count - 1
                            a = TbSet.rdoColumns.Count - TempTable.rdoColumns.Count
                            AddingColumns cn, cn1, TbSet, TempTable, a, errorstring
                          Else 'NOT TBSET.RDOCOLUMNS.COUNT...
                            Till = TbSet.rdoColumns.Count - 1
                        End If
                        For i = 0 To Till
                            If UCase$(TbSet.rdoColumns(i).Name) <> UCase$(TempTable.rdoColumns(i).Name) Then  ''Name Comparison
                                errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & ": Not Found"
                              Else 'NOT UCASE$(TBSET.RDOCOLUMNS(I).NAME)...
                                If TbSet.rdoColumns(i).Type <> TempTable.rdoColumns(i).Type Then
                                    errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Type Mistmatch"
                                  Else 'NOT TBSET.RDOCOLUMNS(I).TYPE...
                                    If TbSet.rdoColumns(i).Size <> TempTable.rdoColumns(i).Size Then
                                        errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Size Mistmatch"
                                      Else 'NOT TBSET.RDOCOLUMNS(I).SIZE...
                                        If TbSet.rdoColumns(i).Required <> TempTable.rdoColumns(i).Required Then
                                            errorstring = errorstring & Chr$(13) & Chr$(10) & Chr$(9) & TbSet.rdoColumns(i).Name & " : Null"
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                        If errorstring <> "" Then
                            Text3 = TempTable.Name & errorstring & Chr$(13) & Chr$(10) & Text3
                        End If
                        Exit For '''the correct table is compared '>---> Next
                    End If
                End If
            Next TempTable
        End If
        If found = False And UCase$(tb!table_type) = "TABLE" Then
            tmp = "cn.CreateQuery "
            AddTable cn, cn1, TbSet, TempTable
            Text3 = "<<< " & TbSet.Name & " Added To Original Database >>>" & Chr$(13) & Chr$(10) & Text3
        End If
        found = False
        Text3.Refresh
        ProgressBar1.Value = ProgressBar1.Value + 1
        tb.Close
    Next TbSet
    ProgressTitle.Caption = "Structure Check Over"
    ProgressTitle.Refresh
    Screen.MousePointer = 0

End Sub

Private Sub QuickDbCheck_Click()

    If QuickDbCheck.Value = 1 Then
        Frame1.Enabled = False
        Defaults.Enabled = False
        StoredProc.Enabled = False
        TriggerCheck.Enabled = False
        Defaults.Value = 0
        StoredProc.Value = 1
        TriggerCheck.Value = 0
      Else 'NOT QUICKDBCHECK.VALUE...
        Frame1.Enabled = True
        Defaults.Enabled = True
        StoredProc.Enabled = True
        TriggerCheck.Enabled = True
        Defaults.Value = 1
        StoredProc.Value = 1
        TriggerCheck.Value = 1
    End If

End Sub

Private Sub QuickDbChecker()

  Dim TextLine As String
  Dim tmp As String
  Dim cn As New rdoConnection
  Dim cpw As rdoQuery
  Dim tb As rdoResultset
  Dim tbf As rdoResultset
  Dim TbSet As rdoTable
  Dim coltb As rdoColumn
  Dim TempCol As rdoColumn
  Dim cn1 As New rdoConnection
  Dim TempTable As rdoTable
  Dim i As Integer
  Dim j As Integer
  Dim a As Integer
  Dim Till As Integer
  Dim found As Boolean
  Dim AnyError As Boolean
  Dim errorstring As String
  Dim ColFound As Boolean
  Dim inderrorstring As String
  Dim dbs As Integer
  Dim sql As String
  '-------------------------

    Text4 = ""
    ProgressBar1.Value = 0
    ProgressTitle.Caption = ""
    Label1.Caption = ""
    Screen.MousePointer = 11
    If Trim$(Text1) = "" Then
        MsgBox "Enter The Script Location!", vbInformation, "VBIG Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Trim$(Text2) = "" Then
        MsgBox "Enter the name of the server!", vbInformation, "VBIG Structure Check"
        Exit Sub '>---> Bottom
    End If
    If DbList.ListCount = 0 Then
        MsgBox "Enter The name(s) Of The Database(s) You Want Checked!", vbInformation, "VBIG Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Len(DbDrive) = 0 Then
        MsgBox "Enter The Drive of The Temporary Database!", vbInformation, "VBIG Structure Check"
        Exit Sub '>---> Bottom
    End If
    If Dir(Trim$(Text1)) = "" Then
        MsgBox "Script Not Found!", vbCritical
        Exit Sub '>---> Bottom
    End If
    passedTesting = True
    Close #1
    Open Trim(Text1) For Input As #1

    cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & DbList.List(0) & ";DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    tmp = "USE master" & Chr$(13) & Chr$(10)
    tmp = tmp & "select * from sysdatabases where name = 'TempStruc'"
    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)
    If tb.EOF And tb.BOF Then
        If Dir(Trim$(DbDrive) & ":\TempStruc.mdf") <> "" Then
            Kill Trim$(DbDrive) & ":\TempStruc.mdf"
        End If
        If Dir(Trim$(DbDrive) & ":\TempStruclogs.ldf") <> "" Then
            Kill Trim$(DbDrive) & ":\TempStruclogs.ldf"
        End If
      Else 'NOT TB.EOF...
        tmp = "USE master" & Chr$(13) & Chr$(10)
        tmp = "if exists (select * from sysdatabases where name = 'TempStruc') drop DATABASE TempStruc" & Chr$(13) & Chr$(10)
        Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
    End If
    DoEvents

    '------------Create Temporary Database---------------'
    tmp = "USE master" & Chr$(13) & Chr$(10)
    tmp = tmp & "CREATE DATABASE TempStruc" & Chr$(13) & Chr$(10)
    tmp = tmp & "ON ( NAME = TempStruc_dat, FILENAME = '" & Trim$(DbDrive) & ":\TempStruc.mdf',SIZE = 10,MAXSIZE = 50,FILEGROWTH = 5)"
    tmp = tmp & "LOG ON ( NAME = 'TempStruc_log', FILENAME = '" & Trim$(DbDrive) & ":\TempStruclogs.ldf',SIZE = 5MB, MAXSIZE = 25MB,  FILEGROWTH = 5MB )"

    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
    cn.Close
    Text4 = "Temporary Database Created" & Text4 & Chr$(13) & Chr$(10)

    '----------Connect to Temporary Database------------'
    cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=TempStruc;DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    '----------Running Script of Temporary Database------'
    tmp = ""
    ProgressTitle.Caption = "Creating Temporary Structure"
    ProgressTitle.Refresh
    Do While Not EOF(1)
        Label1 = Str$(Val(Label1) + 1) & " lines in Script"
        Label1.Refresh
        Line Input #1, TextLine
        If Len(TextLine) > 1 Then
            If (Left$(UCase$(TextLine), 2)) <> "GO" Then
                tmp = tmp & Chr$(13) & Chr$(10) & TextLine

              Else 'NOT (LEFT$(UCASE$(TEXTLINE),...
                On Error Resume Next
                    Set tb = cn.OpenResultset(tmp, 2, rdConcurRowVer)  ', rdAsyncEnable
                    tmp = ""
                End If
              Else 'NOT LEN(TEXTLINE)...
                tmp = tmp & Chr$(13) & Chr$(10) & TextLine
            End If
        Loop
    On Error GoTo 0
    DoEvents
    tb.Close
    '-------------'
    Close #1
    Text4 = "Database Structure Check Result" & Text4
    Text4 = "-----------------------------------------------------" & Chr$(13) & Chr$(10) & Text4
    ProgressTitle.Caption = "Checking Structure"
    ProgressTitle.Refresh

    For dbs = 0 To DbList.ListCount - 1
        ProgressTitle.Caption = "Structure Check For Database: " & DbList.List(dbs) & ""
        ProgressTitle.Refresh
        ProgressBar1.Value = 0
        ProgressBar1.Max = cn.rdoTables.Count
        DoEvents
        cn1.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & DbList.List(dbs) & ";DSN='';"
        cn1.CursorDriver = rdUseOdbc
        cn1.EstablishConnection rdDriverNoPrompt
        cn1.QueryTimeout = 3000
        cn1.rdoTables.Refresh
        found = False

        For Each TbSet In cn.rdoTables
            Set cpw = cn.CreateQuery("", "EXEC sp_tables " & TbSet.Name)
            Set tb = cpw.OpenResultset(2)
            If UCase$(TbSet.Type) = "TABLE" And UCase$(tb!table_type) = "TABLE" Then
                For Each TempTable In cn1.rdoTables
                    errorstring = ""
                    If UCase$(TempTable.Type) = "TABLE" Then
                        If UCase$(TbSet.Name) = UCase$(TempTable.Name) Then
                            found = True
                            For i = 0 To TbSet.rdoColumns.Count - 1
                                ColFound = False
                                For j = 0 To TempTable.rdoColumns.Count - 1
                                    DoEvents
                                    If UCase$(TbSet.rdoColumns(i).Name) = UCase$(TempTable.rdoColumns(j).Name) Then  ''Name Comparison
                                        ColFound = True
                                        If AddDefaults = 1 Then
                                            If TempTable.rdoColumns(j).Required = False Then
                                                ColDefaults cn1, TempTable.Name, TempTable.rdoColumns(j).Name, TempTable.rdoColumns(j).Type, errorstring
                                                DoEvents
                                            End If
                                        End If
                                        Exit For '>---> Next
                                    End If
                                Next j
                                If ColFound = False Then
                                    AddingOneColumn cn, cn1, TbSet, TempTable, errorstring, TbSet.rdoColumns(i).Name
                                    DoEvents
                                End If
                            Next i
                            If errorstring <> "" Then
                                Text4 = TempTable.Name & errorstring & Chr$(13) & Chr$(10) & Text4
                            End If
                            Exit For '''the correct table is compared '>---> Next
                        End If
                    End If
                Next TempTable
            End If
            If found = False And UCase$(tb!table_type) = "TABLE" Then
                tmp = "cn.CreateQuery "
                AddTable cn, cn1, TbSet, TempTable
                Text4 = "<<< " & TbSet.Name & " Added To Original Database >>>" & Chr$(13) & Chr$(10) & Text4
            End If
            found = False
            Text4.Refresh
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            tb.Close
        Next TbSet
        StoredProcedures cn1, errorstring
        Text4 = errorstring & Chr$(13) & Chr$(10) & Text4
        ProgressTitle.Caption = "Structure Check For Database: " & DbList.List(dbs) & ""
        ProgressTitle.Refresh
        Text4 = "  Results of Structure Check For Database: " & DbList.List(dbs) & Chr$(13) & Chr$(10) & Text4
        DoEvents
        cn1.Close
    Next dbs
    ProgressTitle.Caption = "Structure Check Over"
    ProgressTitle.Refresh

    Screen.MousePointer = 0
Bye:

Exit Sub

End Sub

Private Sub StoredProc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub StoredProcedures(cnDest As rdoConnection, errorstring As String)

  Dim DestTb As rdoResultset
  Dim DestCpw As rdoQuery
  Dim DestTb1 As rdoResultset
  Dim DestCpw1 As rdoQuery
  Dim TextLine As String
  Dim sql As String
  Dim mysql As String
  Dim tb As rdoResultset
  Dim PrName As String
  Dim MyName As String

    On Error GoTo StoredProcedures_Error

    '-------Getting New Stored procedures------'
    On Error Resume Next
        sql = ""
        Close #1
        Open Trim(Text1) For Input As #1
        Do While Not EOF(1)
            Line Input #1, TextLine
            If InStr(1, UCase$(TextLine), "CREATE PROCEDURE") Then
                cnDest.BeginTrans
                MyName = Trim$(Mid$(TextLine, 17))
                If InStr(1, UCase$(MyName), " ") <> 0 Then
                    PrName = Trim$(Mid$(MyName, 1, InStr(1, UCase$(MyName), " ") - 1))
                  Else 'NOT INSTR(1,...
                    PrName = MyName
                End If
                Set DestCpw = cnDest.CreateQuery("", "select * from sysobjects where  name='" & PrName & "' and OBJECTPROPERTY(id, N'IsProcedure') = 1")
                Set DestTb = DestCpw.OpenResultset(2)
                If Not (DestTb.EOF And DestTb.BOF) Then
                    cnDest.Execute "drop procedure [dbo].[" & PrName & "]"
                End If

                sql = TextLine & Chr$(13) & Chr$(10)
                Do While Not (InStr(1, UCase$(TextLine), "GO") > 0)
                    If Not EOF(1) Then
                        Line Input #1, TextLine
                        sql = sql & TextLine & Chr$(13) & Chr$(10)
                      Else 'NOT NOT...
                        Exit Do '>---> Loop
                    End If
                Loop
                If InStr(1, sql, "GO") > 0 Then
                    sql = Mid$(sql, 1, InStr(1, sql, "GO") - 1)
                End If
                sql = Replace(sql, Chr$(34), "'")
                Set tb = cnDest.OpenResultset(Trim$(sql), 2, rdConcurRowVer)  ', rdAsyncEnable
                sql = ""
                cnDest.CommitTrans
                errorstring = errorstring & CheckPro(PrName, cnDest.Connect, "S")
            End If
        Loop

    On Error GoTo 0

Exit Sub

StoredProcedures_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StoredProcedures of Form Form1"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Text2_LostFocus()

  Dim tb As rdoResultset
  Dim cpw As rdoQuery
  Dim cn As New rdoConnection
  Dim cnf As New rdoConnection
  Dim tbf As rdoResultset
  Dim found As Boolean
  Dim sql As String

    DbList1.Clear
    DbList.Clear
    PassToPass = ""
    On Error GoTo ErrorinIt
TryAgain:

    If Len(Trim$(Text2)) = 0 Then
        Exit Sub '>---> Bottom
    End If

    cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=Master;DSN='';"
    cn.CursorDriver = rdUseOdbc
    cn.EstablishConnection rdDriverNoPrompt

    Set cpw = cn.CreateQuery("", "Use master select * from sysdatabases ")
    Set tb = cpw.OpenResultset(2)
    Do While Not tb.EOF

        cnf.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & tb!Name & ";DSN='';"
        cnf.CursorDriver = rdUseOdbc
        cnf.EstablishConnection rdDriverNoPrompt

        ' Commented Lines are safeguards to make sure the user can select databases
        ' relating to the intended structure update, they differ from database to database
        ' but we usually choose one table that has an odd name, if it exists , we then
        ' confirm it with the existance of a field in another table

        '    sql = "select * from sysobjects where id = object_id(N'[dbo].[callcycletype]') and OBJECTPROPERTY(id, N'IsUserTable') = 1 "
        '    Set tbf = cnf.OpenResultset(sql)
        '    If Not (tbf.EOF And tbf.BOF) Then
        '        sql = "select * from sysobjects where id = object_id(N'[dbo].[properties]') and OBJECTPROPERTY(id, N'IsUserTable') = 1 "
        '        Set tbf = cnf.OpenResultset(sql)
        '        If Not (tbf.EOF And tbf.BOF) Then

        '            sql = "select name from syscolumns where id = object_id('properties')"
        '            Set tbf = cnf.OpenResultset(sql)
        '            found = False
        '            Do While Not tbf.EOF
        '                If UCase$(tbf!Name) = "CODE" Then
        '                    found = True
        '                    Exit Do
        '                End If
        '                tbf.MoveNext
        '            Loop
        '            If found Then
        DbList1.AddItem tb!Name
        '            End If
        '        End If
        '    End If
        tb.MoveNext
        cnf.Close

    Loop
    tb.Close
    cn.Close
Byebye:

Exit Sub

ErrorinIt:
    If PassToPass <> "" Then
        PassToPass = ""
        Resume TryAgain
      Else 'NOT PASSTOPASS...
        Resume Byebye
    End If

End Sub

Private Sub TriggerCheck_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

'Private Sub ConnectToDb(ByVal Db As String, cn As rdoConnection)
'PassToPass = ""
'On Error GoTo ErrorinIt
'TryAgain:
'cn.Connect = "uid=sa;pwd=" & PassToPass & ";server=" & Text2 & ";driver={SQL Server};database=" & Db & ";DSN='';"
'cn.CursorDriver = rdUseOdbc
'cn.EstablishConnection rdDriverNoPrompt
'cn.QueryTimeout = 3000
'Byebye:
'Exit Sub
'ErrorinIt:
'    If PassToPass <> "" Then
'        PassToPass = ""
'        Resume TryAgain
'    Else
'        Resume Byebye
'    End If
'End Sub
Private Sub Triggers(cnDest As rdoConnection, errorstring As String)

  Dim DestTb As rdoResultset
  Dim DestCpw As rdoQuery
  Dim TextLine As String
  Dim sql As String
  Dim tb As rdoResultset
  Dim TrName As String
  Dim MyName As String

    On Error GoTo Triggers_Error

    '-------Getting New Triggers------'
    On Error Resume Next
        sql = ""
        Close #1
        Open Trim(Text1) For Input As #1
        Do While Not EOF(1)
            Line Input #1, TextLine
            If InStr(1, UCase$(TextLine), "CREATE TRIGGER") Then
                sql = TextLine & Chr$(13) & Chr$(10)
                cnDest.BeginTrans
                MyName = Trim$(Mid$(TextLine, 15))
                If InStr(1, UCase$(MyName), " ") <> 0 Then
                    TrName = Trim$(Mid$(MyName, 1, InStr(1, UCase$(MyName), " ") - 1))
                  Else 'NOT INSTR(1,...
                    TrName = MyName
                End If
                Set DestCpw = cnDest.CreateQuery("", "select * from sysobjects where  name='" & TrName & "' and OBJECTPROPERTY(id, N'IsTrigger') = 1")
                Set DestTb = DestCpw.OpenResultset(2)
                If Not (DestTb.EOF And DestTb.BOF) Then
                    cnDest.Execute "drop trigger [dbo].[" & TrName & "]"
                End If
                Do While Not (InStr(1, UCase$(TextLine), "GO") > 0)
                    If Not EOF(1) Then
                        Line Input #1, TextLine
                        sql = sql & TextLine & Chr$(13) & Chr$(10)
                      Else 'NOT NOT...
                        Exit Do '>---> Loop
                    End If
                Loop
                If InStr(1, sql, "GO") > 0 Then
                    sql = Mid$(sql, 1, InStr(1, sql, "GO") - 1)
                End If
                sql = Replace(sql, Chr$(34), "'")
                Set tb = cnDest.OpenResultset(sql, 2, rdConcurRowVer)  ', rdAsyncEnable
                sql = ""
                cnDest.CommitTrans
                errorstring = errorstring & CheckPro(TrName, cnDest.Connect, "T")
            End If
        Loop

    On Error GoTo 0

Exit Sub

Triggers_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Triggers of Form Form1"

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jun-15 00:55) 11 + 1499 = 1510 Lines
