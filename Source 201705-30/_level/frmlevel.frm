VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmlevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Enry"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   840
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5106
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   4575
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   120
         Picture         =   "frmlevel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "15"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   2580
         Picture         =   "frmlevel.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1575
         Picture         =   "frmlevel.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "18"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   3570
         Picture         =   "frmlevel.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmlevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmlevel
' description   :   Module for Level of rate and cola of each employee (customized for almatech fo the meantime)
' programmer    :   _-=[ srm ]=-_
' date          :   25 jan 2012

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        nLastRow As Integer
    
Sub ShowRecords()
    Dim cSqlStmt As String
    
    CtrlPanel Me, nAdd
    
        
    cSqlStmt = " SELECT LVLCode,Rate,Cola,CMPID FROM pa5380 order by LVLCode "

    Frame2.Enabled = False
    
    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False
    
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        MSHFlexGrid1.Clear
        SetGridColumn myArray, MSHFlexGrid1
    End If
    Command8.Enabled = True
    
    Frame2.Enabled = True
End Sub


Private Sub Command10_Click()
    On Error GoTo ErrSave
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer, _
        nTotAmt As Double
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " level Table Entry?", vbYesNoCancel, App.Title)
    
        Case vbYes
            If nAdd = 1 Then
                With MSHFlexGrid1

                    ShowProgress 0
                    ShowProgress 1, , .Rows - 1

                    For nCtr = 1 To .Rows - 1

                        ShowProgress 2, nCtr

                        If Not (Trim(.TextMatrix(nCtr, 1)) = "") Then

                            ShowProgress 2, nCtr, , , "Saving " & .TextMatrix(nCtr, 1)

                            cSqlStmt = "INSERT INTO pa5380(LVLCode,Rate,Cola,CMPID)VALUES(" & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
                            
'                                MsgBox cSqlStmt

                            OpenQueryDNS cSqlStmt, objdbRs, True

                        End If
                    Next nCtr

                    ShowProgress 4

                End With
            Else
            
                OpenQueryDNS "DELETE FROM pa5380", objdbRs, True
                Script2File "DELETE FROM pa5380 "

                
                With MSHFlexGrid1

                    ShowProgress 0
                    ShowProgress 1, , .Rows - 1

                    For nCtr = 1 To .Rows - 1

                        ShowProgress 2, nCtr

                        If Not (Trim(.TextMatrix(nCtr, 1)) = "") Then

                            ShowProgress 2, nCtr, , , "Saving " & .TextMatrix(nCtr, 1)

                            cSqlStmt = "INSERT INTO pa5380(LVLCode,Rate,Cola,CMPID)VALUES(" & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
                            
'                                MsgBox cSqlStmt

                            OpenQueryDNS cSqlStmt, objdbRs, True

                        End If
                    Next nCtr
                        
                    ShowProgress 4

                End With
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
    Lock2User Name, "level", "level", False     ' --> 20050328

    nAdd = 0
    CtrlPanel Me, nAdd

    MSHFlexGrid1.LeftCol = 2
    ShowRecords
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save level Table Entry", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
'        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, "level", "level", False     ' --> 20050321
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            ShowRecords
        End If
        
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmplevel(" & _
               " [LVLCode] char(2), " & _
               " [Rate] decimal(18,4), " & _
               " [Cola] decimal(18,4), " & _
               " [CMPID] char(4), " & _
               " [CMPNAME] char(99))"

    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
        
ErrCreate:
'    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmplevel", oTempADO, True
End Sub


Private Sub Command7_Click()
    nAdd = 1
    CtrlPanel Me, nAdd
    
    MSHFlexGrid1.Clear
    SetGridColumn myArray, MSHFlexGrid1
    
    MSHFlexGrid1.SetFocus

End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        gCmpName As String

    CreateTemp

    With MSHFlexGrid1

        ShowProgress 0, , .Rows - 1

        For nCtr = 1 To .Rows - 1

            ShowProgress 2, nCtr

            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
            
                ShowProgress 2, nCtr, , , "Copying " & Trim(.TextMatrix(nCtr, 1)) & "..."
                
                 OpenQueryDNS " select cmpname from  di2660 where cmpid= " & cQuote & gCompanyID & cQuote, objdbRs, False
                gCmpName = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
                                
                cSqlStmt = " INSERT INTO tmplevel(LVLCode,Rate,Cola,CMPID,CMPNAME)VALUES(" & _
                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & .TextMatrix(nCtr, 2) & "," & _
                           .TextMatrix(nCtr, 3) & "," & .TextMatrix(nCtr, 4) & "," & _
                           cQuote & gCmpName & cQuote & ")"
                           
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, oRecordSet, True
            End If
        Next

        ShowProgress 3
        
        GenerateReport "Level Table Preview", "PRV5380.RPT", , True

        ShowProgress 4
    End With

    Set oRecordSet = Nothing

End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, "level", "level") Then
        Lock2User Me.Name, "level", "level", True
        
        nAdd = 2
        CtrlPanel Me, nAdd
        
    '    SetGridColumn myArray, MSHFlexGrid1
        
        MSHFlexGrid1.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
'    myArray = Array("NUM:[RANGE 1]:13.2:True", _
'                    "NUM:[RANGE 2]:13.2:True", _
'                    "NUM:[Sal. Credit]:13.2:True", _
'                    "NUM:[ER SS]:12.2:True", _
'                    "NUM:[EE SS]:12.2:True", _
'                    "NUM:[SS Total]:15.2:True", _
'                    "NUM:[ER EC]:10.2:True", _
'                    "NUM:[ER Contri]:12.2:True", _
'                    "NUM:[EE Contri]:12.2:True", _
'                    "NUM:[Total Contri]:15.2:True", _
'                    "TXT:[CMPID]:5:Flase")
    
    myArray = Array("NUM:[Level Code]:13.2:True", _
                    "NUM:[Rate]:13.2:True", _
                    "NUM:[Cola]:13.2:True", _
                    "TXT:[CMPID]:5:Flase")
    
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM pa5380 ORDER BY LVLCode", oTempADO, False
    GetFields Me, oTempADO
    ShowRecords
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (nAdd > 0) Then
        MsgBox "Please click CANCEL to abort this entry...", vbOKOnly, App.Title
        Cancel = 1
    Else
        Log2Audit Name, "CLOSE"
    End If
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub MSHFlexGrid1_DblClick()
    MSHFlexGrid1_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid1_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyDown
                If (nLastRow = .Row) And (nLastRow = .Rows - 1) Then
                    If (Trim(.TextMatrix(.Rows - 1, 1)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        
                        RefreshGrid MSHFlexGrid1, True
                        
                        .LeftCol = 1
                        .Col = 1
                        .ColSel = 1
                    End If
                Else
                    nLastRow = .Row
                End If

            Case vbKeyUp
                If .Rows - 1 > 1 Then
                    If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then
                        nLastRow = nLastRow - 1
                        .Rows = .Rows - 1
                    End If
                End If
                
            Case vbKeyInsert    ' --> 20050908
                If .TextMatrix(.RowSel, 1) <> "" Then
                    .AddItem "", .RowSel
                    .RowHeight(.RowSel) = 285
                    
                    RefreshGrid MSHFlexGrid1, True
                    
                    '.Row = .RowSel + 1
                    .SetFocus
                End If
                
            Case vbKeyReturn
                Select Case .ColSel
                    Case 1, 2, 3
                        Command11.Cancel = False
                        txtFlex.ZOrder 0
                        txtFlex.Visible = True
                        txtFlex.Width = .CellWidth + 25
                        txtFlex.Height = .CellHeight
                        txtFlex.left = .CellLeft + .left
                        txtFlex.top = .CellTop + .top - 10
                        txtFlex.Text = .Text
                        txtFlex.SelStart = 0
                        txtFlex.SelLength = Len(.Text)
                        txtFlex.SetFocus
                End Select
            
            Case vbKeyDelete
                If (.RowSel < .Rows) Then
                    If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                        If MsgBox("Delete Record ?", vbYesNo, App.Title) = vbYes Then
                            If .Rows - 1 = 1 Then
                                .AddItem "", .Rows
                                .RowHeight(.RowSel + 1) = 285
                            End If
                            .RemoveItem .RowSel
                            RefreshGrid MSHFlexGrid1, True
                        End If
                    Else
                        .RemoveItem .RowSel
                        RefreshGrid MSHFlexGrid1, True
                    End If
                    
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LeaveCell()
    nLastRow = MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = Screen.ActiveForm.ActiveControl.Name <> "txtFlex"
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                
    '            myArray = Array("NUM:1[RANGE 1]:15:True", _
    '                            "NUM:2[RANGE 2]:15:True", _
    '                            "NUM:3[Sal. Credit]:15:True", _
    '                            "NUM:4[ER SS]:12:True", _
    '                            "NUM:5[EE SS]:12:True", _
    '                            "NUM:6[SS Total]:15:True", _
    '                            "NUM:7[ER EC]:10:True", _
    '                            "NUM:8[ER Contri]:12:True", _
    '                            "NUM:9[EE Contri]:12:True", _
    '                            "NUM:0[Total Contri]:15:True", _
    '                            "TXT:1[CMPID]:5:Flase")
                Select Case MSHFlexGrid1.ColSel
                    Case 1
                        .TextMatrix(.Row, .ColSel) = txtFlex.Text
                    Case 2, 3
                        .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
                        
                        
                End Select
                
                txtFlex_LostFocus
                .SetFocus
                 
            Case vbKeyEscape
                txtFlex_LostFocus
                .SetFocus
        End Select
    End With
End Sub
 
Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub txtFlex_Validate(Cancel As Boolean)
    If MSHFlexGrid1.ColSel = 10 Then
        Cancel = Not IsNumeric(txtFlex.Text)
    End If
End Sub
