VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPhilhealth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhilHealth Table Entry"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11550
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3885
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5895
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   10398
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
      Height          =   915
      Left            =   6555
      TabIndex        =   5
      Top             =   5940
      Width           =   4920
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   120
         Picture         =   "frmPhilhealth.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "15"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   2940
         Picture         =   "frmPhilhealth.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1935
         Picture         =   "frmPhilhealth.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "18"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   3930
         Picture         =   "frmPhilhealth.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPhilhealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmPhilhealth
' description   :   Module for Maintenance of Philhealth Deduction
' programmer    :   _-=[ srm ]=-_
' date          :   16 Oct 2005

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset

Sub ShowRecords()
    Dim cSqlStmt As String
    
    CtrlPanel Me, nAdd
    
    cSqlStmt = " SELECT msal_brac, range1, range2, sal_base, " & _
               " mtot_cont, ps, es FROM pa7454 order by msal_brac "

    Frame2.Enabled = False
    
    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False
    
    If objdbRs.RecordCount > 0 Then
'        Command7.Enabled = False
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , True, True
    Else
        MSHFlexGrid1.Clear
'        Command7.Enabled = True
'        Command8.Enabled = True
        SetGridColumn myArray, MSHFlexGrid1
    End If
    
    Command8.Enabled = True
    
    Frame2.Enabled = True
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmpPhil_h([MSAL_BRAC] integer,    " & _
               " [RANGE1] decimal(18,4),    [RANGE2] decimal(18,4), " & _
               " [SAL_BASE] decimal(18,4),  [MTOT_CONT] decimal(18,4), " & _
               " [PS] decimal(18,4),        [ES] decimal(18,4), " & _
               " [CMPID] char(4))"
               

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpPhil_h", oTempADO, True
End Sub


Private Sub Command10_Click()
    On Error GoTo ErrSave
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer, _
        nTotAmt As Double
    
    Select Case MsgBox("Save Philhealth Table Entry?", vbYesNoCancel, App.Title)
    
        Case vbYes
            OpenQueryDNS "DELETE FROM PA7454", objdbRs, True
            Script2File "DELETE FROM PA7454 "

            
            With MSHFlexGrid1

                ShowProgress 0

                For nCtr = 1 To (.Rows - 1)

                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , , "Updating " & .TextMatrix(nCtr, 4)

                    If Not ((Trim(.TextMatrix(nCtr, 2)) = "") And (Trim(.TextMatrix(nCtr, 4)) = "")) Then

                        cSqlStmt = "INSERT INTO PA7454(MSAL_BRAC,RANGE1,RANGE2,SAL_BASE,MTOT_CONT,PS,ES,CMPID)VALUES(" & _
                                   cQuote & nCtr & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 4) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 5) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & _
                                   cQuote & gCompanyID & cQuote & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt

                    End If
                    
                Next nCtr
                    
                ShowProgress 4

            End With
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
'    Lock2User Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050328

    nAdd = 0
    CtrlPanel Me, nAdd

    MSHFlexGrid1.LeftCol = 2
    ShowRecords
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save Philhealth Table Entry", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
        
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            ShowRecords
        End If
        
    End If

End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset

    CreateTemp
    
    With MSHFlexGrid1
        
        ShowProgress 0, , .Rows - 1

        For nCtr = 1 To .Rows - 1
            
            ShowProgress 2, nCtr

            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                ShowProgress 2, nCtr, , , "Copying " & Trim(.TextMatrix(nCtr, 4)) & "..."
                cSqlStmt = " INSERT INTO tmpPhil_h(MSAL_BRAC,RANGE1,RANGE2,SAL_BASE,MTOT_CONT,PS,ES,CMPID)VALUES(" & _
                           .TextMatrix(nCtr, 1) & "," & .TextMatrix(nCtr, 2) & "," & _
                           .TextMatrix(nCtr, 3) & "," & .TextMatrix(nCtr, 4) & "," & _
                           .TextMatrix(nCtr, 5) & "," & .TextMatrix(nCtr, 6) & "," & _
                           .TextMatrix(nCtr, 7) & "," & cQuote & gCompanyID & cQuote & ")"
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, oRecordSet, True
            End If
        Next
        
        GenerateReport "PhilHealth Table Preview", "PRVpa7454.RPT", , True

        ShowProgress 4
        
    End With
    
    Set oRecordSet = Nothing
End Sub

Private Sub Command7_Click()
    nAdd = 1
    CtrlPanel Me, nAdd
    
    MSHFlexGrid1.Clear
    SetGridColumn myArray, MSHFlexGrid1
    
    MSHFlexGrid1.SetFocus

End Sub

Private Sub Command8_Click()
    nAdd = 2
    CtrlPanel Me, nAdd
    
'    SetGridColumn myArray, MSHFlexGrid1
    
    MSHFlexGrid1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("NUM:[Sal Bracket]:12:False", _
                    "NUM:[RANGE1]:12.2:True", _
                    "NUM:[RANGE2]:12.2:True", _
                    "NUM:[Salary Base]:12.2:True", _
                    "NUM:[Tot Contri.]:12.2:True", _
                    "NUM:[EE Share]:12.2:True", _
                    "NUM:[ER Share]:12.2:True", _
                    "TXT:[CMPID]:5:False")
    
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM PA7454 ORDER BY MSAL_BRAC", oTempADO, False
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
                If .Row = .Rows - 1 Then
                    If (Trim(.TextMatrix(.Rows - 1, 2)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        
                        RefreshGrid MSHFlexGrid1, True, True
                        
                        .LeftCol = 2
                        .Col = 2
                        .ColSel = 2
                    End If
                End If

            Case vbKeyUp
                If .Rows - 1 > 1 Then
                    If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then
                        .Rows = .Rows - 1
                    End If
                End If
                
            Case vbKeyInsert    ' --> 20050908
                If .TextMatrix(.RowSel, 1) <> "" Then
                    .AddItem "", .RowSel
                    .RowHeight(.RowSel) = 285
                    
                    RefreshGrid MSHFlexGrid1, True, True
                    
                    '.Row = .RowSel + 1
                    .SetFocus
                End If
                
            Case vbKeyReturn
                If nAdd <> 0 Then
                    Select Case .ColSel
                        Case 2, 3, 4, 5
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
                End If
            
            Case vbKeyDelete
            
                If (.RowSel < .Rows) Then
                    If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                        If MsgBox("Delete Record ?", vbYesNo, App.Title) = vbYes Then
                            If .Rows - 1 = 1 Then
                                .AddItem "", .Rows
                                .RowHeight(.RowSel + 1) = 285
                            End If
                            .RemoveItem .RowSel
                            RefreshGrid MSHFlexGrid1, True, True
                        End If
                    Else
                        If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                            .RemoveItem .RowSel
                            RefreshGrid MSHFlexGrid1, True, True
                        End If
                    End If
                    .SetFocus
                End If
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = Screen.ActiveForm.ActiveControl.Name <> "txtFlex"
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cSqlStmt As String, _
        nRow As Integer
    
    Select Case KeyCode
    
        Case vbKeyReturn
            With MSHFlexGrid1
                Select Case .ColSel
                
        '    myArray = Array("NUM:1[Sal Bracket]:12:False", _
        '                    "NUM:2[RANGE1]:12.2:True", _
        '                    "NUM:3[RANGE2]:12.2:True", _
        '                    "NUM:4[Salary Base]:12.2:True", _
        '                    "NUM:5[Tot Contri.]:12.2:True", _
        '                    "NUM:6[EE Share]:12.2:True", _
        '                    "NUM:7[ER Share]:12.2:True", _
        '                    "TXT:8[CMPID]:5:False")
                
                    Case 2, 3
                        .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
                        
                        ' --> optional only...
                        If .ColSel = 3 Then
                            .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 2))
                            .TextMatrix(.Row, 5) = (Val(.TextMatrix(.Row, 2)) * 0.0125) * 2
                            .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 2)) * 0.0125
                            .TextMatrix(.Row, 7) = Val(.TextMatrix(.Row, 2)) * 0.0125
                        End If
                        
                        .Col = .ColSel + 1
                        
                    Case 4
                        .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
                        .TextMatrix(.Row, 5) = (Val(txtFlex.Text) * 0.0125) * 2
                        .TextMatrix(.Row, 6) = Val(txtFlex.Text) * 0.0125
                        .TextMatrix(.Row, 7) = .TextMatrix(.Row, 6)
                        
                        MSHFlexGrid1_KeyDown vbKeyDown, 0
                        
'                    Case 5
'                        .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
'                        If .Col = 5 Then
'                            SendKeys "{DOWN}"
'                        End If
                
                End Select
            End With
            
            txtFlex_LostFocus
            MSHFlexGrid1.SetFocus
             
        Case vbKeyEscape
            txtFlex_LostFocus
            MSHFlexGrid1.SetFocus
    End Select
End Sub

Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub txtFlex_Validate(Cancel As Boolean)
    If MSHFlexGrid1.ColSel = 7 Then
        Cancel = Not IsNumeric(txtFlex.Text)
    End If
End Sub
