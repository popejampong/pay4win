VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Begin VB.Form frmEmpDed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Deduction/Taxes Entry"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFlex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3690
      TabIndex        =   9
      Top             =   2175
      Visible         =   0   'False
      Width           =   1035
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   585
      Left            =   4395
      TabIndex        =   8
      Top             =   3855
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1032
      LicValid        =   -1  'True
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Tag             =   "21"
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Tag             =   "20"
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&dit"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Tag             =   "18"
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3495
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1635
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1740
      TabIndex        =   4
      ToolTipText     =   "TXT:EMPID"
      Top             =   90
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Top             =   390
      Width           =   5385
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3105
      Left            =   75
      TabIndex        =   10
      Top             =   720
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   5477
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   450
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmEmpDed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmEmpDed
' programmer    :   _-=[ srm ]=-_
' description   :   Module for Employee's Deduction
' date          :   20 Oct 2005

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset
        
Sub ShowRecords()
    Dim cSqlStmt As String
    
    cSqlStmt = "SELECT A.DEDID, IFNULL(B.DEDNAME,'Undefined Deduction') as DEDNAME, if(IFNULL(B.AUTO_DED,0)=1,'Auto-Compute',A.DEF_AMT) as DEF_AMT, A.ACC_AMT,A.CUT_OFF_AMT, if(A.PERIOD1=1,'Yes','No'), if(A.PERIOD2=1,'Yes','No'), IFNULL(B.AUTO_DED,0) AS AUTO_DED, IFNULL(B.FIX_DED,0) AS FIX_DED " & _
               " FROM DI3673 A LEFT JOIN PA3330 B ON A.DEDID=B.DEDID WHERE A.EMPID=" & cQuote & Text1.Text & cQuote & _
               " ORDER BY A.DEDID"
    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount = 0 Then
'        OpenQueryDNS "SELECT DEDID, DEDNAME, IF(AUTO_DED=1,'Auto-Compute',DEF_AMT) AS DEF_AMT, 0, CUT_OFF_AMT, if(PERIOD1=1,'Yes','No'), if(PERIOD2=1,'Yes','No'), AUTO_DED, FIX_DED FROM PA3330 ORDER BY DEDID", objdbRs, False
'    End If
    
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        MSHFlexGrid1.Clear
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub chkFlex_Click()
    If Screen.ActiveForm.ActiveControl.Name = "chkFlex" Then chkFlex_KeyDown vbKeyReturn, 0
End Sub

Private Sub chkFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If (MSHFlexGrid1.ColSel = 6) Or (MSHFlexGrid1.ColSel = 7) Then
                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.Col) = IIf(chkFlex.Value = vbChecked, "Yes", "No")
            End If
            chkFlex.Visible = False
            Command3.Cancel = True
            MSHFlexGrid1.SetFocus
            
        Case vbKeyEscape
            chkFlex.Visible = False
            Command3.Cancel = True
            MSHFlexGrid1.SetFocus
    End Select
End Sub

Private Sub chkFlex_LostFocus()
    chkFlex.Visible = False
    Command3.Cancel = True
End Sub

Private Sub Command1_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        CtrlPanel Me, nAdd
    End If
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrEmpSave
    Dim nCtr As Integer, _
        cSqlStmt As String
    
    Select Case MsgBox("Save/Update employee deduction entry?", vbYesNoCancel, "Employee Deduction Entry...")
        Case vbYes
            With MSHFlexGrid1
                If IfExists("DI3673", "EMPID=" & cQuote & Text1.Text & cQuote) Then
                    OpenQueryDNS "DELETE FROM DI3673 WHERE EMPID=" & cQuote & Text1.Text & cQuote, objdbRs, True
                    Script2File "DELETE FROM DI3673 WHERE EMPID=" & cQuote & Text1.Text & cQuote
                
                    Log2Audit Name, "Delete Deduction entry of Empl. ID#" & Text1.Text
                End If
                For nCtr = 1 To .Rows - 1
                    If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                        cSqlStmt = "INSERT INTO DI3673(EMPID,DEDID,DEF_AMT,ACC_AMT,CUT_OFF_AMT,PERIOD1,PERIOD2)VALUES(" & _
                                   cQuote & Text1.Text & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                   Val(.TextMatrix(nCtr, 3)) & "," & _
                                   Val(.TextMatrix(nCtr, 4)) & "," & _
                                   Val(.TextMatrix(nCtr, 5)) & "," & _
                                   IIf(.TextMatrix(nCtr, 6) = "Yes", 1, 0) & "," & _
                                   IIf(.TextMatrix(nCtr, 7) = "Yes", 1, 0) & ")"
    '                    MsgBox cSqlStmt
                        OpenQueryDNS cSqlStmt, objdbRs, True
                    
                        Log2Audit Name, "Add Deduction ID#" & .TextMatrix(nCtr, 1) & " to Empl. ID#" & Text1.Text
                        Script2File cSqlStmt
                    End If
                Next nCtr
            End With
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321

    nAdd = 0
    CtrlPanel Me, nAdd

    ShowRecords

endsave:
    Exit Sub
    
ErrEmpSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
    Resume endsave
End Sub

Private Sub Command3_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            nAdd = 0
            CtrlPanel Me, nAdd
            
            ShowRecords
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    Tag = nAccess_Tag
    
    nAdd = 0
    
    CtrlPanel Me, nAdd
    
    myArray = Array("TXT:[Ded ID]:3:False", _
                    "TXT:[DedName]:30:True", _
                    "NUM:[Amount]:14:True", _
                    "NUM:[Accrued Amt]:14:True", _
                    "NUM:[Cut Off Amt]:14:True", _
                    "NUM:[Period 1]:8:True", _
                    "NUM:[Period 2]:8:True", _
                    "NUM:[Auto-Compute]:1:False", _
                    "NUM:[Fix]:1:False", _
                    "TXT:[CMPID]:5:Flase")
    ShowRecords
'    MSHFlexGrid1.FixedCols = 3
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
                    If (Trim(.TextMatrix(.Rows - 1, 1)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        
                        RefreshGrid MSHFlexGrid1, True
                        
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
                    
                    RefreshGrid MSHFlexGrid1, True
                    
                    '.Row = .RowSel + 1
                    .SetFocus
                End If
                
            Case vbKeyReturn
                Select Case .Col
                    Case 2
                        frmLookup.showPopup 7
                        frmLookup.Show 1
                        
                        If Trim(cResult) <> "" Then
'                    myArray = Array("TXT:1[Ded ID]:3:False", _
'                                    "TXT:2[DedName]:30:True", _
'                                    "NUM:3[Amount]:14:True", _
'                                    "NUM:4[Accrued Amt]:14:True", _
'                                    "NUM:5[Cut Off Amt]:14:True", _
'                                    "NUM:6[Period 1]:8:True", _
'                                    "NUM:7[Period 2]:8:True", _
'                                    "NUM:8[Auto-Compute]:1:False", _
'                                    "NUM:9[Fix]:1:False", _
'                                    "TXT:0[CMPID]:5:Flase")
                            
                            OpenQueryDNS "SELECT * FROM PA3330 WHERE DEDID=" & cQuote & cResult & cQuote, objdbRs, False
                            If objdbRs.RecordCount > 0 Then
                                .TextMatrix(.Row, 1) = cResult
                                .TextMatrix(.Row, 2) = DecodeStr(objdbRs("DEDNAME"))
                                .TextMatrix(.Row, 3) = IIf(objdbRs("AUTO_DED") = 1, "Auto-Compute", objdbRs("DEF_AMT"))
                                .TextMatrix(.Row, 4) = 0
                                .TextMatrix(.Row, 5) = objdbRs("CUT_OFF_AMT")
                                .TextMatrix(.Row, 6) = IIf(objdbRs("PERIOD1") = 1, "Yes", "No")
                                .TextMatrix(.Row, 7) = IIf(objdbRs("PERIOD2") = 1, "Yes", "No")
                                .TextMatrix(.Row, 8) = objdbRs("AUTO_DED")
                                .TextMatrix(.Row, 9) = objdbRs("FIX_DED")
                            End If
                        End If
                        .SetFocus
                    
                    Case 3, 5
                        If (.Col = 3) And (Val(.TextMatrix(.Row, 8)) = 1) Then Exit Sub
                        
                        If (.Col = 5) And (Val(.TextMatrix(.Row, 4)) > 0) Then
                            MsgBox "Warning!!!" & Chr$(13) & Chr$(10) & _
                                   "System does not allow editing of an active deduction...", vbCritical, "System Advisory"
                            Exit Sub
                        End If
                        
                        Command3.Cancel = False
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
                        
                    Case 6, 7
                        If Val(.TextMatrix(.Row, 9)) = 1 Then Exit Sub
                        Command3.Cancel = False
                        chkFlex.ZOrder 0
                        chkFlex.Visible = True
                        chkFlex.top = .CellTop + .top - 10 '+ (chkFlex.Height / 2)
                        chkFlex.Height = .CellHeight
                        chkFlex.left = .CellLeft + .left '+ (chkFlex.Width / 2)
                        chkFlex.Width = .CellWidth - 10
                        chkFlex.Value = IIf(.Text = "Yes", vbChecked, vbUnchecked)
                        chkFlex.SetFocus
                End Select
                
            Case vbKeyDelete
                If (.RowSel < .Rows) Then
                    
                    If (Val(.TextMatrix(.Row, 4)) > 0) Then
                        MsgBox "Warning!!!" & Chr$(13) & Chr$(10) & _
                               "System does not allow deletion of an active deduction...", vbCritical, "System Advisory"
                        Exit Sub
                    End If
                    
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
                        If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                            .RemoveItem .RowSel
                            RefreshGrid MSHFlexGrid1, True
                        End If
                    End If
                    .SetFocus
                End If
                
        End Select
        
    End With
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex") And (Screen.ActiveForm.ActiveControl.Name <> "chkFlex")
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                Select Case .ColSel
                    Case 3, 5
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
    Command3.Cancel = True
End Sub

Private Sub txtFlex_Validate(Cancel As Boolean)
    If (MSHFlexGrid1.ColSel = 3) Or (MSHFlexGrid1.ColSel = 5) Then
        Cancel = Not IsNumeric(txtFlex.Text)
    End If
End Sub
