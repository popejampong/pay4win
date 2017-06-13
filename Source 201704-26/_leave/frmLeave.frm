VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaxppanel.ocx"
Begin VB.Form frmLeave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incentive Leave Entry"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5070
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2265
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   825
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3105
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4875
      _ExtentX        =   8599
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
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   585
      Left            =   870
      TabIndex        =   2
      Top             =   3225
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1032
      LicValid        =   -1  'True
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Tag             =   "21"
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Tag             =   "20"
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&dit"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Tag             =   "18"
         Top             =   45
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-In Payroll System
' module        :   frmLeave
' description   :   Module for Incentive Leave
' programmer    :   _-=[ srm ]=-_
' date          :   28 Oct 2005

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        oTempADO As New ADODB.Recordset


Sub ShowRecords()
    Dim cSqlStmt As String
    
    cSqlStmt = "SELECT RANGE1, RANGE2, SL, VL FROM PA53283 ORDER BY RANGE1"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , True, True
    Else
        MSHFlexGrid1.Clear
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub Command1_Click()
    nAdd = 2
    CtrlPanel Me, nAdd
    
'    MSHFlexPeriod.Enabled = False
    MSHFlexGrid1.SetFocus
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrEmpSave
    Dim nCtr As Integer, _
        cSqlStmt As String
    
    Select Case MsgBox("Save/Update Incentive Leave entry?", vbYesNoCancel, "Incentive Leave Entry...")
        Case vbYes
            OpenQueryDNS "delete from PA53283", objdbRs, True
            Script2File "delete form pa53283"
            
            With MSHFlexGrid1
                For nCtr = 1 To .Rows - 1
                    cSqlStmt = "INSERT INTO PA53283(RANGE1, RANGE2, SL, VL)VALUES(" & _
                               Val(.TextMatrix(nCtr, 1)) & "," & _
                               Val(.TextMatrix(nCtr, 2)) & "," & _
                               Val(.TextMatrix(nCtr, 3)) & "," & _
                               Val(.TextMatrix(nCtr, 4)) & ")"
'                    MsgBox cSqlStmt
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                
                Next nCtr
                Log2Audit Name, "Modify Incentive Leave Entry"
            End With
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
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
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
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
    
    myArray = Array("NUM:[Range 1]:12.2:True", _
                    "NUM:[Range 2]:12.2:True", _
                    "NUM:[SL]:10.2:True", _
                    "NUM:[VL]:10.2:True")
    Tag = nAccess_Tag
    nAdd = 0
    
    CtrlPanel Me, nAdd
    
    ShowRecords
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (nAdd > 0) Then
        MsgBox "Please click CANCEL to abort this entry...", vbOKOnly, App.Title
        Cancel = 1
    Else
        Log2Audit Me.Name, "CLOSE"
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
                    
                    .SetFocus
                End If
                
            Case vbKeyReturn
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
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex")
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
'            myArray = Array("NUM:1[Range 1]:12:True", _
'                            "NUM:2[Range 2]:12:True", _
'                            "NUM:3[SL]:10:True", _
'                            "NUM:4[VL]:10:True")
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
                
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
    Cancel = Not IsNumeric(txtFlex.Text)
End Sub
