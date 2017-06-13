VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Code Entry"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   45
      TabIndex        =   8
      Top             =   4095
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmTax.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmTax.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmTax.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmTax.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmTax.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmTax.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmTax.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmTax.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmTax.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmTax.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3810
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Left            =   1695
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:TAXNAME"
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   1695
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:TAXCODE"
      Top             =   405
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1695
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:TAXID"
      Top             =   90
      Width           =   525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3045
      Left            =   1695
      TabIndex        =   7
      Top             =   1050
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5371
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Top             =   735
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax ID"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code"
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
      TabIndex        =   3
      Top             =   420
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   4110
      Left            =   0
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmTax
' description   :   Tax Code module
' programmer    :   _-=[ srm ]=-_
' date          :   24 feb 2006

Option Explicit
    Dim nAdd As Integer, myArray As Variant
    Dim cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        nLastRow As Integer

Sub FillGrid()
    Dim nCtr As Integer
    
    With MSHFlexGrid1
        .Redraw = False
        .Clear

        SetGridColumn myArray, MSHFlexGrid1
        
        DoEvents
        
        For nCtr = 0 To UBound(aTaxAmt)
            If (aTaxAmt(nCtr) > 0) Or (aTaxPct(nCtr) > 0) Then
                .Rows = nCtr + 2
                .TopRow = .Rows - 1
                
                .RowHeight(nCtr + 1) = 285
                .TextMatrix(nCtr + 1, 1) = 0
                .TextMatrix(nCtr + 1, 2) = aTaxAmt(nCtr)
                .TextMatrix(nCtr + 1, 3) = aTaxPct(nCtr)
            End If
        Next nCtr
        
        RefreshGrid MSHFlexGrid1, True
        
        .Redraw = True
    End With
End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    cSqlStmt = "SELECT DED_AMT2, " & _
               "       DED_AMT, " & _
               "       DED_PCT " & _
               " FROM PA8293 " & _
               " WHERE TAXID=" & cQuote & Text1.Text & cQuote & _
               " ORDER BY SEQ_NO"
    Frame2.Enabled = False
    
    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        MSHFlexGrid1.Clear
        SetGridColumn myArray, MSHFlexGrid1
    End If
    
    Frame2.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        ShowRecords
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrTaxSave
    Dim cString As String, _
        nCtr As Integer, _
        cSqlStmt As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & "Tax Code entry?", vbYesNoCancel, "Tax Code Entry...")
    
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA8290", "TAXID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Tax Id already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA8290"), oTempADO, True
                    Script2File InsertFields(Me, "PA8290")
                    
                    Log2Audit Name, "ADD TAX ID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                    
                    ShowProgress 0
                    
                    With MSHFlexGrid1
                    
                        For nCtr = 1 To (.Rows - 1)
                        
                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                            
                            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                                cSqlStmt = "INSERT INTO PA8293(TAXID,DED_AMT2,DED_AMT,DED_PCT,SEQ_NO)VALUES(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           Val(.TextMatrix(nCtr, 1)) & "," & _
                                           Val(.TextMatrix(nCtr, 2)) & "," & _
                                           Val(.TextMatrix(nCtr, 3)) & "," & _
                                           nCtr & ")"
                                OpenQueryDNS cSqlStmt, objdbRs, True
                            End If
                            
                        Next nCtr
                        
                    End With
                    
                    ShowProgress 4
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "PA8290", "PA8290.TAXID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA8290", "PA8290.TAXID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT TAX ID -->" & Trim(Text1.Text)
                Log2Audit Name, "EDIT ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
            
                OpenQueryDNS "DELETE FROM PA8293 WHERE TAXID=" & cQuote & Text1.Text & cQuote, objdbRs, True
                Script2File "DELETE FROM PA8293 WHERE TAXID=" & cQuote & Text1.Text & cQuote
                
                ShowProgress 0
                
                With MSHFlexGrid1
                
                    For nCtr = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            cSqlStmt = "INSERT INTO PA8293(TAXID,DED_AMT2,DED_AMT,DED_PCT,SEQ_NO)VALUES(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       Val(.TextMatrix(nCtr, 1)) & "," & _
                                       Val(.TextMatrix(nCtr, 2)) & "," & _
                                       Val(.TextMatrix(nCtr, 3)) & "," & _
                                       nCtr & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                    Next nCtr
                    
                End With
                
                ShowProgress 4
                
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "TAX", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    Text2.Enabled = False
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "TAXID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    ShowRecords

endsave:
    Exit Sub
    
ErrTaxSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            
            If Text1.Text <> cSeries Then ResetSeries "TAX", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "TAXID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    
    frmLookup.showPopup 8
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "TAXID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    FillGrid
    
    cSeries = GenerateSeries("TAX")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA8290", "PA8290.TAXID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("TAX")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    Dim nCtr As Integer
    
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        If ((MSHFlexGrid1.Rows - 1) = 1) And MSHFlexGrid1.TextMatrix(1, 1) = "" Then FillGrid
        
        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA8290 WHERE TAXID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM PA8290 WHERE TAXID=" & cQuote & Text1.Text & cQuote
        
        OpenQueryDNS "DELETE FROM PA8293 WHERE TAXID=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "DELETE FROM PA8293 WHERE TAXID=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE " & Trim(Text2.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text3.Text)))
        
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
    
    Exit Sub
    
ErrDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("NUM:[Bracket]:12.2:True", _
                    "NUM:[Amount]:12.2:True", _
                    "NUM:[%]:7.2:True")
    
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
        
    OpenQueryDNS "SELECT * FROM PA8290 ORDER BY TAXID", oTempADO, False
    
    GetFields Me, oTempADO
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

Private Sub MSHFlexGrid1_LeaveCell()
    nLastRow = MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = Screen.ActiveForm.ActiveControl.Name <> "txtFlex"
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Command11.Cancel = True
End Sub

Private Sub txtFlex_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtFlex.Text)
End Sub

