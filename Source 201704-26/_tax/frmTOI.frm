VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTOI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   75
      TabIndex        =   7
      Top             =   5670
      Width           =   4920
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   3930
         Picture         =   "frmTOI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1935
         Picture         =   "frmTOI.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "18"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   2940
         Picture         =   "frmTOI.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   120
         Picture         =   "frmTOI.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "15"
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   975
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   1875
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "NUM:S_AMT"
      Top             =   4125
      Width           =   1140
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
      Left            =   1875
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "NUM:H_AMT"
      Top             =   4425
      Width           =   1140
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
      Left            =   1875
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "NUM:M_AMT"
      Top             =   4725
      Width           =   1140
   End
   Begin VB.TextBox Text4 
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
      Left            =   1875
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "NUM:EX_AMT"
      Top             =   5055
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   60
      TabIndex        =   0
      Top             =   3765
      Width           =   4920
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3495
      Left            =   75
      TabIndex        =   6
      Top             =   225
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   6165
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
      Caption         =   "Income Tax Rates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   75
      TabIndex        =   17
      Top             =   0
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Single"
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
      Left            =   60
      TabIndex        =   15
      Top             =   4140
      Width           =   1740
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Head of the Family"
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
      Left            =   60
      TabIndex        =   14
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Married Individual"
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
      Left            =   60
      TabIndex        =   13
      Top             =   4740
      Width           =   1740
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Add'l Exemption (per dependent not exceeding 4)"
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
      Height          =   570
      Left            =   60
      TabIndex        =   12
      Top             =   5070
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   0
      Top             =   4110
      Width           =   1860
   End
   Begin VB.Label Label2 
      Caption         =   "Allowance for Personal Exemption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   75
      TabIndex        =   16
      Top             =   3885
      Width           =   3795
   End
End
Attribute VB_Name = "frmTOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmTOI
' description   :   Module for Maintenance of TOI Deduction
' programmer    :   _-=[ srm ]=-_
' date          :   05 jan 2007

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        nLastRow As Integer
    
Sub ShowRecords()
    Dim cSqlStmt As String
    
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "select * from pa4870", objdbRs, False
    GetFields Me, objdbRs
    
    cSqlStmt = "SELECT range1, range2, AMOUNT, format(PERCENT,0) as PERCENT FROM PA4870 order by range1 "

    Frame2.Enabled = False
    
    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False
    
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
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
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " TOI Table Entry?", vbYesNoCancel, App.Title)
    
        Case vbYes
            If nAdd = 1 Then
                With MSHFlexGrid1

                    ShowProgress 0

                    For nCtr = 1 To .Rows - 1
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
    
                        If Not ((Trim(.TextMatrix(nCtr, 1)) = "") And (Trim(.TextMatrix(nCtr, 2)) = "")) Then

                            cSqlStmt = "INSERT INTO PA4870(RANGE1,RANGE2,AMOUNT,PERCENT,S_AMT,H_AMT,M_AMT,EX_AMT,CMPID)VALUES(" & _
                                       Val(.TextMatrix(nCtr, 1)) & "," & _
                                       Val(.TextMatrix(nCtr, 2)) & "," & _
                                       Val(.TextMatrix(nCtr, 3)) & "," & _
                                       Val(.TextMatrix(nCtr, 4)) & "," & _
                                       Val(Format(Text2.Text, "###.#00")) & "," & _
                                       Val(Format(Text1.Text, "###.#00")) & "," & _
                                       Val(Format(Text3.Text, "###.#00")) & "," & _
                                       Val(Format(Text4.Text, "###.#00")) & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt

                        End If
                    Next nCtr

                    ShowProgress 4

                End With
            Else
                OpenQueryDNS "DELETE FROM PA4870", objdbRs, True
                Script2File "DELETE FROM PA4870 "
                With MSHFlexGrid1

                        ShowProgress 0

                        For nCtr = 1 To .Rows - 1

                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100

                            If Not ((Trim(.TextMatrix(nCtr, 1)) = "") And (Trim(.TextMatrix(nCtr, 2)) = "")) Then

                                cSqlStmt = "INSERT INTO PA4870(RANGE1,RANGE2,AMOUNT,PERCENT,S_AMT,H_AMT,M_AMT,EX_AMT,CMPID)VALUES(" & _
                                           Val(.TextMatrix(nCtr, 1)) & "," & _
                                           Val(.TextMatrix(nCtr, 2)) & "," & _
                                           Val(.TextMatrix(nCtr, 3)) & "," & _
                                           Val(.TextMatrix(nCtr, 4)) & "," & _
                                           Val(Format(Text2.Text, "###.#00")) & "," & _
                                           Val(Format(Text1.Text, "###.#00")) & "," & _
                                           Val(Format(Text3.Text, "###.#00")) & "," & _
                                           Val(Format(Text4.Text, "###.#00")) & "," & _
                                           cQuote & gCompanyID & cQuote & ")"
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt

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
    
    Lock2User Name, "TOI", "TOI", False     ' --> 20050328

    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd

    MSHFlexGrid1.LeftCol = 2
    ShowRecords
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save TOI Table Entry", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
'        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, "TOI", "TOI", False     ' --> 20050321
            
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
    
    cSqlStmt = " CREATE TABLE tmpTOI(" & _
               " [RANGE1] decimal(18,4),    [RANGE2] decimal(18,4), " & _
               " [AMOUNT] decimal(18,4),    [PERCENT] decimal(18,4), " & _
               " [CMPID] char(4))"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpTOI", oTempADO, True
End Sub


Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        aUserInfo As Variant

    aUserInfo = Array("", "", "", "", "", "")
    
    CreateTemp
    
    With MSHFlexGrid1
        
        ShowProgress 0

        For nCtr = 1 To .Rows - 1
            
            ShowProgress 2, nCtr

            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
            
                ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                
                cSqlStmt = " INSERT INTO tmpTOI(RANGE1,RANGE2,[AMOUNT],[PERCENT],CMPID)VALUES(" & _
                           Val(.TextMatrix(nCtr, 1)) & "," & Val(.TextMatrix(nCtr, 2)) & "," & _
                           Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & _
                           cQuote & gCompanyID & cQuote & ")"
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, oRecordSet, True
            End If
        Next
        
        ShowProgress 3
        GenerateReport "Tax On Individual Table Preview", "PRVPA4870.RPT", , True

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
    If Not isDataLock(Me.Name, "TOI", "TOI") Then
        Lock2User Me.Name, "TOI", "TOI", True
        
        nAdd = 2
        ClearAll Me, True, False
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
    
    myArray = Array("NUM:[RANGE 1]:13.2:True", _
                    "NUM:[RANGE 2]:13.2:True", _
                    "NUM:[Amount]:13.2:True", _
                    "NUM:[Pecentage %]:13.2:True", _
                    "TXT:[CMPID]:5:Flase")
    
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM PA4870 ORDER BY RANGE1", oTempADO, False
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
                    Case 1, 2, 3, 4
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
                
                Select Case MSHFlexGrid1.ColSel
                    Case 1, 2, 3, 4
                        .TextMatrix(.Row, .ColSel) = Val(txtFlex.Text)
                        
                        If .Col = 4 Then
                            SendKeys "{DOWN}"
                        End If
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


