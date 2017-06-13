VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManualFin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Contract "
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   12165
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmManualFin.frx":0000
      Left            =   330
      List            =   "frmManualFin.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   690
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3825
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtFlex 
      Height          =   375
      Left            =   195
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   187367424
      CurrentDate     =   38381
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5820
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   10266
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
      Left            =   8775
      TabIndex        =   4
      Top             =   5820
      Width           =   3345
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   120
         Picture         =   "frmManualFin.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "16"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   2355
         Picture         =   "frmManualFin.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "App&ly"
         Height          =   660
         Left            =   1395
         Picture         =   "frmManualFin.frx":3324
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmManualFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmManualFin
' description   :   Module for Automation of finish contract
' programmer    :   _-=[ srm ]=-_
' date          :   17 April 2006

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset

Sub ShowRecords()
    Dim cSqlStmt As String

    cSqlStmt = " SELECT a.empid,concat(a.lastname,', ', a.firstname,' ', a.mname ) as fullname,b.linename,c.posname,a.date_fin,'Finish','', 2,'' " & _
               " FROM di3670 a left join di5463 b on a.depid=b.lineid left join di7670 c on a.posid=c.posid " & _
               " where (a.active = 0) and (a.emp_stat<>2) and (a.date_fin <= " & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & ") " & _
               " order by a.depid,c.posid"
    
    Frame2.Enabled = False

    DoEvents
    OpenQueryDNS cSqlStmt, objdbRs, False

    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
'    Command8.Enabled = True

    Frame2.Enabled = True
End Sub


Private Sub Command10_Click()
    On Error GoTo EndManual_Fin:
    Dim nCtr As Integer, _
        cSqlStmt As String, _
        cString As String
    
'    cString = Text1.Text

    Select Case MsgBox("Apply changes in employment status?", vbYesNoCancel, App.Title)
        Case vbYes
            If nAdd = 2 Then
                With MSHFlexGrid1
                    ShowProgress 0, , 1
                    
                    For nCtr = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , , "Updating status of " & Trim(.TextMatrix(nCtr, 2))
                        If (Trim(.TextMatrix(nCtr, 8)) <> "") And (Trim(.TextMatrix(nCtr, 8)) = 0) Then
                            If (Trim(Format(.TextMatrix(nCtr, 7), "yyyy-mm-dd")) <> Trim(Format(Now, "yyyy-mm-dd"))) Then
'                            MsgBox "DITO YUNG PART NA YUNG ACTIVE = 0"
                                cSqlStmt = "update di3670 set date_fin =" & cQuote & Format(Trim(.TextMatrix(nCtr, 7)), "yyyy-mm-dd") & cQuote & "  where empid =" & cQuote & Trim(.TextMatrix(nCtr, 1)) & cQuote
                            Else
                                cSqlStmt = "update di3670 set active = 2 where empid =" & cQuote & Trim(.TextMatrix(nCtr, 1)) & cQuote
                            End If
                        Else
'                            MsgBox "DITO YUNG PART NA YUNG ACTIVE = 2"
                            cSqlStmt = "update di3670 set active = 2 where empid =" & cQuote & Trim(.TextMatrix(nCtr, 1)) & cQuote
                        End If
                        
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                        
                    Next nCtr
                    
                    ShowProgress 4
                    
                End With
            End If
        
        Case vbNo
            cString = ""
        
        Case vbCancel
            GoTo EndManual_Fin
            
    End Select
    
      
    nAdd = 0
    CtrlPanel Me, nAdd
    
    Command11_Click
'    ShowRecords
    
EndManual_Fin:
    Exit Sub
    
ErrManual_Fin:
    ErrorMsg Err.Number, Err.Description, "Apply Button", Name
    Resume EndManual_Fin
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            nAdd = 0
            
            Unload Me
'            CtrlPanel Me, nAdd
'
'            ShowRecords
        End If
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmp346626( " & _
               " [DATE_FIN] date,       [EMPID] char(6), " & _
               " [FULLNAME] char(100),  [LineName] char(100), " & _
               " [POSNAME] char(100),   [CMPName] char(100))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmp346626", oTempADO, True
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        cCmpname As String, _
        nCtr As Integer

    CreateTemp
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
             
            OpenQueryDNS "select * from di2660 where cmpid = " & cQuote & gCompanyID & cQuote, objdbRs, False
            cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
            
            cSqlStmt = "insert into tmp346626(EMPID,FULLNAME,LineName,POSNAME,DATE_FIN,CMPName)values(" & _
                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 2))) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 3))) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 4))) & cQuote & "," & _
                       cQuote & Format(.TextMatrix(nCtr, 5), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & cCmpname & cQuote & ")"
                       
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
        Next nCtr
        
        ShowProgress 3

        GenerateReport "Finish Contract Preview ", "PRV346626.RPT"

        ShowProgress 4
        
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Emp ID]:8:True", _
                    "TXT:[Name]:40:True", _
                    "TXT:[Department]:25:True", _
                    "TXT:[Position]:20:True", _
                    "TXT:[Date Finish]:12:True", _
                    "TXT:[Status]:10:True", _
                    "TXT:[Date Extended]:14:True", _
                    "NUM:[active]:1:False", _
                    "TXT:[Remark]:30:True")
    
    Tag = 1111
    nAdd = 2
    CtrlPanel Me, nAdd
    
    Command6.Enabled = True
    
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
    Dim cSqlStmt As String
    
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                Select Case .ColSel
                    Case 6     ' --> for Status ....
                        Command11.Cancel = False
                        cmbFlex.ZOrder 0
                        cmbFlex.Visible = True
                        cmbFlex.left = .CellLeft + .left - (cmbFlex.Width - .CellWidth)
                        cmbFlex.top = .CellTop + .top - 10
                        cmbFlex.ListIndex = Val(.TextMatrix(.Rows - 1, 6))
                        cmbFlex.SetFocus
                   
                    Case 7
                        
                        If .TextMatrix(.Row, .ColSel) <> "" Then
                            Command11.Cancel = False
                            dtFlex.Visible = True
                            dtFlex.left = .CellLeft + .left - (dtFlex.Width - .CellWidth)
                            dtFlex.top = .CellTop + .top - 10
                            dtFlex.Value = IIf(Trim(.Text) = "", Format(Now, "yyyy-mm-dd"), .Text)
                            dtFlex.SetFocus
                        End If
                   
                    Case 9
                        Command11.Cancel = False
                        txtFlex.ZOrder 0
                        txtFlex.Visible = True
                        txtFlex.Width = .CellWidth + 25
                        txtFlex.Height = .CellHeight
                        txtFlex.left = .CellLeft + .left
                        txtFlex.top = .CellTop + .top - 10
                        txtFlex.Text = .Text
                        txtFlex.SetFocus
                End Select
                
            Case vbKeyDelete
                If .ColSel = 4 Then
                    If MsgBox("Delete shift entry for this day?", vbYesNo, "Confirm shift deletion...") = vbYes Then
                        .TextMatrix(.Row, 3) = ""
                        .TextMatrix(.Row, 4) = ""
                        .TextMatrix(.Row, 5) = ""
                        .TextMatrix(.Row, 6) = ""
                    End If
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex") And _
                                                          (Screen.ActiveForm.ActiveControl.Name <> "cmbFlex") And _
                                                          (Screen.ActiveForm.ActiveControl.Name <> "dtFlex")
End Sub

Private Sub cmbFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                If cmbFlex.ListIndex = 0 Then
                    .TextMatrix(.RowSel, 6) = cmbFlex.Text
                    .TextMatrix(.RowSel, 7) = ""
                    .TextMatrix(.RowSel, 8) = 2
                Else
                    .TextMatrix(.RowSel, 6) = cmbFlex.Text
                    .TextMatrix(.RowSel, 7) = FormatDateTime(Now, vbShortDate)
                    .TextMatrix(.RowSel, 8) = 0
                End If
                cmbFlex_LostFocus
                .SetFocus
    
            Case vbKeyEscape
                cmbFlex_LostFocus
                .SetFocus
                
        End Select
    End With
End Sub

Private Sub cmbFlex_LostFocus()
    cmbFlex.Visible = False
    Command11.Cancel = True
End Sub


Private Sub dtFlex_DblClick()
    dtFlex_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cString As String
    
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                If .ColSel = 7 Then
                    .TextMatrix(.Row, 7) = FormatDateTime(dtFlex.Value, vbShortDate)
                    dtFlex_LostFocus
                    .SetFocus
                End If
                
            Case vbKeyEscape
                dtFlex_LostFocus
                .SetFocus
        End Select
    End With
End Sub

Private Sub dtFlex_LostFocus()
    dtFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                If .ColSel = 9 Then .TextMatrix(.Row, .ColSel) = txtFlex.Text
                
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

'Private Sub txtFlex_Validate(Cancel As Boolean)
'    If MSHFlexGrid1.ColSel = 9 Then
'        Cancel = Not IsNumeric(txtFlex.Text)
'    End If
'End Sub


