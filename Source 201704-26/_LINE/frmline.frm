VERSION 5.00
Begin VB.Form frmline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Entry"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmline.frx":0000
      Left            =   1725
      List            =   "frmline.frx":0022
      TabIndex        =   25
      Tag             =   "1"
      Text            =   "Combo1"
      ToolTipText     =   "NUM:ERPPOSCODE"
      Top             =   1335
      Width           =   2070
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
      Left            =   1725
      TabIndex        =   21
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   1035
      Width           =   1290
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   315
      Left            =   3060
      TabIndex        =   20
      Top             =   1035
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   3060
      TabIndex        =   18
      Top             =   735
      Width           =   495
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
      Left            =   1725
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "TXT:COSTCENTERID"
      Top             =   735
      Width           =   1290
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   0
      TabIndex        =   5
      Top             =   1665
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmline.frx":0043
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmline.frx":19C5
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmline.frx":3347
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmline.frx":4CC9
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmline.frx":664B
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmline.frx":7FCD
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmline.frx":994F
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmline.frx":B2D1
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmline.frx":CC53
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmline.frx":E5D5
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Production Line"
      Height          =   285
      Left            =   2490
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "NUM:PRODUCTION"
      Top             =   150
      Width           =   2295
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
      Left            =   1725
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:LINEID"
      Top             =   135
      Width           =   720
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
      Left            =   1725
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:LINENAME"
      Top             =   435
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Left            =   90
      TabIndex        =   24
      Top             =   1365
      Width           =   1530
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center Code"
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
      Left            =   90
      TabIndex        =   23
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3615
      TabIndex        =   22
      Top             =   1080
      Width           =   5310
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3615
      TabIndex        =   19
      Top             =   780
      Width           =   5310
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Center Code"
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
      Left            =   90
      TabIndex        =   17
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label3 
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
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   465
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1905
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmLine
' description   :   Module for Maintenance of Department/Line
' programmer    :   _-=[ srm ]=-_
' date          :   19 jan 2005
' note          :   copied from DICAS

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset

Private Sub Command10_Click()
    On Error GoTo ErrDeptSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Line file entry?", vbYesNoCancel, "Line File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("di5463", "LINEID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Deparment ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "di5463"), oTempADO, True
                    Script2File InsertFields(Me, "di5463")
                    
                    Log2Audit Name, "ADD LINE ID -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "di5463", "di5463.LINEID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "di5463", "di5463.LINEID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT LINE ID -->" & Trim(Text1.Text)
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select
            
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "LINE", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "LINEID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    ShowRecords
    
endsave:
    Exit Sub
ErrDeptSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "LINE", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "LINEID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords
            
        End If
        
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
    ShowRecords
End Sub

Private Sub Command2_Click()
    frmLookup.showPopup 20
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text4.Text = cResult
            Label14.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
        End If
    End If
    
    Text4.SetFocus
End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 21
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA97722 WHERE WORKCENTERID=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text3.Text = cResult
            Label2.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
        End If
    End If
    
    Text4.SetFocus
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
        frmLookup.showPopup 2
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "LINEID='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
            If Not oTempADO.EOF Then GetFields Me, oTempADO
            ShowRecords
        End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Label14.Caption = ""
    Command2.Enabled = True
    
    Label2.Caption = ""
    Command3.Enabled = True
    
    cSeries = GenerateSeries("LINE")
    While IfExists("di5463", "di5463.LINEID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("LINE")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    ' --> modified 20050321
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Command2.Enabled = True
        Command3.Enabled = True

        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDeptDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM di5463 WHERE LINEID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM di5463 WHERE LINEID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
'        OpenQueryDNS "SELECT * FROM di5463 ORDER BY LINEID", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
Exit Sub
ErrDeptDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM di5463 ORDER BY LINEID", oTempADO, False
    
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

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(Text3.Text) = "" Then
        Command3_Click
    Else
        OpenQueryDNS "SELECT * FROM pa97722 WHERE WORKCENTERID=" & cQuote & Text3.Text & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Label2.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
            Text3.SetFocus
        Else
            Label2.Caption = ""
            MsgBox "Work Center Ref. No. Not Found!", vbCritical, App.Title
            Text3.SetFocus
        End If
    End If
End If


End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(Text4.Text) = "" Then
        Command2_Click
    Else
        OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text4.Text & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Label14.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
            Text3.SetFocus
        Else
            Label14.Caption = ""
            MsgBox "Cost Center Ref. No. Not Found!", vbCritical, App.Title
            Text4.SetFocus
        End If
    End If
End If


End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    
     ' ---> Cost Center 201207-25
    OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text4.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        Label14.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
    Else
        Label14.Caption = ""
    End If
    
     ' ---> Cost Center 201207-25
    OpenQueryDNS "SELECT * FROM pa97722 WHERE WORKCENTERID=" & cQuote & Text3.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        Label2.Caption = objdbRs("DESCRIPTION") & " / " & objdbRs("COMPCODE")
    Else
        Label2.Caption = ""
    End If
    
    Command2.Enabled = nAdd <> 0
    Command3.Enabled = nAdd <> 0
End Sub


