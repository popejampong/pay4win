VERSION 5.00
Begin VB.Form frmWorkCenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Center Entry"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
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
      Left            =   1755
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "TXT:COMPCODE"
      Top             =   930
      Width           =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   2640
      TabIndex        =   19
      Top             =   930
      Width           =   495
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
      Left            =   1755
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "TXT:COSTCENTERID"
      Top             =   630
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   315
      Left            =   3270
      TabIndex        =   15
      Top             =   630
      Width           =   495
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
      Left            =   1755
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   30
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   0
      TabIndex        =   12
      Top             =   1305
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmWorkCenter.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmWorkCenter.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmWorkCenter.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmWorkCenter.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmWorkCenter.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmWorkCenter.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmWorkCenter.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmWorkCenter.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmWorkCenter.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmWorkCenter.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
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
      Left            =   1755
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:DESCRIPTION"
      Top             =   330
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company Code"
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
      Left            =   105
      TabIndex        =   22
      Top             =   945
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3210
      TabIndex        =   21
      Top             =   975
      Width           =   5310
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   75
      TabIndex        =   18
      Top             =   645
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3825
      TabIndex        =   17
      Top             =   675
      Width           =   5310
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
      TabIndex        =   14
      Top             =   345
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   0
      TabIndex        =   13
      Top             =   75
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "frmWorkCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmCostCenter
' description   :   Module for Maintenance of ERP-Cost Center
' programmer    :   _-=[ srm ]=-_
' date          :   20 July 2012

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset

Private Sub Command10_Click()
    On Error GoTo ErrDeptSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Work Center file entry?", vbYesNoCancel, "Work Center File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("pa97722", "WORKCENTERID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Refernce no already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "pa97722"), oTempADO, True
                    Script2File InsertFields(Me, "pa97722")
                    
                    Log2Audit Name, "ADD Work Center ID -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "pa97722", "pa97722.WORKCENTERID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "pa97722", "pa97722.WORKCENTERID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT Work Center ID -->" & Trim(Text1.Text)
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select
            
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "ERPWC", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "WORKCENTERID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
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
            
            If Text1.Text <> cSeries Then ResetSeries "ERPWC", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "WORKCENTERID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
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
    frmLookup.showPopup 22
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA2660 WHERE COMPCODE=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text4.Text = cResult
            Label5.Caption = objdbRs("COMPNAME")
        End If
    End If
    
    Text4.SetFocus

End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 20
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text3.Text = cResult
            Label4.Caption = objdbRs("DESCRIPTION")
        End If
    End If
    
    Text4.SetFocus

End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
        frmLookup.showPopup 21
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "WORKCENTERID='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
            If Not oTempADO.EOF Then GetFields Me, oTempADO
        End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
'    cSeries = GenerateSeries("ERPWC")
'    While IfExists("pa97722", "pa97722.WORKCENTERID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
'        cSeries = GenerateSeries("ERPWC")
'        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
'    Wend
'    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
'
'    Text1.SetFocus

    Label4.Caption = ""
    Command2.Enabled = True
    
    Label5.Caption = ""
    Command3.Enabled = True
    
    cSeries = GenerateSeries("ERPWC")
    While IfExists("pa97722", "pa97722.WORKCENTERID='" & cPrefix & PadStr(cSeries, "0", nCodeLen) & "'")
        cSeries = GenerateSeries("ERPWC")
        
        Text1.Text = cPrefix & PadStr(cSeries, "0", nCodeLen)
    Wend
    Text1.Text = cPrefix & PadStr(cSeries, "0", nCodeLen)
        
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
        OpenQueryDNS "DELETE FROM pa97722 WHERE WORKCENTERID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM pa97722 WHERE WORKCENTERID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
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
    
    OpenQueryDNS "SELECT * FROM pa97722 ORDER BY WORKCENTERID", oTempADO, False
    
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
            OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text3.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label4.Caption = objdbRs("DESCRIPTION")
                Text4.SetFocus
            Else
                Label4.Caption = ""
                MsgBox "Cost Center ID Not Found!", vbCritical, App.Title
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
                OpenQueryDNS "SELECT * FROM PA2660 WHERE COMPCODE=" & cQuote & Text4.Text & cQuote, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    Label5.Caption = objdbRs("COMPName")
                Else
                    Label5.Caption = ""
                    MsgBox "Company Code Not Found!", vbCritical, App.Title
                    Text4.SetFocus
                End If
            End If
    End If

End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text3.Text & cQuote, objdbRs, False
    Label4.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
    OpenQueryDNS "SELECT * FROM pa2660 WHERE COMPCODE=" & cQuote & Text4.Text & cQuote, objdbRs, False
    Label5.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("COMPName"), "")

    Command3.Enabled = nAdd <> 0
    Command2.Enabled = nAdd <> 0
        
End Sub

