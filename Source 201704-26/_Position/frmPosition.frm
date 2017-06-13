VERSION 5.00
Begin VB.Form frmPosition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Position Entry"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmPosition.frx":0000
      Left            =   1485
      List            =   "frmPosition.frx":000A
      TabIndex        =   18
      Tag             =   "1"
      Top             =   960
      Width           =   1255
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   1325
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmPosition.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmPosition.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmPosition.frx":3324
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmPosition.frx":4CA6
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmPosition.frx":6628
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmPosition.frx":7FAA
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmPosition.frx":992C
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
         Picture         =   "frmPosition.frx":B2AE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmPosition.frx":CC30
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmPosition.frx":E5B2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Management and Staff"
      Height          =   285
      Left            =   2235
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "NUM:STAFF"
      Top             =   50
      Width           =   2295
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
      Left            =   1485
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "NUM:ALLOWANCE"
      Top             =   665
      Width           =   1230
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
      Left            =   1485
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:POSID"
      Top             =   35
      Width           =   675
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
      Left            =   1485
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:POSNAME"
      Top             =   350
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Height          =   375
      Left            =   370
      TabIndex        =   19
      Top             =   1000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Food Allowance"
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
      TabIndex        =   6
      Top             =   680
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Position ID"
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
      Left            =   240
      TabIndex        =   5
      Top             =   80
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Left            =   105
      TabIndex        =   4
      Top             =   380
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1340
      Left            =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmPosition
' description   :   Module for Maintenance of Employee's Position
' programmer    :   _-=[ srm ]=-_
' date          :   17 jan 2005

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset
    Dim cSqlStmt As String


Private Sub Command10_Click()
    On Error GoTo ErrDeptSave
    Dim cString As String
    
    cString = Text1.Text

    '---> (201704-07) (TLC)
    If Combo1.ListIndex = -1 Then
        MsgBox "Please specify designation!", vbCritical, App.Title
        Combo1.SetFocus
        GoTo endsave
    End If


    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Position entry?", vbYesNoCancel, "Position Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("DI7670", "POSID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Employee Position ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "DI7670"), oTempADO, True
                    Script2File InsertFields(Me, "DI7670")
    
    '------------------(201704-07)(TLC)>
                    cSqlStmt = "insert into DI7673(posid,designation)values(" & _
                           cQuote & Text1.Text & cQuote & "," & _
                           cQuote & Combo1.ListIndex & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
    '-----------------(201704-07)(TLC)<

                    
                    Log2Audit Name, "ADD POSITION ID -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "DI7670", "DI7670.POSID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "DI7670", "DI7670.POSID=" & cQuote & Text1.Text & cQuote)
                
 '------------------(201704-07)(TLC) >
                If IfExists("DI7673", "posid=" & cQuote & Text1.Text & cQuote) Then
                     cSqlStmt = "UPDATE DI7673 SET designation=" & cQuote & Combo1.ListIndex & cQuote & _
                                " where posid=" & cQuote & Text1.Text & cQuote
                     OpenQueryDNS cSqlStmt, objdbRs, True
               Else
                     cSqlStmt = "insert into DI7673(posid,designation)values(" & _
                           cQuote & Text1.Text & cQuote & "," & _
                           cQuote & Combo1.ListIndex & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
               
                End If
'------------------(201704-07)(TLC) <

                
                
                Log2Audit Name, "EDIT POSITION ID -->" & Trim(Text1.Text)
            End If
            
        Case vbNo
            cString = ""
        Case vbCancel
            GoTo endsave
    End Select
            
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "POSITION", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "POSID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    
    
'--------------->(201704-07)(TLC) >
        If IfExists("DI7673", "POSID=" & cQuote & Text1.Text & cQuote) Then
                    OpenQueryDNS "SELECT * FROM DI7673 WHERE POSID=" & cQuote & Text1.Text & cQuote, objdbRs, False
            Combo1.ListIndex = objdbRs("designation")
            If Combo1.ListIndex = 0 Then
                Combo1.ListIndex = 0
            ElseIf Combo1.ListIndex = 1 Then
                Combo1.ListIndex = 1
            Else
                Combo1.ListIndex = -1
            End If
    Else
           Combo1.ListIndex = -1
    End If

'--------------->(201704-07)(TLC) <
    
    
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
            
            If Text1.Text <> cSeries Then ResetSeries "POSITION", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "POSID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        End If
        
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
    
    '--------------->(201704-07)(TLC) >
        If IfExists("DI7673", "POSID=" & cQuote & Text1.Text & cQuote) Then
                    OpenQueryDNS "SELECT * FROM DI7673 WHERE POSID=" & cQuote & Text1.Text & cQuote, objdbRs, False
            Combo1.ListIndex = objdbRs("designation")
            If Combo1.ListIndex = 0 Then
                Combo1.ListIndex = 0
            ElseIf Combo1.ListIndex = 1 Then
                Combo1.ListIndex = 1
            Else
                Combo1.ListIndex = -1
            End If
    Else
           Combo1.ListIndex = -1
    End If

'--------------->(201704-07)(TLC) <

    
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
        frmLookup.showPopup 4
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "POSID='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
            If Not oTempADO.EOF Then GetFields Me, oTempADO
        End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    cSeries = GenerateSeries("POSITION")
    While IfExists("DI7670", "DI7670.POSID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("POSITION")
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
        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDeptDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM DI7670 WHERE POSID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        OpenQueryDNS "DELETE FROM DI7673 WHERE POSID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM DI7670 WHERE POSID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
'        OpenQueryDNS "SELECT * FROM di5463 ORDER BY LINEID", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
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
    
    OpenQueryDNS "SELECT * FROM DI7670 ORDER BY POSID", oTempADO, False
    
    GetFields Me, oTempADO
    
'--------------->(201704-07)(TLC) >
        If IfExists("DI7673", "POSID=" & cQuote & Text1.Text & cQuote) Then
                    OpenQueryDNS "SELECT * FROM DI7673 WHERE POSID=" & cQuote & Text1.Text & cQuote, objdbRs, False
            Combo1.ListIndex = objdbRs("designation")
            If Combo1.ListIndex = 0 Then
                Combo1.ListIndex = 0
            ElseIf Combo1.ListIndex = 1 Then
                Combo1.ListIndex = 1
            Else
                Combo1.ListIndex = -1
            End If
    Else
           Combo1.ListIndex = -1
    End If

'--------------->(201704-07)(TLC) <

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

