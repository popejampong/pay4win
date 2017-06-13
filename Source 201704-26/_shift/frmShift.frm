VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Entry"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   45
      TabIndex        =   21
      Top             =   2160
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmShift.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmShift.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmShift.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmShift.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmShift.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmShift.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmShift.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmShift.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmShift.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmShift.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Default Shift"
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5160
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "NUM:DEFAULT"
      Top             =   105
      Width           =   2340
   End
   Begin VB.TextBox Text6 
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
      Left            =   1440
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "NUM:ALLOWANCE"
      Top             =   1500
      Width           =   645
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "NUM:BTIME"
      Top             =   1125
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "NUM:REG_HR"
      Top             =   780
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Night Differential Day"
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2610
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "NUM:NDIFF"
      Top             =   105
      Width           =   2340
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
      Left            =   1440
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "TXT:REMARK"
      Top             =   1815
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
      Left            =   1440
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:DESCRIPTION"
      Top             =   435
      Width           =   5655
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
      Left            =   1440
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:SHIFTID"
      Top             =   120
      Width           =   900
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TIM:TIME1"
      Top             =   750
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56426498
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "TIM:TIME2"
      Top             =   1110
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56426498
      CurrentDate     =   38623
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Minute(s)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   1530
      Width           =   690
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grace Period"
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
      Left            =   165
      TabIndex        =   18
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hour"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4890
      TabIndex        =   17
      Top             =   1155
      Width           =   600
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hour"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4890
      TabIndex        =   16
      Top             =   810
      Width           =   600
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Breaktime"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   3315
      TabIndex        =   15
      Top             =   1155
      Width           =   1170
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   3315
      TabIndex        =   14
      Top             =   810
      Width           =   1170
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
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
      Left            =   165
      TabIndex        =   11
      Top             =   1845
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
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
      Left            =   165
      TabIndex        =   10
      Top             =   1185
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
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
      Left            =   165
      TabIndex        =   9
      Top             =   825
      Width           =   1095
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
      Left            =   165
      TabIndex        =   8
      Top             =   465
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Shift No"
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
      Left            =   165
      TabIndex        =   7
      Top             =   150
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmShift
' programmer    :   _-=[ srm ]=-_
' date          :   03 Nov 2005

Option Explicit
    Dim nAdd As Integer, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset

Private Sub Check1_Click()
    Dim cSqlStmt As String
    
    If (nAdd = 0) Or (Check1.Value = vbUnchecked) Then Exit Sub
    
    cSqlStmt = "select * from pa74380 where (`default`=1) limit 1"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If objdbRs("shiftid") <> Text1.Text Then
            cSqlStmt = "Warning!!!" & vbCrLf & _
                       "You are about to change the default shift to " & Text1.Text & vbCrLf & _
                       "The following module will be affected once you proceed" & vbCrLf & _
                       "   - default shift for new Employee" & vbCrLf & _
                       "   - default shift for new entry in Shift Schedule" & vbCrLf & _
                       "Do you wish to proceed?"
            If MsgBox(cSqlStmt, vbYesNo, "System Advisory!!!") = vbNo Then
                Check1.Value = vbUnchecked
            End If
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrShiftSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Shift Information file entry?", vbYesNoCancel, "Shift File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA74380", "SHIFTID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Shift ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA74380"), oTempADO, True
                    Script2File InsertFields(Me, "PA74380")
                    
                    Log2Audit Name, "ADD SHIFTID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD Description -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                End If
            Else
                
                OpenQueryDNS EditField(Me, "PA74380", "PA74380.SHIFTID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA74380", "PA74380.SHIFTID=" & cQuote & Text1.Text & cQuote)
                    
                Log2Audit Name, "EDIT SHIFTID -->" & Trim(Text1.Text)
                Log2Audit Name, "EDIT Description -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
            End If
            
            ' --> reset default to 0 excluding the current shift...
            If Check1.Value = vbChecked Then
                OpenQueryDNS "update pa74380 set `default`=0 where shiftid<>" & cQuote & Text1.Text & cQuote, objdbRs, True
                Script2File "update pa74380 set `default`=0 where shiftid<>" & cQuote & Text1.Text & cQuote
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "SHIFT", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "SHIFTID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    
endsave:
    Exit Sub
    
ErrShiftSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
Dim cString As String
    If nAdd = 0 Then
        Unload Me
    Else
        cString = IIf(nAdd = 2, Text1.Text, "")
        
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            If Text1.Text <> cSeries Then ResetSeries "SHIFT", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SHIFTID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
    frmLookup.showPopup 9
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "SHIFTID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then GetFields Me, oTempADO
    End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Text3.Enabled = False
    Text5.Enabled = False
    
    cSeries = GenerateSeries("SHIFT")
    While IfExists("PA74380", "PA74380.SHIFTID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("SHIFT")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Text1.Enabled = False
        Text3.Enabled = False
        Text5.Enabled = False
        
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA74380 WHERE SHIFTID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM PA74380 WHERE SHIFTID=" & cQuote & Text1.Text & cQuote
        
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
'        OpenQueryDNS "SELECT * FROM PA74380 ORDER BY SHIFTID", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
    End If
Exit Sub
ErrDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub DTPicker2_Validate(Cancel As Boolean)
    Dim nSeconds As Double
    If nAdd <> 0 Then
        If DTPicker2.Value < DTPicker1.Value Then
            nSeconds = (DateDiff("s", DTPicker1.Value, "23:59:59") + DateDiff("s", "00:00:00", DTPicker2.Value)) / 3600
        Else
            nSeconds = DateDiff("s", DTPicker1.Value, DTPicker2.Value) / 3600
        End If
        If Round(nSeconds, 2) < 7.5 Then
            MsgBox "End Time is less than the regular hours allowed!", vbCritical, "System Advisory!!!"
            DTPicker2.SetFocus
            Cancel = True
        Else
            Text3.Text = Format(IIf(Round(nSeconds, 2) > 8, 8, 7.5), "#0.00")
            Text5.Text = Format(IIf(Round(nSeconds, 2) > 8, Round(nSeconds, 2) - 8, 0.5), "#0.00")
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
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM PA74380 ORDER BY SHIFTID", oTempADO, False
    GetFields Me, oTempADO
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
