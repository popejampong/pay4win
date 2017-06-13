VERSION 5.00
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaxpframe.ocx"
Begin VB.Form frmDeduction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Entry"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
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
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "TXT:DEDERPID"
      Top             =   2235
      Width           =   1200
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
      Left            =   1695
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "TXT:SHORT_DESC"
      Top             =   1005
      Width           =   1560
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:DEDNAME2"
      Top             =   705
      Width           =   4290
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmDeduction.frx":0000
      Left            =   1680
      List            =   "frmDeduction.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "NUM:DEDTYPE"
      Top             =   1305
      Width           =   2340
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmDeduction.frx":0028
      Left            =   5385
      List            =   "frmDeduction.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "NUM:DEDTAG"
      Top             =   1320
      Width           =   3675
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Auto-compute"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2775
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "NUM:AUTO_DED"
      Top             =   120
      Width           =   1500
   End
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   750
      Left            =   6315
      TabIndex        =   28
      Top             =   60
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1323
      Alignment       =   2
      Caption         =   " Inclusive Period "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.CheckBox Check2 
         Caption         =   "Period 1 ( 1 - 15 )"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   165
         TabIndex        =   21
         Tag             =   "1"
         ToolTipText     =   "NUM:PERIOD1"
         Top             =   180
         Value           =   1  'Checked
         Width           =   2370
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Period 2 ( 16 - end of month )"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   165
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "NUM:PERIOD2"
         Top             =   420
         Value           =   1  'Checked
         Width           =   2370
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fixed Entry"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4440
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "NUM:FIX_DED"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
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
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "NUM:CUT_OFF_AMT"
      Top             =   1935
      Width           =   1560
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
      Left            =   1695
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "NUM:DEF_AMT"
      Top             =   1635
      Width           =   1560
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
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:DEDID"
      Top             =   105
      Width           =   825
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
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:DEDNAME"
      Top             =   405
      Width           =   4290
   End
   Begin VB.Frame Frame2 
      Height          =   930
      Left            =   75
      TabIndex        =   23
      Top             =   2520
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7080
         Picture         =   "frmDeduction.frx":0058
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "20"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6240
         Picture         =   "frmDeduction.frx":19DA
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "19"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5400
         Picture         =   "frmDeduction.frx":335C
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "18"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2640
         Picture         =   "frmDeduction.frx":4CDE
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "12"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1800
         Picture         =   "frmDeduction.frx":6660
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "14"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   960
         Picture         =   "frmDeduction.frx":7FE2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "13"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8040
         Picture         =   "frmDeduction.frx":9964
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "21"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4560
         Picture         =   "frmDeduction.frx":B2E6
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "17"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3600
         Picture         =   "frmDeduction.frx":CC68
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "15"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   120
         Picture         =   "frmDeduction.frx":E5EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "11"
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ERP Item Code"
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
      TabIndex        =   33
      Top             =   2265
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Short Desc"
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
      TabIndex        =   32
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alternate Desc"
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
      TabIndex        =   31
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      TabIndex        =   30
      Top             =   1335
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Additional Information"
      Height          =   270
      Left            =   5400
      TabIndex        =   29
      Top             =   1110
      Width           =   2625
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default Amount"
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
      TabIndex        =   27
      Top             =   1665
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      TabIndex        =   26
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cut Off Amount"
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
      TabIndex        =   25
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction ID"
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
      TabIndex        =   24
      Top             =   135
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   2730
      Left            =   0
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmDeduction
' description   :   Module for Maintenance of Deduction
' programmer    :   _-=[ srm ]=-_
' date          :   17 Oct 2005

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset


Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrMatColorSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Deduction file entry?", vbYesNoCancel, "Deduction File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA3330", "DEDID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Deduction Id already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA3330"), oTempADO, True
                    Script2File InsertFields(Me, "PA3330")
                    
                    Log2Audit Name, "ADD DEDID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD DEDNAME -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                End If
            Else
                OpenQueryDNS EditField(Me, "PA3330", "PA3330.DEDID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA3330", "PA3330.DEDID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT DEDID -->" & Trim(Text1.Text)
                Log2Audit Name, "EDIT DEDNAME -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "DEDUCTION", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    Text2.Enabled = False
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "DEDID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO

endsave:
    Exit Sub
ErrMatColorSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "DEDUCTION", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Text2.Enabled = False
           
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "DEDID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 7
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "DEDID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
        End If
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    cSeries = GenerateSeries("DEDUCTION")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA3330", "PA3330.DEDID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("DEDUCTION")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        Text1.Enabled = False
        
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrPeriodDelete
    
    If Check1.Value = vbChecked Then
        MsgBox "You are not allowed to delete this fixed entry!", vbCritical, "System Advisory!!!"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deduction Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM PA3330 WHERE DEDID=" & cQuote & Text1.Text & cQuote, oTempADO, True

        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM PA3330 WHERE DEDID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
    End If
    
    Exit Sub
    
ErrPeriodDelete:
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
        
    OpenQueryDNS "SELECT * FROM PA3330 ORDER BY DEDID", oTempADO, False
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
