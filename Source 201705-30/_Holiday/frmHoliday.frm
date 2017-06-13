VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHoliday 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Holiday"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmHoliday.frx":0000
      Left            =   1380
      List            =   "frmHoliday.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "NUM:TAG"
      Top             =   1020
      Width           =   3645
   End
   Begin VB.CheckBox Check2 
      Caption         =   "&Non-Working Holiday"
      Height          =   315
      Left            =   6615
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "NUM:WITHPAY"
      Top             =   720
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   45
      TabIndex        =   19
      Top             =   1380
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmHoliday.frx":003A
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmHoliday.frx":19BC
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmHoliday.frx":333E
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmHoliday.frx":4CC0
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmHoliday.frx":6642
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmHoliday.frx":7FC4
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmHoliday.frx":9946
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmHoliday.frx":B2C8
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
         Picture         =   "frmHoliday.frx":CC4A
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
         Picture         =   "frmHoliday.frx":E5CC
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Fixed Holiday"
      Height          =   315
      Left            =   6615
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "NUM:FIX_DAY"
      Top             =   390
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
      Left            =   1395
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:HOLIDAYID"
      Top             =   90
      Width           =   645
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
      Left            =   1395
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:DESCRIPTION"
      Top             =   705
      Width           =   4290
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE"
      Top             =   405
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16384000
      CurrentDate     =   38623
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Condition"
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
      TabIndex        =   20
      Top             =   1065
      Width           =   1110
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
      TabIndex        =   18
      Top             =   735
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      TabIndex        =   17
      Top             =   435
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Holiday ID"
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
      TabIndex        =   16
      Top             =   150
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1395
      Left            =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "frmHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmholiday
' description   :   Module for Maintenance of Holidays
' programmer    :   _-=[ srm ]=-_
' date          :   16 Oct 2005

Option Explicit
    Dim nAdd As Integer, myArray As Variant
    Dim cSeries As String, _
        oTempADO As New ADODB.Recordset

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        DTPicker1.Format = dtpCustom
        DTPicker1.CustomFormat = "MMMM d"
    Else
        DTPicker1.Format = dtpLongDate
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrMatColorSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Holiday file entry?", vbYesNoCancel, "Holiday File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA4329", "HOLIDAYID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Holiday Id already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA4329"), oTempADO, True
                    Script2File InsertFields(Me, "PA4329")
                    
                    Log2Audit Name, "ADD HOLIDAYID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                End If
            Else
                OpenQueryDNS EditField(Me, "PA4329", "PA4329.HOLIDAYID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA4329", "PA4329.HOLIDAYID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT HOLIDAYID -->" & Trim(Text1.Text)
                Log2Audit Name, "EDIT ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
            End If
        Case vbNo
            cString = ""
        Case vbCancel
            GoTo endsave
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "HOLIDAY", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    Text2.Enabled = False
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "HOLIDAYID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
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
            
            If Text1.Text <> cSeries Then ResetSeries "HOLIDAY", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Text2.Enabled = False
           
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "HOLIDAYID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 6
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "HOLIDAYID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
        End If
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    cSeries = GenerateSeries("HOLIDAY")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA4329", "PA4329.HOLIDAYID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("HOLIDAY")
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
        DTPicker1.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrPeriodDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Holiday Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM PA4329 WHERE HOLIDAYID=" & cQuote & Text1.Text & cQuote, oTempADO, True

        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM PA4329 WHERE HOLIDAYID=" & cQuote & Text1.Text & cQuote
        
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
        
    OpenQueryDNS "SELECT * FROM PA4329 ORDER BY HOLIDAYID", oTempADO, False
    
    GetFields Me, oTempADO
    
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

