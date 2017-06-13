VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Period Entry"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Tag "
      ForeColor       =   &H00400000&
      Height          =   660
      Left            =   1695
      TabIndex        =   28
      Top             =   15
      Width           =   4725
      Begin VB.OptionButton Option1 
         Caption         =   "Annual Withheld Tax Period"
         Height          =   420
         Index           =   2
         Left            =   2925
         TabIndex        =   30
         Tag             =   "1"
         ToolTipText     =   "NUM:WTAX"
         Top             =   195
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "13th Month Period"
         Height          =   420
         Index           =   1
         Left            =   1335
         TabIndex        =   29
         Tag             =   "1"
         ToolTipText     =   "NUM:13MONTH"
         Top             =   195
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "N/A"
         Height          =   420
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   105
      TabIndex        =   17
      Top             =   2910
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmPeriod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmPeriod.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmPeriod.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmPeriod.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmPeriod.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmPeriod.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmPeriod.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmPeriod.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmPeriod.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmPeriod.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPeriod.frx":FF14
      Left            =   1695
      List            =   "frmPeriod.frx":FF1E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "NUM:PCLOSE"
      Top             =   1935
      Width           =   2040
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
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "TXT:DURATION"
      Top             =   1620
      Width           =   4290
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
      ToolTipText     =   "TXT:PERIODID"
      Top             =   690
      Width           =   615
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
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "NUM:HOLIDAYS"
      Top             =   2580
      Width           =   885
   End
   Begin VB.TextBox Text5 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7995
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "NUM:STATUS"
      Top             =   465
      Visible         =   0   'False
      Width           =   915
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
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "NUM:WORKINDAYS"
      Top             =   2265
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1695
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_START"
      Top             =   1005
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   49741824
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   1695
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_END"
      Top             =   1305
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   49741824
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   300
      Left            =   3855
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_CLOSE"
      Top             =   1950
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   529
      _Version        =   393216
      Format          =   49741824
      CurrentDate     =   38623
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Period No"
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
      TabIndex        =   15
      Top             =   750
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      TabIndex        =   13
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
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
      TabIndex        =   12
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inclusive Date"
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
      TabIndex        =   11
      Top             =   1650
      Width           =   1470
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Period Close"
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
      TabIndex        =   10
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays"
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
      TabIndex        =   9
      Top             =   2625
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   2925
      Left            =   0
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmPeriod
' description   :   Module for Maintenance of Payroll Period
' programmer    :   _-=[ srm ]=-_
' date          :   07 Oct 2005

Option Explicit
    Dim nAdd As Integer, myArray As Variant
    Dim cSeries As String, _
        cParam As String
    Dim oTempADO As New ADODB.Recordset, _
        Dvalue1, Dvalue2 As Date
        
Sub ProcessPeriod()
    Dim nCtr As Integer, _
        cSqlStmt, _
        cDateStart, _
        cDateEnd As String
        
    ShowProgress 0, , 100
    For nCtr = 2 To 25
        ShowProgress 2, (nCtr / 25) * 100, , , "Please wait processing " & MonthName(Format(Int(nCtr / 2), "00")) & " period..."
        cDateStart = Year(Now) & "-" & Format(Int(nCtr / 2), "00") & "-" & Format(IIf(nCtr / 2 = Int(nCtr / 2), "1", "16"), "00")
        If nCtr / 2 <> Int(nCtr / 2) Then
            cDateEnd = "LAST_DAY(" & cQuote & cDateStart & cQuote & ")"
        Else
            cDateEnd = cQuote & Year(Now) & "-" & Format(Int(nCtr / 2), "00") & "-15" & cQuote
        End If
        
        
        cSeries = GenerateSeries("PERIOD")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
        While IfExists("PA7730", "PA7730.PERIODID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
            cSeries = GenerateSeries("PERIOD")
            Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
        Wend

        cSqlStmt = "INSERT INTO PA7730(PERIODID,DATE_START,DATE_END,WORKINDAYS,STATUS)VALUES(" & _
                   cQuote & Text1.Text & cQuote & "," & cQuote & cDateStart & cQuote & "," & cDateEnd & "," & _
                   "DATEDIFF(" & cDateEnd & "," & cQuote & cDateStart & cQuote & ")," & _
                   IIf(nCtr / 2 <> Int(nCtr / 2), 2, 1) & ")"
'        MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, objdbRs, True
    Next nCtr
    ShowProgress 4
    
    OpenQueryDNS "SELECT * FROM PA7730 ORDER BY PERIODID", oTempADO, False
    If oTempADO.RecordCount > 0 Then
        ShowProgress 0, , 100, , "Please wait performing verification..."
        While Not oTempADO.EOF
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            GetFields Me, oTempADO
            getdate
            OpenQueryDNS EditField(Me, "PA7730", "PERIODID=" & cQuote & oTempADO("PERIODID") & cQuote), objdbRs, True
            oTempADO.MoveNext
        Wend
        ShowProgress 4
    End If

    OpenQueryDNS "SELECT * FROM PA7730 ORDER BY PERIODID", oTempADO, False
    GetFields Me, oTempADO
    CtrlPanel Me, nAdd
End Sub

Function WorkDay(oDate1 As DTPicker, oDate2 As DTPicker) As Double
    Dim nCtr As Integer, _
        nWrkDays As Integer
    For nCtr = 0 To (oDate2.Value - oDate1.Value)
        nWrkDays = nWrkDays + IIf(Weekday(oDate1.Value + nCtr) <> 1, 1, 0)
    Next nCtr
    WorkDay = nWrkDays
End Function

Sub getdate()
    Dim nCtr As Integer, _
        nDOW As DTPicker, _
        nWrkDays As Long, _
        nSundayCount As Integer
    
    DoEvents
    
    If Option1(1).Value Then
        Text2.Text = "13th Month Pay " & Year(DTPicker1.Value)
        Text3.Text = 26
    Else
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "Start Date is greater than End Date!", vbCritical, "System Advisory!"
            Text2.Text = "Start Date is greater than End Date!"
        Else
            Text5.Text = IIf(Day(DTPicker2.Value) > 15, 2, 1)
            If Month(DTPicker1.Value) = Month(DTPicker2.Value) Then
                Text2.Text = MonthName(Month(DTPicker1.Value)) & " " & Day(DTPicker1.Value) & "-" & Day(DTPicker2.Value) & ", " & Year(DTPicker2.Value)
            Else
                Text2.Text = MonthName(Month(DTPicker1.Value)) & " " & Day(DTPicker1.Value) & "-" & MonthName(Month(DTPicker2.Value)) & " " & Day(DTPicker2.Value) & ", " & Year(DTPicker2.Value)
            End If
    
            Text3.Text = WorkDay(DTPicker1, DTPicker2)
        End If
    End If
End Sub


Private Sub Check1_Click()
    If Option1(1).Value Then
        DTPicker1.Value = "01/01/" & Year(Now)
        DTPicker2.Value = "12/31/" & Year(Now)
        getdate
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 1 Then
        DTPicker3.Visible = True
        DTPicker3.Value = Now
    Else
        DTPicker3.Visible = False
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        CtrlPanel Me, nAdd, oTempADO("pclose") <> 1
        Option1_Click 0
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrMatColorSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Period file entry?", vbYesNoCancel, "Period File Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA7730", "PERIODID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Period Id already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA7730"), oTempADO, True
                    Script2File InsertFields(Me, "PA7730")
                    
                    Log2Audit Name, "ADD PERIODID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                    
                End If
            Else
                    OpenQueryDNS EditField(Me, "PA7730", "PA7730.PERIODID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                    Script2File EditField(Me, "PA7730", "PA7730.PERIODID=" & cQuote & Text1.Text & cQuote)
                    
                    Log2Audit Name, "EDIT PERIODID -->" & Trim(Text1.Text)
                    Log2Audit Name, "EDIT ADD INCLUSIVE DATE -->" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
                    
            End If
        Case vbNo
            cString = ""
        Case vbCancel
            GoTo endsave
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "PERIOD", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    Text2.Enabled = False
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "PERIODID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    Option1_Click 0

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
            
            If Text1.Text <> cSeries Then ResetSeries "PERIOD", cSeries
            
            nAdd = 0
            ClearAll Me, False, True
            
            Text2.Enabled = False
           
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "PERIODID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            CtrlPanel Me, nAdd
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 5
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "PERIODID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            CtrlPanel Me, nAdd, oTempADO("pclose") <> 1
            Option1_Click 0
        End If
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    getdate
    
    Text2.Enabled = False
    
    Combo1_Click
    Option1_Click 0
    
    cSeries = GenerateSeries("PERIOD")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA7730", "PA7730.PERIODID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("PERIOD")
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
        
        Option1_Click 0
        
        Text1.Enabled = False
        Text2.Enabled = False
        
        DTPicker1.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrPeriodDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Employee Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM PA7730 WHERE PERIODID=" & cQuote & Text1.Text & cQuote, oTempADO, True

        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM PA7730 WHERE PERIODID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        
        OpenQueryDNS "SELECT * FROM PA7730 ORDER BY PERIODID", oTempADO, False
        GetFields Me, oTempADO
        
        CtrlPanel Me, nAdd, oTempADO("pclose") <> 1
        Option1_Click 0
    End If
    
    Exit Sub
    
ErrPeriodDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub DTPicker1_Change()
    getdate
End Sub

Private Sub DTPicker2_Change()
    getdate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
        
    OpenQueryDNS "SELECT * FROM PA7730 ORDER BY PERIODID", oTempADO, False
    GetFields Me, oTempADO
    If oTempADO.RecordCount = 0 Then
        If MsgBox("Would you like the system to auto-generate period for the year " & Year(Now) & "?", vbYesNo, App.Title) = vbYes Then ProcessPeriod
    End If

    CtrlPanel Me, nAdd, oTempADO("pclose") <> 1
    Option1_Click 0
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

Private Sub Option1_Click(Index As Integer)
    Option1(0).Enabled = nAdd <> 0
    If Not (Option1(1).Value Or Option1(2).Value) Then
        Option1(0).Value = True
    End If
End Sub
