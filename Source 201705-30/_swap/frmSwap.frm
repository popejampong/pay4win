VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSwap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swap Date Entry"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   10125
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   60
      TabIndex        =   15
      Top             =   1020
      Width           =   10020
      Begin VB.CommandButton Command2 
         Caption         =   "A&pply"
         Height          =   660
         Left            =   8025
         Picture         =   "frmSwap.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmSwap.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmSwap.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmSwap.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmSwap.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmSwap.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmSwap.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   9000
         Picture         =   "frmSwap.frx":B28E
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
         Picture         =   "frmSwap.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmSwap.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmSwap.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
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
      Left            =   1590
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:CTRL_NO"
      Top             =   75
      Width           =   915
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1590
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE1"
      Top             =   390
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56623104
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   1590
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE2"
      Top             =   705
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56623104
      CurrentDate     =   38623
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Control Number"
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
      Top             =   105
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date to Swap"
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
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Swap Date"
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
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   405
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSwap
' description   :   Swap Date module
' programmer    :   _-=[ srm ]=-_
' date          :   17 nov 2006

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String, _
        oTempADO As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrSave
    Dim cString As String
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Swap Date entry?", vbYesNoCancel, "Swap Date Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA7927", "CTRL_NO=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Control Number already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA7927"), oTempADO, True
                    Script2File InsertFields(Me, "PA7927")
                    
                    Log2Audit Name, "ADD CONTROL NO -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "PA7927", "PA7927.CTRL_NO=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA7927", "PA7927.CTRL_NO=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT CONTROL NO -->" & Trim(Text1.Text)
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
            
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False
    
    If Text1.Text <> cSeries Then ResetSeries "SWAP", cSeries
    
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "CTRL_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
endsave:
    Exit Sub
    
ErrSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "SWAP", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "CTRL_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        
            If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
        End If
        
    End If
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrApply
    
    Dim lProceed As Boolean, _
        cSqlStmt As String, _
        cString As String
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If

    If lProceed Then
        If MsgBox("Apply this Swap date entry?", vbYesNo) = vbYes Then
        
            cSqlStmt = "update pa7927 set status=1, " & _
                       "                  date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                       " where ctrl_no=" & cQuote & Text1.Text & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "CTRL_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
        
            If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this swap date entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If

    Exit Sub
    
ErrApply:
    ErrorMsg Err.Number, Err.Description, "Apply Swap Entry #" & Text1.Text, Name
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
    frmLookup.showPopup 15
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "CTRL_NO='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then GetFields Me, oTempADO
    End If
    Frame2.Enabled = True
End Sub

Private Sub Command7_Click()
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    cSeries = GenerateSeries("SWAP")
    While IfExists("PA7927", "PA7927.CTRL_NO=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("SWAP")
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
        DTPicker1.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDeptDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA7927 WHERE CTRL_NO=" & cQuote & Text1.Text & cQuote, oTempADO, True
        
        Log2Audit Name, "DELETE SWAP DATE CONTROL NUMBER " & Trim(Text1.Text)
        
        Script2File "DELETE FROM PA7927 WHERE CTRL_NO=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
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
    
    OpenQueryDNS "SELECT * FROM PA7927 ORDER BY CTRL_NO", oTempADO, False
    GetFields Me, oTempADO
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
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
