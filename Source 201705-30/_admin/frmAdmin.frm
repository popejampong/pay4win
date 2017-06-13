VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Administrator File Entry"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   11310
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   11310
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmAdmin.frx":08CA
      Left            =   1605
      List            =   "frmAdmin.frx":08DA
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Tag             =   "1"
      ToolTipText     =   "NUM:GROUPID"
      Top             =   2490
      Width           =   2850
   End
   Begin VB.TextBox Text10 
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
      Left            =   1620
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "TXT:POSITION"
      Top             =   2190
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "System User"
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "NUM:SYSUSER"
      Top             =   2325
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "with Authority"
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "NUM:USERLEVEL"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   1680
      TabIndex        =   38
      Top             =   0
      Width           =   7455
      Begin VB.OptionButton Option1 
         Caption         =   "All User"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Non-System User"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "System User"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " User Status "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   7095
      TabIndex        =   33
      Top             =   585
      Width           =   4110
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAdmin.frx":092D
         Left            =   1140
         List            =   "frmAdmin.frx":0937
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "NUM:STATUS"
         Top             =   1125
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "TXT:TIME"
         Top             =   525
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "1"
         ToolTipText     =   "TXT:WSID"
         Top             =   225
         Width           =   660
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "ddd d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   23
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_LOG"
         Top             =   825
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd - MMM dd, yyyy"
         Format          =   56295427
         CurrentDate     =   38338
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   60
         TabIndex        =   37
         Top             =   1170
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Date Logged"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   60
         TabIndex        =   36
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Logged Time"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   60
         TabIndex        =   35
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Station ID"
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   60
         TabIndex        =   34
         Top             =   285
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   1620
      TabIndex        =   32
      Top             =   2775
      Width           =   9585
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7410
         Picture         =   "frmAdmin.frx":094D
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6570
         Picture         =   "frmAdmin.frx":22CF
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "19"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5730
         Picture         =   "frmAdmin.frx":3C51
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "18"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2730
         Picture         =   "frmAdmin.frx":55D3
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "12"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1890
         Picture         =   "frmAdmin.frx":6F55
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "14"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   1050
         Picture         =   "frmAdmin.frx":8FAF
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "13"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8490
         Picture         =   "frmAdmin.frx":A931
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4890
         Picture         =   "frmAdmin.frx":C2B3
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "17"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3810
         Picture         =   "frmAdmin.frx":DC35
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "15"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   210
         Picture         =   "frmAdmin.frx":F5B7
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "11"
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   5295
      TabIndex        =   28
      Top             =   585
      Width           =   1740
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   60
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1035
         Width           =   1590
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   60
         PasswordChar    =   "*"
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "TXT:PWORD"
         Top             =   390
         Width           =   1590
      End
      Begin VB.Label Label7 
         Caption         =   "Confirm Password"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   31
         Top             =   825
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "New Password:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   30
         Top             =   180
         Width           =   1380
      End
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
      Left            =   1620
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "TXT:LASTNAME"
      Top             =   1890
      Width           =   3615
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
      Left            =   1620
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:MNAME"
      Top             =   1590
      Width           =   3615
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
      Left            =   1620
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:FIRSTNAME"
      Top             =   1290
      Width           =   3615
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
      Left            =   1620
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:USERID"
      Top             =   990
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "DAT:DATEREG"
      Top             =   690
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56295424
      CurrentDate     =   38338
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
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
      Height          =   360
      Left            =   75
      TabIndex        =   45
      Top             =   2550
      Width           =   1425
   End
   Begin VB.Label Label15 
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
      Height          =   360
      Left            =   75
      TabIndex        =   43
      Top             =   2235
      Width           =   1425
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group by"
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
      Height          =   360
      Left            =   0
      TabIndex        =   42
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Registered"
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
      Height          =   360
      Left            =   75
      TabIndex        =   29
      Top             =   750
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Height          =   360
      Left            =   75
      TabIndex        =   27
      Top             =   1935
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
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
      Height          =   360
      Left            =   75
      TabIndex        =   26
      Top             =   1635
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Height          =   360
      Left            =   75
      TabIndex        =   25
      Top             =   1335
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Height          =   360
      Left            =   75
      TabIndex        =   24
      Top             =   1035
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   3795
      Left            =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Cost Accounting System
' module        :   frmAdmin
' description   :   Administrator module
' programmer    :   _-=[ srm ]=-_
' date          :   13 jan 2005

Option Explicit
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset
    Dim nAdd As Integer
    Dim nIndex As Integer

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        Text6.Text = Text5.Text
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrAdminSave
    Dim cString, _
        cSqlStmt As String
    
    cString = Text1.Text
    
    If Text5.Text <> Text6.Text Then
        MsgBox "Please confirm your password again!", vbCritical, App.Title
        Text6.SetFocus
    Else
        Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " user entry?", vbYesNoCancel, "User Entry...")
            Case vbYes
                If nAdd = 1 Then
                    If IfExists("pa2360", "USERID=" & cQuote & Text1.Text & cQuote) Then
                        MsgBox "User ID already exists!", vbOKOnly, App.Title
                        GoTo endsave
                    Else
'                        OpenQueryDNS InsertFields(Me, "pa2360"), oTempADO, True
'                        Script2File InsertFields(Me, "pa2360")      ' --> added 20050311
                        
                        cSqlStmt = "INSERT INTO pa2360(USERID,DATEREG,FIRSTNAME,MNAME,LASTNAME,POSITION,`PASSWORD`,USERLEVEL,`STATUS`,GROUPID,SYSUSER)VALUES(" & _
                                   cQuote & Text1.Text & cQuote & "," & _
                                   cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                   cQuote & EncodeStr(Text2.Text) & cQuote & "," & _
                                   cQuote & EncodeStr(Text3.Text) & cQuote & "," & _
                                   cQuote & EncodeStr(Text4.Text) & cQuote & "," & _
                                   cQuote & Text10.Text & cQuote & "," & _
                                   "AES_ENCRYPT(" & cQuote & EncodeStr(Text5.Text) & cQuote & ",UCASE(USERID))," & _
                                   Check1.Value & "," & _
                                   Combo1.ListIndex & "," & _
                                   Combo2.ListIndex & "," & _
                                   Check2.Value & ")"
                        OpenQueryDNS cSqlStmt, oTempADO, True
                        Script2File cSqlStmt
                        
                        Log2Audit Name, "ADD " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text4.Text)))
                        
                    End If
                Else
'                    OpenQueryDNS EditField(Me, "pa2360", "pa2360.USERID=" & cQuote & Text1.Text & cQuote), oTempADO, True
'                    Script2File EditField(Me, "pa2360", "pa2360.USERID=" & cQuote & Text1.Text & cQuote)       ' --> added 20050311
                    
                    cSqlStmt = "UPDATE pa2360 SET DATEREG=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                               "                  FIRSTNAME = " & cQuote & EncodeStr(Text2.Text) & cQuote & "," & _
                               "                  MNAME = " & cQuote & EncodeStr(Text3.Text) & cQuote & "," & _
                               "                  LASTNAME = " & cQuote & EncodeStr(Text4.Text) & cQuote & "," & _
                               "                  POSITION = " & cQuote & Text10.Text & cQuote & "," & _
                               "                  `PASSWORD` = AES_ENCRYPT(" & cQuote & EncodeStr(Text5.Text) & cQuote & ",UCASE(USERID))," & _
                               "                  USERLEVEL = " & Check1.Value & "," & _
                               "                  STATUS = " & Combo1.ListIndex & "," & _
                               "                  GROUPID = " & Combo2.ListIndex & "," & _
                               "                  SYSUSER = " & Check2.Value & _
                               " WHERE USERID = " & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, oTempADO, True
                    Script2File cSqlStmt
                    
                    Log2Audit Name, "EDIT User ID#" & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text4.Text)))
                End If
                
            Case vbNo
                cString = ""
                
            Case vbCancel
                GoTo endsave
                
        End Select
        
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
        
        If Text1.Text <> cSeries Then ResetSeries "ADMIN", cSeries
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        Text6.Enabled = False
        
        Frame4.Enabled = True
        
        Option1_Click nIndex
        
        oTempADO.Requery adAsyncFetch
        If Trim(cString) <> "" Then oTempADO.Find "USERID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
        If oTempADO.RecordCount > 0 Then
            If Not oTempADO.EOF Then
                GetFields Me, oTempADO
            Else
                oTempADO.MoveFirst
            End If
        End If
'        Text5.Text = Dekryp(Text5.Text)
        
        Text6.Text = Text5.Text
    End If
endsave:
    Exit Sub
    
ErrAdminSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "ADMIN", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Text6.Enabled = False
            
            Frame4.Enabled = True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "USERID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            Text6.Text = Text5.Text
        End If
    End If
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 1, IIf(nIndex = 0, "WHERE SYSUSER=1", IIf(nIndex = 1, "WHERE SYSUSER=0", ""))
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "USERID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            Text6.Text = Text5.Text
        End If
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    
    Text6.Enabled = True
    Text6.Text = Text5.Text
    
    Frame4.Enabled = False
    
    cSeries = GenerateSeries("ADMIN")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("pa2360", "pa2360.USERID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("ADMIN")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    ' --> modified 20050321
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        cSeries = Text1.Text
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Frame4.Enabled = False
        
        Text1.Enabled = False
        Text6.Enabled = True
'        Text5.Text = Dekryp(Text5.Text, Text1.Text)
        Text6.Text = Text5.Text
        
        
        'Revision 20150822 --> for deactivation of group combo box for not admin user
        If lSuperUser <> True Then
            OpenQueryDNS " select * from pa2360 where userid = " & cQuote & gUserID & cQuote, objdbRs, False
            If (objdbRs("groupid") = 0) Or (objdbRs("groupid") = 1) Then
                Combo2.Enabled = False
                Text5.Enabled = False
                Text6.Enabled = False
            Else
                Combo2.Enabled = True
            End If
        Else
            Combo2.Enabled = True
        End If
                        
        
        Text2.SetFocus
    End If
    ' --> end modified
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrAdminDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "User Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM pa2360 WHERE USERID=" & cQuote & Text1.Text & cQuote, oTempADO, True

        ' --> delete user right on di2798
        OpenQueryDNS "DELETE FROM pa2798 WHERE USERID=" & cQuote & Text1.Text & cQuote, oTempADO, True

        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text4.Text)))
        
        ' --> added 20050311
        Script2File "DELETE FROM pa2360 WHERE USERID=" & cQuote & Text1.Text & cQuote
        Script2File "DELETE FROM pa2798 WHERE USERID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        
        Text6.Enabled = False

        Option1_Click nIndex
'        OpenQueryDNS "SELECT * FROM pa2360", oTempADO, False
'        GetFields Me, oTempADO
''        Text5.Text = Dekryp(Text5.Text)
'        Text6.Text = Text5.Text
    End If
    
    Exit Sub
    
ErrAdminDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    K_Press KeyAscii
End Sub

Private Sub Form_Load()
    Dim cSqlStmt As String
    
    Log2Audit Name, "OPEN"
    
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    Text6.Enabled = False
    
    cSqlStmt = "SELECT USERID,DATEREG,FIRSTNAME,MNAME,LASTNAME,GROUPID,AES_DECRYPT(`PASSWORD`,UCASE(USERID)) AS PWORD, `POSITION`,USERLEVEL," & _
               " SYSUSER,`TIME`,DATE_LOG,`STATUS`,WSID, CMPID FROM PA2360 WHERE SYSUSER=1 ORDER BY USERID"
    OpenQueryDNS cSqlStmt, oTempADO, False
'    OpenQueryDNS "SELECT * FROM pa2360 WHERE SYSUSER = 1 ORDER BY USERID", oTempADO, False
    GetFields Me, oTempADO
    
'    Text5.Text = dekryp(Text5.Text, Text1.Text)
    Text6.Text = Text5.Text
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

Private Sub Option1_Click(Index As Integer)
    Dim cSqlStmt, cParam As String
    nIndex = Index
    
    cSqlStmt = "SELECT USERID,DATEREG,FIRSTNAME,MNAME,LASTNAME,GROUPID,AES_DECRYPT(`PASSWORD`,UCASE(USERID)) AS PWORD, `POSITION`,USERLEVEL," & _
               " SYSUSER,`TIME`,DATE_LOG,`STATUS`,WSID, CMPID FROM pa2360"
'    cSqlStmt = "SELECT * FROM pa2360"
    
    Select Case Index
        Case 0
            cParam = " WHERE SYSUSER = 1 "
        Case 1
            cParam = " WHERE SYSUSER = 0 "
        Case 2
            cParam = " "
    End Select
    
    CtrlPanel Me, nAdd
    OpenQueryDNS cSqlStmt & cParam & "ORDER BY USERID", oTempADO, False
    GetFields Me, oTempADO
    Text6.Text = Text5.Text
End Sub
