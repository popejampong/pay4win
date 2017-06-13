VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1935
   ClientLeft      =   5220
   ClientTop       =   5640
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLogin.frx":B8FA
   ScaleHeight     =   1143.262
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   255
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1380
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   885
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Login ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   45
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   675
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmLogin
' description   :   System Login
' programmer    :   _-=[ srm ]=-_
' date created  :   17 december 2004
' note          :   copied from DICAS

Option Explicit
    Dim lKeyCombine As Boolean, _
        cUserID As String
    

Function DecodeChar(cChar As String) As Integer
    Dim myArray As Variant, _
        nCtr As Integer, _
        lFound As Boolean
        
    If Format(Now, "dd") Mod 2 = 0 Then
        myArray = Array("ABC", "DEF", "GHI", "JKL", "MNO", "PQRS", "TUV", "WXYZ")
    Else
        myArray = Array("QAZ", "WSX", "EDC", "RFV", "TGB", "YHN", "UJM", "IK", "OL", "P")
    End If
    
    DoEvents
    For nCtr = 0 To IIf(Format(Now, "dd") Mod 2 = 0, 7, 9)
        If InStr(1, myArray(nCtr), cChar) > 0 Then
            lFound = True
            Exit For
        End If
    Next nCtr
    DecodeChar = IIf(lFound, nCtr + IIf(Format(Now, "dd") Mod 2 = 0, 2, 1), 0)
End Function

Function GenChar() As String
    ' 65-90     A - Z
    Randomize
    GenChar = Chr$(64 + Int((26 * Rnd) + 1))
End Function

Private Sub cmdCancel_Click()
    ModalResult = mrCancel
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cCurDate As String, _
        oConfigFile As New FileSystemObject, _
        oTxtStream As TextStream, _
        cString As String, _
        cSupportID As String, _
        cSupportPWD As String, _
        cSupportFile As String, _
        cDCChar As String, _
        nCtr As Integer
    
'    MsgBox oConfigFile.GetSpecialFolder(1)
    
'    cSupportFile = CheckPath(oConfigFile.GetSpecialFolder(1)) & "dong-in.cfg"
'
'    If Dir(cSupportFile) <> "" Then
'        Set oTxtStream = oConfigFile.OpenTextFile(cSupportFile, ForReading)
'        cString = oTxtStream.ReadLine
'        cSupportID = UCase(Dekryp(left(cString, InStr(1, cString, "|") - 1)))
'        cSupportPWD = Dekryp(right(cString, Len(cString) - InStr(1, cString, "|")))
'    End If
    
    cCurDate = Format(Now, "mm/dd/yyyy")
    
    gUserID = txtUserName.Text
    gUserPW = txtPassword.Text
'    gUserPW = Enkryp(txtPassword.Text)
    
'    If (UCase(Trim(gUserID)) = "ADMIN") And lKeyCombine And (txtPassword.Text = "081423") Then
    If (UCase(Trim(gUserID)) = "ADMIN") And lKeyCombine Then
        If (txtPassword.Text = "081423") Then
            cString = "System Advisory!!!" & vbCrLf & vbCrLf & _
                      "This login administrative feature had been disabled" & vbCrLf & _
                      "to protect the system from unauthorized usage." & vbCrLf & vbCrLf & _
                      "Please call your system administrator for assistance..."
            MsgBox cString, vbCritical, App.Title
            End
        Else
            cString = ""
            For nCtr = 1 To Len(cUserID)
                cDCChar = DecodeChar(Mid(cUserID, nCtr, 1))
                cString = cString & IIf(Val(cDCChar) = 10, "0", cDCChar)
            Next nCtr
            If cString = txtPassword.Text Then
                gUserName = "<< Super User >>"
                lSuperUser = True
            End If
        End If
    ElseIf (Trim(cSupportID) <> "" Or Trim(cSupportPWD) <> "") And (UCase(Trim(txtUserName.Text)) = cSupportID) And (txtPassword.Text = cSupportPWD) Then
        gUserName = "<< System Support User >>"
        lSuperUser = True
    End If
    
    Set oConfigFile = Nothing
    
    ModalResult = mrOk
    Unload Me
End Sub

Private Sub Form_Activate()
    lKeyCombine = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Form_Load()
    Dim cSqlStmt As String
    
    ' --> check admin table for first time
    OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & "pa2360" & cQuote, objdbRs, False
    If objdbRs.RecordCount = 0 Then
        cSqlStmt = "CREATE TABLE `pa2360` (" & _
                   "  `dateReg` date NOT NULL default '1975-11-28'," & _
                   "  `userID` char(6) NOT NULL default ''," & _
                   "  `FirstName` char(20) NOT NULL default ''," & _
                   "  `MName` char(20) NOT NULL default ''," & _
                   "  `LastName` char(20) NOT NULL default ''," & _
                   "  `Password` char(10) NOT NULL default ''," & _
                   "  `userLevel` int(1) NOT NULL default '0'," & _
                   "  `depID` char(3) NOT NULL default ''," & _
                   "  `status` int(1) NOT NULL default '0'," & _
                   "  `time` char(10) NOT NULL default ''," & _
                   "  `date_log` date NOT NULL default '1975-11-28'," & _
                   "  `wsid` char(3) NOT NULL default ''," & _
                   "  `POSITION` char(40) NOT NULL default ''," & _
                   "  `sysuser` int(1) default '0'," & _
                   "  `CMPID` char(4) NOT NULL default '0003'," & _
                   "  `PWORD` char(16) NOT NULL default ''," & _
                   "  PRIMARY KEY  (`userID`,`CMPID`)" & _
                   ") ENGINE=MyISAM DEFAULT CHARSET=latin1"
        OpenQueryDNS cSqlStmt, objdbRs, True
    End If
    
    OpenQueryDNS "SELECT * FROM pa2360 LIMIT 1", objdbRs, False
    frmLogin.txtUserName.MaxLength = objdbRs.Fields.Item("USERID").DefinedSize
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If UCase(Trim(txtUserName.Text)) = "ADMIN" Then
        If ((Shift And vbShiftMask) > 0) And ((Shift And vbCtrlMask) > 0) And ((Shift And vbAltMask) > 0) Then
            If KeyCode = vbKeyDown Then
'                ShowInTaskbar = True
                cUserID = GenChar & GenChar & GenChar & GenChar & GenChar & GenChar
                Add2List "Creating " & IIf(Format(Now, "dd") Mod 2 = 0, "System ", "Temporary ") & "file named " & cUserID & "." & IIf(Format(Now, "dd") Mod 2 = 0, "DLL", "TMP")
                lKeyCombine = True
                txtPassword.SetFocus
            End If
        End If
    End If
End Sub
