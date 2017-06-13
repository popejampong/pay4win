VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGenBCID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bio-Clock ID Generator"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmGenBCID.frx":0000
      Left            =   60
      List            =   "frmGenBCID.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5700
      Width           =   3450
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmGenBCID.frx":0052
      Left            =   4530
      List            =   "frmGenBCID.frx":0068
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5760
      Width           =   2250
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   8235
      TabIndex        =   2
      Top             =   5535
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Default         =   -1  'True
      Height          =   495
      Left            =   6885
      TabIndex        =   1
      Top             =   5535
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5445
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   9604
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   5505
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Device"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   5565
      Width           =   1935
   End
End
Attribute VB_Name = "frmGenBCID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmGenBCID
' description   :   Generate Bio-Clock ID number for information purposes only...
' programmer    :   _-=[ srm ]=-_
' date          :   19 May 2006

Option Explicit

Dim oTempADO As New ADODB.Recordset, _
    myArray, myArray2


Private Sub Combo2_Click()
    With MSHFlexGrid2
        .Redraw = False
        .Col = Combo2.ListIndex + 1
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With
End Sub

Private Sub Command1_Click()
    Dim nCtr As Integer, _
        cSqlStmt As String
    
    With MSHFlexGrid2
        .Redraw = False
        .Clear
        SetGridColumn myArray, MSHFlexGrid2
        For nCtr = 1 To 999
            .Rows = nCtr + 1
            .RowHeight(nCtr) = 285
            .TextMatrix(nCtr, 1) = ((Combo1.ListIndex + 1) * IIf(gCompanyID = "0002", 1000, 10000)) + nCtr
            cSqlStmt = "select a.tcid, " & _
                       "       a.empid, " & _
                       "       concat(a.lastname,', ',a.firstname,' ',if(trim(a.mname)='',' ',concat(left(a.mname,1),'.'))) as fullname, " & _
                       "       ifnull(b.posname,'') as position, " & _
                       "       ifnull(c.linename,'') as linename, " & _
                       "       if(a.active=0,'',if(a.active=1,'Resigned','FC')) as status," & _
                       "       a.emp_stat, " & _
                       "       a.firstname, " & _
                       "       a.lastname, " & _
                       "       a.emp_stat, " & _
                       "       a.active, " & _
                       "       a.depid,a.wap " & _
                       "from di3670 a left join di7670 b on a.posid=b.posid " & _
                       " left join di5463 c on a.depid=c.lineid " & _
                       "where a.tcid=" & cQuote & ((Combo1.ListIndex + 1) * IIf(gCompanyID = "0002", 1000, 10000)) + nCtr & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nCtr, 2) = objdbRs("empid")
                .TextMatrix(nCtr, 3) = objdbRs("fullname")
                .TextMatrix(nCtr, 4) = objdbRs("position")
                .TextMatrix(nCtr, 5) = objdbRs("linename")
                .TextMatrix(nCtr, 6) = objdbRs("status")
            End If
            .TopRow = nCtr
        Next nCtr
        .Redraw = True
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Log2Audit Name, "Open"
    
    myArray = Array("TXT:[TCID]:6:True", _
                    "TXT:[Emp ID]:8:True", _
                    "TXT:[Fullname]:30:True", _
                    "TXT:[Position]:20:True", _
                    "TXT:[Department]:20:True", _
                    "TXT:[Status]:30:True", _
                    "NUM:[emp stat]:1:False", _
                    "TXT:[FName]:20:False", _
                    "TXT:[LName]:20:False", _
                    "NUM:[Emp Stat]:1:False", _
                    "NUM:[Active]:1:False", _
                    "NUM:[Dep ID]:3:False", _
                    "NUM:[WAP Status]:1:False")
                    
    SetGridColumn myArray, MSHFlexGrid2
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log2Audit Name, "CLOSE"
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


