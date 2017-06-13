VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Begin VB.Form frmAutomate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DTR Automatic Upload"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   14655
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   556
      HasRightBorder  =   0   'False
      HasTopBorder    =   0   'False
      HasBottomBorder =   0   'False
      LicValid        =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded DTR"
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
         Height          =   345
         Left            =   105
         TabIndex        =   2
         Top             =   90
         Width           =   2070
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   14085
      Top             =   4320
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6297
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3570
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6297
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
   Begin ciaXPPanel.XPPanel XPPanel2 
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   3990
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   556
      HasRightBorder  =   0   'False
      HasTopBorder    =   0   'False
      HasBottomBorder =   0   'False
      LicValid        =   -1  'True
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded DTR"
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
         Height          =   345
         Left            =   105
         TabIndex        =   5
         Top             =   90
         Width           =   2070
      End
   End
End
Attribute VB_Name = "frmAutomate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'               $`````$
'             $( o  o )$
'    >------oOO--(_)--OOo------------------------------------------------------------------------------<
'    "Intelligent people can be bored. They just know a lot. But Smart people
'    are never bored, because they're always looking for something to engage
'    their minds."
'    >------oooo(O) (0)oooo----------------------------------------------------------------------------<

' project name  :   Dong-in Payroll System
' module        :   frmAutomate
' description   :   Module for monitoring of automatic downloading of bioclock data...
' programmer    :   _-=[ srm ]=-_
' date          :   3 july 2007

Option Explicit
    Dim myArray As Variant, _
        myarray1 As Variant, _
        myArray2 As Variant, _
        cSeries As String, _
        nCtr As Integer
        
    Dim oTempADO As ADODB.Recordset

Sub ShowRecords()
    Dim cSqlStmt As String, _
        cParam As String
       
    cSqlStmt = "SELECT a.empid, " & _
               " a.tcid, " & _
               " ifnull(concat(b.lastname,', ',b.firstname,' ',if(trim(b.mname)='',' ',concat(left(b.mname,1),'.'))),'unknown entry') as fullname, " & _
               " ifnull(c.linename,'') as department," & _
               " a.transdate, " & _
               " a.trantime, " & _
               " if(a.trantype=0,'In','Out') as trantype " & _
               "FROM att2000 a " & _
               " left join di3670 b on a.empid = b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " where (a.tag = 1) and (a.transdate=curdate()) " & _
               " order by a.transdate desc, a.trantime desc "
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If

    cSqlStmt = "SELECT a.empid, " & _
               " a.tcid, " & _
               " ifnull(concat(b.lastname,', ',b.firstname,' ',if(trim(b.mname)='',' ',concat(left(b.mname,1),'.'))),'') as fullname, " & _
               " a.transdate, " & _
               " a.trantime, " & _
               " if(a.trantype=0,'In','Out') as trantype, " & _
               " if(trim(a.empid)='','Undefined Employee','No Shifting Schedule') as remark " & _
               "FROM att2000 a " & _
               " left join di3670 b on a.empid = b.empid " & _
               " where (a.tag = 0) and (a.transdate=curdate()) " & _
               " order by a.transdate desc, a.trantime desc "
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid2, myarray1
    Else
        SetGridColumn myarray1, MSHFlexGrid2
    End If
    
    
'    cSqlStmt = cParam & " where a.tag = 1 " & _
'                        " order by a.transdate desc "
'    OpenQueryDNS cSqlStmt, objdbRs, False
'
'    If objdbRs.RecordCount > 0 Then
'        QueryAttach objdbRs, MSHFlexGrid2, myarray1, False
'    Else
'        SetGridColumn myarray1, MSHFlexGrid2
'    End If
'
'    OpenQueryDNS "SELECT distinct bcid from att2000", objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        QueryAttach objdbRs, MSHFlexGrid3, myArray2, False
'    Else
'        SetGridColumn myArray2, MSHFlexGrid3
'    End If
End Sub


Private Sub Form_Load()
    Log2Audit Name, "OPEN"
        
    myArray = Array("TXT:[Emp ID]:10:True", _
                    "TXT:[TCID]:8:True", _
                    "TXT:[FULLNAME]:50:True", _
                    "TXT:[DEPARTMENT]:40:True", _
                    "TXT:[Date]:15:True", _
                    "TXT:[Time]:12:True", _
                    "TXT:[Type]:8:True")
                    
    myarray1 = Array("TXT:[Emp ID]:10:True", _
                    "TXT:[TCID]:8:True", _
                    "TXT:[FULLNAME]:40:True", _
                    "TXT:[Date]:15:True", _
                    "TXT:[Time]:10:True", _
                    "TXT:[Type]:8:True", _
                    "TXT:[Remark]:50:True")
                    
    myArray2 = Array("TXT:[BCID]:10:True")

    SetGridColumn myArray, MSHFlexGrid1
    SetGridColumn myarray1, MSHFlexGrid2

    Timer1.Enabled = True
    Timer1.Interval = 5000
    
    ShowRecords
End Sub

Private Sub Timer1_Timer()
    ShowRecords
End Sub
