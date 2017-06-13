VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaxppanel.ocx"
Object = "{B8C20E45-C574-4CC7-BEC9-6392691436BE}#1.1#0"; "ciaxpspin.ocx"
Object = "{083C8784-F106-4CC2-9930-876218A6B74C}#1.1#0"; "ciaxpbutton.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaxpframe.ocx"
Begin VB.Form frmEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Entry"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text34 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4635
      TabIndex        =   145
      Tag             =   "1"
      ToolTipText     =   "TXT:S_REMARK"
      Top             =   6120
      Width           =   5100
   End
   Begin VB.TextBox Text33 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4635
      TabIndex        =   144
      Tag             =   "1"
      ToolTipText     =   "TXT:REMARK"
      Top             =   5760
      Width           =   5100
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Index           =   0
      ItemData        =   "frmEmployee.frx":0000
      Left            =   8400
      List            =   "frmEmployee.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   142
      Tag             =   "1"
      ToolTipText     =   "NUM:ERP_ACTIVE"
      Top             =   5055
      Width           =   1860
   End
   Begin ciaXPPanel.XPPanel XPPanel10 
      Height          =   1635
      Left            =   6960
      TabIndex        =   126
      Top             =   7320
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   2884
      LicValid        =   -1  'True
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   315
         Left            =   2640
         TabIndex        =   140
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         TabIndex        =   139
         Tag             =   "1"
         ToolTipText     =   "TXT:BEPWORKCENTERID"
         Top             =   540
         Width           =   840
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   315
         Left            =   3540
         TabIndex        =   133
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         TabIndex        =   132
         Tag             =   "1"
         ToolTipText     =   "TXT:COSTCENTERID"
         Top             =   900
         Width           =   1740
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         TabIndex        =   130
         Tag             =   "1"
         ToolTipText     =   "TXT:WORKCENTERID"
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   315
         Left            =   2640
         TabIndex        =   129
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3240
         TabIndex        =   141
         Top             =   585
         Width           =   2355
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Center Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   136
         Top             =   930
         Width           =   1545
      End
      Begin VB.Label Label54 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   135
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1215
         TabIndex        =   134
         Top             =   1260
         Width           =   4500
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3210
         TabIndex        =   131
         Top             =   225
         Width           =   2355
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "BEP Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   128
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Center Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   127
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "frmEmployee.frx":0021
      Left            =   4635
      List            =   "frmEmployee.frx":002B
      TabIndex        =   88
      Tag             =   "1"
      Text            =   "Combo7"
      ToolTipText     =   "NUM:LABORTYPE"
      Top             =   9045
      Width           =   1695
   End
   Begin ciaXPPanel.XPPanel XPPanel8 
      Height          =   1875
      Left            =   9870
      TabIndex        =   120
      Top             =   5385
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3307
      LicValid        =   -1  'True
      Begin VB.TextBox Text20 
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
         Left            =   210
         TabIndex        =   124
         Tag             =   "1"
         ToolTipText     =   "NUM:POS_ALLOW"
         Top             =   1050
         Width           =   1560
      End
      Begin VB.TextBox Text29 
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
         Left            =   225
         TabIndex        =   28
         Tag             =   "1"
         ToolTipText     =   "TXT:BACCNTNO"
         Top             =   510
         Width           =   2430
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pos Allowance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   123
         Top             =   855
         Width           =   1410
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   255
         TabIndex        =   122
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   210
         TabIndex        =   121
         Top             =   15
         Width           =   2685
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   4635
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_FIN"
      Top             =   5400
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56754176
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   285
      Left            =   4635
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_RES"
      Top             =   5400
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56754176
      CurrentDate     =   38623
   End
   Begin VB.TextBox Text23 
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
      Left            =   4635
      TabIndex        =   26
      Tag             =   "1"
      ToolTipText     =   "TXT:PHEALTHNUM"
      Top             =   7995
      Width           =   2235
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Select All"
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   120
      TabIndex        =   106
      Top             =   9015
      Width           =   2340
   End
   Begin ciaXPPanel.XPPanel XPPanel6 
      Height          =   435
      Left            =   4635
      TabIndex        =   100
      Top             =   15
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   767
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmEmployee.frx":004D
         Left            =   2370
         List            =   "frmEmployee.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   60
         Width           =   2790
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status Filter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   90
         TabIndex        =   102
         Top             =   105
         Width           =   2325
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel5 
      Height          =   630
      Left            =   6945
      TabIndex        =   95
      Top             =   9000
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   1111
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.TextBox Text11 
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
         Left            =   825
         TabIndex        =   90
         Tag             =   "1"
         ToolTipText     =   "TXT:SHIFTID"
         Top             =   75
         Width           =   630
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   315
         Left            =   1530
         TabIndex        =   91
         Top             =   60
         Width           =   405
      End
      Begin ciaXPButton.XPButton XPButton4 
         Height          =   525
         Left            =   4155
         TabIndex        =   99
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   926
         Caption         =   "Custom Shifting Schedule"
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   0
         TabIndex        =   125
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "08:00 AM - 05:00 PM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   765
         TabIndex        =   98
         Top             =   390
         Width           =   3180
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Regular Shift"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1980
         TabIndex        =   97
         Top             =   120
         Width           =   2190
      End
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmEmployee.frx":0093
      Left            =   6495
      List            =   "frmEmployee.frx":00A3
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "NUM:ACTIVE"
      Top             =   5055
      Width           =   1860
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   315
      Left            =   5235
      TabIndex        =   87
      Top             =   8580
      Width           =   405
   End
   Begin VB.TextBox Text16 
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
      Left            =   4635
      TabIndex        =   29
      Tag             =   "1"
      ToolTipText     =   "TXT:TAXID"
      Top             =   8595
      Width           =   570
   End
   Begin ciaXPPanel.XPPanel XPPanel2 
      Height          =   1650
      Left            =   10785
      TabIndex        =   85
      Top             =   5505
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2910
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin ciaXPButton.XPButton XPButton2 
         Height          =   525
         Left            =   45
         TabIndex        =   43
         Top             =   555
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   926
         Caption         =   "VL/SL Accrual"
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPButton.XPButton XPButton1 
         Height          =   525
         Left            =   45
         TabIndex        =   42
         Top             =   120
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   926
         Caption         =   "Deductions"
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPButton.XPButton XPButton3 
         Height          =   525
         Left            =   45
         TabIndex        =   44
         Top             =   1080
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   926
         Caption         =   "Finger Print"
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
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
      Left            =   4635
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "TXT:MNAME"
      Top             =   1830
      Width           =   3585
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
      Left            =   4635
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "TXT:LASTNAME"
      Top             =   2130
      Width           =   3585
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
      Left            =   4635
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   1230
      Width           =   855
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
      Left            =   4635
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "TXT:FIRSTNAME"
      Top             =   1530
      Width           =   3585
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
      Left            =   4635
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "TXT:DEPID"
      Top             =   3945
      Width           =   450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   5085
      TabIndex        =   61
      Top             =   3930
      Width           =   495
   End
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
      Left            =   4635
      TabIndex        =   21
      Tag             =   "1"
      ToolTipText     =   "NUM:RATE_AMT"
      Top             =   6555
      Width           =   1560
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmEmployee.frx":00CF
      Left            =   4635
      List            =   "frmEmployee.frx":00DC
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "1"
      ToolTipText     =   "NUM:PAYSTATUS"
      Top             =   6870
      Width           =   1560
   End
   Begin VB.TextBox Text8 
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
      Left            =   8175
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "NUM:COLA_AMT"
      Top             =   6555
      Width           =   1560
   End
   Begin VB.TextBox Text9 
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
      Left            =   4635
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "TXT:PAGIBIGNO"
      Top             =   7395
      Width           =   2235
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
      Left            =   4635
      TabIndex        =   25
      Tag             =   "1"
      ToolTipText     =   "TXT:SSNUM"
      Top             =   7695
      Width           =   2235
   End
   Begin VB.TextBox Text12 
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
      Left            =   4635
      TabIndex        =   27
      Tag             =   "1"
      ToolTipText     =   "TXT:TIN"
      Top             =   8295
      Width           =   2235
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmEmployee.frx":00FB
      Left            =   4635
      List            =   "frmEmployee.frx":0105
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "NUM:SEX"
      Top             =   2760
      Width           =   1560
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmEmployee.frx":0117
      Left            =   4635
      List            =   "frmEmployee.frx":0121
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "NUM:ISUNION"
      Top             =   4620
      Width           =   900
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmEmployee.frx":012E
      Left            =   4635
      List            =   "frmEmployee.frx":013B
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "NUM:EMP_STAT"
      Top             =   5055
      Width           =   1860
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   315
      Left            =   5085
      TabIndex        =   60
      Top             =   4245
      Width           =   495
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
      Left            =   4635
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "TXT:POSID"
      Top             =   4260
      Width           =   450
   End
   Begin ciaXPFrame.XPFrame XPFrame2 
      Height          =   1665
      Left            =   10560
      TabIndex        =   45
      Top             =   3630
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   2937
      Caption         =   " Year-to-Date Info "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         Tag             =   "1"
         ToolTipText     =   "NUM:YTD_WTAX"
         Top             =   1260
         Width           =   1785
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         Tag             =   "1"
         ToolTipText     =   "NUM:YTD_BASIC"
         Top             =   810
         Width           =   1785
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         Tag             =   "1"
         ToolTipText     =   "NUM:YTD_GROSS"
         Top             =   345
         Width           =   1785
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "YTD W/Tax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   51
         Top             =   1065
         Width           =   990
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Basic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   50
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Gross"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   49
         Top             =   165
         Width           =   990
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   1350
      Left            =   10545
      TabIndex        =   52
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   2381
      Caption         =   " VL / SL Availment "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   54
         Tag             =   "1"
         Text            =   "Text16"
         ToolTipText     =   "NUM:SL_USE"
         Top             =   885
         Width           =   600
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   53
         Tag             =   "1"
         Text            =   "Text16"
         ToolTipText     =   "NUM:VL_USE"
         Top             =   540
         Width           =   600
      End
      Begin ciaXPSpin.XPSpin XPSpin1 
         Height          =   315
         Left            =   390
         TabIndex        =   40
         Tag             =   "1"
         ToolTipText     =   "NUM:VL_AVAIL"
         Top             =   540
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         MouseIcon       =   "frmEmployee.frx":015A
         LicValid        =   -1  'True
      End
      Begin ciaXPSpin.XPSpin XPSpin2 
         Height          =   315
         Left            =   390
         TabIndex        =   41
         Tag             =   "1"
         ToolTipText     =   "NUM:SL_AVAIL"
         Top             =   885
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         MouseIcon       =   "frmEmployee.frx":0176
         LicValid        =   -1  'True
      End
      Begin VB.Line Line2 
         X1              =   105
         X2              =   1900
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Alloted"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   135
         TabIndex        =   58
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Accrued"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1290
         TabIndex        =   57
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "SL"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   945
         Width           =   540
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "VL"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   540
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   2010
      Left            =   10485
      TabIndex        =   59
      Top             =   60
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   3545
      LicValid        =   -1  'True
      Begin VB.Image Image1 
         Height          =   1890
         Left            =   60
         ToolTipText     =   "Click to change picture"
         Top             =   60
         Width           =   1965
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel3 
      Height          =   3135
      Left            =   4635
      TabIndex        =   81
      Top             =   825
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   5530
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.TextBox Text26 
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
         Left            =   3510
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "TXT:ADD_CITY"
         Top             =   2280
         Width           =   1185
      End
      Begin VB.TextBox Text28 
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
         Left            =   0
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "TXT:TEL_NUM"
         Top             =   2805
         Width           =   1980
      End
      Begin VB.TextBox Text27 
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
         Left            =   2265
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "TXT:ADD_BRGY"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text25 
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
         Left            =   0
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "TXT:ADD_NO"
         Top             =   2280
         Width           =   2235
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Import"
         Height          =   315
         Left            =   885
         TabIndex        =   113
         Top             =   375
         Width           =   900
      End
      Begin VB.TextBox Text24 
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
         Left            =   3675
         TabIndex        =   112
         Tag             =   "1"
         ToolTipText     =   "TXT:REF_EMPID"
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text17 
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
         Left            =   2730
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "TXT:TCID"
         Top             =   405
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   -15
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "DAT:BIRTHDAY"
         Top             =   1620
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56754176
         CurrentDate     =   38623
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_HIRE"
         Top             =   75
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56754176
         CurrentDate     =   38623
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3525
         TabIndex        =   119
         Top             =   2595
         Width           =   1155
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Bataan"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4710
         TabIndex        =   117
         Top             =   2325
         Width           =   1290
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Brgy"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2295
         TabIndex        =   116
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "No. and Street"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   15
         TabIndex        =   115
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bioclock ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1185
         TabIndex        =   93
         Top             =   465
         Width           =   1455
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel4 
      Height          =   765
      Left            =   4635
      TabIndex        =   82
      Top             =   6495
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   1349
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.TextBox Text21 
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
         Left            =   1635
         TabIndex        =   103
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "NUM:WAP"
         Top             =   75
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.O.L.A."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2040
         TabIndex        =   86
         Top             =   135
         Width           =   1410
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame5 
      Height          =   9030
      Left            =   60
      TabIndex        =   104
      Top             =   30
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   15928
      Alignment       =   2
      Caption         =   " Department List "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8580
         Left            =   105
         TabIndex        =   105
         Top             =   240
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   15134
         _Version        =   393216
         GridColor       =   10416117
         GridColorUnpopulated=   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
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
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   3690
      TabIndex        =   94
      Top             =   9600
      Width           =   9015
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7065
         Picture         =   "frmEmployee.frx":0192
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmEmployee.frx":1B14
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmEmployee.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmEmployee.frx":4E18
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmEmployee.frx":679A
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmEmployee.frx":811C
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8025
         Picture         =   "frmEmployee.frx":9A9E
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmEmployee.frx":B420
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmEmployee.frx":CDA2
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmEmployee.frx":E724
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel7 
      Height          =   435
      Left            =   4635
      TabIndex        =   107
      Top             =   4575
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   767
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   108
         Tag             =   "1"
         Text            =   "Text16"
         ToolTipText     =   "NUM:UL_USE"
         Top             =   60
         Visible         =   0   'False
         Width           =   600
      End
      Begin ciaXPSpin.XPSpin XPSpin3 
         Height          =   315
         Left            =   2385
         TabIndex        =   16
         Tag             =   "1"
         ToolTipText     =   "NUM:UL_AVAIL"
         Top             =   60
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         MouseIcon       =   "frmEmployee.frx":100A6
         LicValid        =   -1  'True
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Accrued"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3240
         TabIndex        =   110
         Top             =   105
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Alloted"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1140
         TabIndex        =   109
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   285
      Left            =   4635
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "DAT:DATEREG"
      Top             =   495
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56754176
      CurrentDate     =   38623
   End
   Begin VB.Label Label60 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Remarks"
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
      Left            =   3120
      TabIndex        =   146
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label label59 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   3720
      TabIndex        =   143
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ERP Position Code"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8520
      TabIndex        =   138
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "ERP Position Code"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8520
      TabIndex        =   137
      Top             =   4005
      Width           =   1365
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No / CP No"
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
      Left            =   3120
      TabIndex        =   118
      Top             =   3645
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Finished"
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
      Left            =   3105
      TabIndex        =   84
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Resigned"
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
      Left            =   3105
      TabIndex        =   83
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   3120
      TabIndex        =   114
      Top             =   3105
      Width           =   1455
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Philhealth No."
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
      Left            =   3105
      TabIndex        =   111
      Top             =   8055
      Width           =   1455
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Type"
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
      Left            =   3105
      TabIndex        =   96
      Top             =   9075
      Width           =   1455
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "SRM"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5685
      TabIndex        =   92
      Top             =   8655
      Width           =   1290
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code"
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
      Left            =   3105
      TabIndex        =   89
      Top             =   8655
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   3105
      TabIndex        =   80
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3105
      TabIndex        =   79
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   3105
      TabIndex        =   78
      Top             =   1290
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   3105
      TabIndex        =   77
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   3105
      TabIndex        =   76
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Hire"
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
      Left            =   3105
      TabIndex        =   75
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      Left            =   3105
      TabIndex        =   74
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   3105
      TabIndex        =   73
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5595
      TabIndex        =   72
      Top             =   4005
      Width           =   2745
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   3105
      TabIndex        =   71
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Status"
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
      Left            =   3105
      TabIndex        =   70
      Top             =   6945
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pag-IBIG No."
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
      Left            =   3105
      TabIndex        =   69
      Top             =   7455
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Number"
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
      Left            =   3105
      TabIndex        =   68
      Top             =   7755
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIN Number"
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
      Left            =   3105
      TabIndex        =   67
      Top             =   8355
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday"
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
      Left            =   3105
      TabIndex        =   66
      Top             =   2475
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Left            =   3105
      TabIndex        =   65
      Top             =   2790
      Width           =   1455
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Union Member"
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
      Left            =   3105
      TabIndex        =   64
      Top             =   4695
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Empl. Status"
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
      Left            =   3105
      TabIndex        =   63
      Top             =   5085
      Width           =   1455
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5595
      TabIndex        =   62
      Top             =   4320
      Width           =   4245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   9540
      Left            =   3090
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmEmployee --> Employee Module
' description   :   Module for Maintenance of Employee
' programmer    :   _-=[ srm ]=-_
' date          :   27 jun 2005
' note          :   copied from DICAS

Option Explicit
    Dim nAdd As Integer, _
        cSeries As String, _
        cParam As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant, _
        cEmpiId As String
    
Sub add2LstBox(ByVal oRecordSet As ADODB.Recordset, ByVal oListBox As ListView, ByVal aField As Variant)
    Dim lstItem As ListItem
    
    If oRecordSet.RecordCount > 0 Then
        oListBox.ListItems.Clear
        While Not oRecordSet.EOF
            Set lstItem = oListBox.ListItems.Add()
            lstItem.Text = objdbRs(aField(1))
            lstItem.SubItems(1) = objdbRs(aField(0))
            oRecordSet.MoveNext
        Wend
    End If
End Sub

Sub SetFilter(ByVal cValue As String, nEmpStatus As Integer)
    cParam = cValue
End Sub
    
Sub CreateFinish()
    Dim cSqlStmt As String
        
    cSqlStmt = " INSERT INTO PA3674(empid, time_history, date_history, datereg, date_hire, date_fin ,date_res," & _
               " depid,posid, rate_amt, active, emp_stat, cmpid)VALUES(" & _
               cQuote & Text1.Text & cQuote & "," & _
               cQuote & Format(Time, "hh:ss:mm") & cQuote & "," & _
               cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
               cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & "," & _
               cQuote & Format(DTPicker4.Value, "yyyy-mm-dd") & cQuote & "," & _
               cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "," & _
               cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "," & _
               cQuote & Text5.Text & cQuote & "," & _
               cQuote & Text6.Text & cQuote & "," & _
               cQuote & Text7.Text & cQuote & "," & _
               cQuote & Combo5.ListIndex & cQuote & "," & _
               cQuote & Combo2.ListIndex & cQuote & "," & _
               cQuote & gCompanyID & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, True
End Sub
    
    
Sub ShowRecords()
    Dim cSqlStmt As String

'    cSqlStmt = "SELECT A.DEDID, IFNULL(B.DEDNAME,'Undefined Deduction') as DEDNAME, if(IFNULL(B.AUTO_DED,0)=1,'Auto-Compute',A.DEF_AMT) as DEF_AMT, A.ACC_AMT,A.CUT_OFF_AMT, if(A.PERIOD1=1,'Yes','No'), if(A.PERIOD2=1,'Yes','No'), IFNULL(B.AUTO_DED,0) AS AUTO_DED, IFNULL(B.FIX_DED,0) AS FIX_DED " & _
'               " FROM DI3673 A LEFT JOIN PA3330 B ON A.DEDID=B.DEDID WHERE A.EMPID=" & cQuote & Text1.Text & cQuote & _
'               " ORDER BY A.DEDID"
'    DoEvents
'    OpenQueryDNS cSqlStmt, objdbRs, False
''    If objdbRs.RecordCount = 0 Then
''        OpenQueryDNS "SELECT DEDID, DEDNAME, IF(AUTO_DED=1,'Auto-Compute',DEF_AMT) AS DEF_AMT, 0, CUT_OFF_AMT, if(PERIOD1=1,'Yes','No'), if(PERIOD2=1,'Yes','No'), AUTO_DED, FIX_DED FROM PA3330 ORDER BY DEDID", objdbRs, False
''    End If
''
'    If objdbRs.RecordCount > 0 Then
'        QueryAttach objdbRs, MSHFlexGrid1, myArray
'    Else
'        MSHFlexGrid1.Clear
'        SetGridColumn myArray, MSHFlexGrid1
'    End If
    
    OpenQueryDNS "SELECT * FROM DI5463 WHERE LINEID=" & cQuote & Text5.Text & cQuote, objdbRs, False
    Label14.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("LINENAME"), "")
    
    Select Case IIf(objdbRs.RecordCount > 0, objdbRs("ERPPOSCODE"), 7)
        Case 0 ' "A"
            Label57.Caption = "A"
        Case 1 ' "B"
            Label57.Caption = "B"
        Case 2 ' "C"
            Label57.Caption = "C"
        Case 3 ' "D"
            Label57.Caption = "D"
        Case 4 ' "E"
            Label57.Caption = "E"
        Case 5 ' "F"
            Label57.Caption = "F"
        Case 6 ' "G"
            Label57.Caption = "G"
        Case 7 ' "Z"
            Label57.Caption = "Z"
    End Select
    
    OpenQueryDNS "SELECT * FROM DI7670 WHERE POSID=" & cQuote & Text6.Text & cQuote, objdbRs, False
    Label20.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("POSNAME"), "")
    
    ' --> Tax
    OpenQueryDNS "SELECT TAXCODE FROM PA8290 WHERE TAXID=" & cQuote & Text16.Text & cQuote, objdbRs, False
    Label35.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TAXCODE"), "")
    
    ' --> Shift
    OpenQueryDNS "SELECT `DESCRIPTION`,CONCAT(TIME_FORMAT(TIME1,'%h:%i %p'),' - ',TIME_FORMAT(TIME2,'%h:%i %p')) AS `TIME` FROM PA74380 WHERE SHIFTID=" & cQuote & Text11.Text & cQuote, objdbRs, False
    Label32.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    Label33.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TIME"), "")
      
     ' ---> Work Center 201207-25
    OpenQueryDNS "select * from pa97722 where workcenterid=" & cQuote & Text31.Text & cQuote, objdbRs, False
    Label53.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
     ' ---> BEP Work Center 20130216
    OpenQueryDNS "select * from pa97722 where workcenterid=" & cQuote & Text32.Text & cQuote, objdbRs, False
    Label58.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
    OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text30.Text & cQuote, objdbRs, False
    Label50.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
    
    Command2.Enabled = nAdd <> 0
    Command3.Enabled = nAdd <> 0
    Command4.Enabled = nAdd <> 0
    Command6.Enabled = nAdd <> 0
    Command12.Enabled = nAdd <> 0
    Command14.Enabled = nAdd <> 0
    Command15.Enabled = nAdd <> 0
    Command13.Enabled = nAdd <> 0
    
    Text29.Enabled = False
    Text20.Enabled = False
    
    Combo2_Click
    Combo5_Click
    Combo4_Click
End Sub

Private Sub Check1_Click()
    cParam = ""
    Combo6_Click
    MSHFlexGrid1.Enabled = Check1.Value <> vbChecked
End Sub

Private Sub Combo2_Click()
    If nAdd = 0 Then Exit Sub
    If gCompanyID <> "0019" Then
        Select Case Combo2.ListIndex
            Case 0, 1       ' --> WAP/Contractual
                DTPicker2.Visible = True
                'DTPicker2.Value = DateDiff("d", 1, DateAdd("m", IIf(Combo2.ListIndex = 0, 3, IIf(Val(Text21.Text) = 0, 5, 3)), DTPicker4.Value))
                DTPicker2.Value = DateDiff("d", 0, DateAdd("m", IIf(Combo2.ListIndex = 0, 3, IIf(Val(Text21.Text) = 0, 5, 3)), DTPicker4.Value))
                If nAdd = 1 Then
                    If Val(Text21.Text) = 0 Then
                        Text7.Text = IIf(Combo2.ListIndex = 0, gBasicRate * 0.75, gBasicRate)
                        Text8.Text = IIf(Combo2.ListIndex = 0, gColaAmt * 0.75, gColaAmt)
                    End If
                End If
            Case 2
                DTPicker2.Visible = False
        End Select
    Else
    
        Select Case Combo2.ListIndex
            Case 0       ' --> WAP
                DTPicker2.Visible = False
            Case 1       ' --> Contractual
                DTPicker2.Visible = True
                DTPicker2.Enabled = False
                If nAdd = 1 Then
                    If Val(Text21.Text) = 0 Then
                        Text7.Text = gBasicRate
                        Text8.Text = gColaAmt
                    End If
                End If
            Case 2
                DTPicker2.Visible = False
        End Select
    End If
End Sub

Private Sub Combo4_Click()
'    Label38.Visible = Combo4.ListIndex = 0
'    Label39.Visible = Combo4.ListIndex = 0
'    XPSpin3.Visible = Combo4.ListIndex = 0
'    Text22.Visible = Combo4.ListIndex = 0
End Sub

Private Sub Combo5_Click()
    Dim cParam As String
    
    Select Case Combo5.ListIndex
        Case 0      ' --> Active
            Label12.Visible = False
            DTPicker3.Visible = False
'          For ERP  201310-02
            label59.Visible = False
            Text33.Visible = False
            Text34.Visible = False
            Label60.Visible = False

            If oTempADO.RecordCount > 0 Then
                If gCompanyID <> "0019" Then
                    Label28.Visible = oTempADO("emp_stat") <> 2
                    DTPicker2.Visible = oTempADO("emp_stat") <> 2
                Else
                    Label28.Visible = True
                    DTPicker2.Visible = True
                    
'          For ERP  201310-02
                    label59.Visible = True
                    Text33.Visible = True
                    Label60.Visible = True
                    Text34.Visible = True
                    
                End If
'                If (nAdd <> 0) Then --> remarked 20060511
                If (nAdd = 2) Then
                    If MsgBox("Is employee re-hired?", vbYesNo, App.Title) = vbNo Then
                        If gCompanyID <> "0019" Then
                            If DTPicker2.Visible Then DTPicker2.SetFocus
                        Else
                            DTPicker2.Value = DateAdd("yyyy", 2, DTPicker4.Value)
                            'DTPicker2.Value = "2012-09-15"
                            Combo5.SetFocus
                        End If
                        
                    Else
                        DTPicker4.Value = Now
                        Combo2_Click
                    End If
                End If
            Else
                Label28.Visible = False
                DTPicker2.Visible = False
            End If
        Case 1, 3   ' --> Resigned/Terminated
            Label12.Visible = True
            DTPicker3.Visible = True
            If nAdd <> 0 Then DTPicker3.Value = Now
            
            Label28.Visible = False
            DTPicker2.Visible = False
            
'             For ERP  201310-02
            
            label59.Visible = True
            Text33.Visible = True
            Label60.Visible = True
            Text34.Visible = True
            
        Case 2      ' --> Finished
            Label12.Visible = False
            DTPicker3.Visible = False
            
            Label28.Visible = True
            DTPicker2.Visible = True
            
'          For ERP  201310-02
            
            label59.Visible = True
            Text33.Visible = True
            Text34.Visible = False
            Label60.Visible = False
            
            If gCompanyID <> "0019" Then
                If nAdd <> 0 Then DTPicker2.Value = Now
            Else
                If nAdd <> 0 Then DTPicker2.Value = "2012-09-15"
            End If
            
    End Select
End Sub

Private Sub Combo6_Click()
    Dim cSqlStmt As String
    
    If nAdd <> 0 Then Exit Sub
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    cSqlStmt = "SELECT * FROM DI3670 " & _
               IIf(Trim(cParam) <> "", " WHERE DEPID IN " & cParam, "") & _
               IIf(Combo6.ListIndex = 2, "", IIf(Trim(cParam) <> "", " AND ", " WHERE ") & " ACTIVE IN (" & IIf(Combo6.ListIndex = 0, "0", "1,2") & ")") & _
               " ORDER BY EMPID"
    OpenQueryDNS cSqlStmt, oTempADO, False
    GetFields Me, oTempADO
    
    ShowRecords
End Sub







Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        ShowRecords
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrEmplSave
    Dim cString As String, _
        cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset
    
    cString = Text1.Text
    
    If Combo1.ListIndex = -1 Then
        MsgBox "Please specify employment term!", vbCritical, App.Title
        Combo1.SetFocus
        GoTo endsave
    End If
    
    If Combo2.ListIndex = -1 Then
        MsgBox "Please specify employment status!", vbCritical, App.Title
        Combo2.SetFocus
        GoTo endsave
    End If
    
    If Text30.Text = "" Then
        MsgBox "Please Specify Cost Center ID"
        Text30.SetFocus
        GoTo endsave
    End If
    
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " employee entry?", vbYesNoCancel, "Employee Entry...")
        Case vbYes
            If nAdd = 1 Then
            
                If IfExists("DI3670", "EMPID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Employee ID already exists!", vbOKOnly, App.Title
                    GoTo endsave
                Else
'                    MsgBox InsertFields(Me, "DI3670")
                    OpenQueryDNS InsertFields(Me, "DI3670"), oTempADO, True
                    Script2File InsertFields(Me, "DI3670")      ' --> added 20050311
                    
                    Log2Audit Name, "ADD " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text4.Text)))
                    
                    ' --> save shifting sched if there's any... 20070108
                    cSqlStmt = "select a.periodid, b.date, b.shiftid, " & _
                               "  ifnull(d.description,'') as description, " & _
                               "  ifnull(d.time1,'') as time1, " & _
                               "  ifnull(d.time2,'') as time2, " & _
                               "  ifnull(d.remark, b.remark) As remark " & _
                               "from di546370 a left join di546373 b on a.sched_no=b.sched_no and b.date >= " & cQuote & Format(DTPicker4.Value, "yyyy-mm-dd") & cQuote & _
                               "  left join pa7730 c on a.periodid=c.periodid " & _
                               "  left join pa74380 d on b.shiftid=d.shiftid " & _
                               "where a.depid=" & cQuote & Text5.Text & cQuote & " and a.status=1 " & _
                               "  and " & cQuote & Format(DTPicker4.Value, "yyyy-mm-dd") & cQuote & " between c.date_start and c.date_end "
                    OpenQueryDNS cSqlStmt, oRecordSet, False
                    If oRecordSet.RecordCount > 0 Then
                    
                        ShowProgress 0
                        
                        While Not oRecordSet.EOF
                        
                            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
                            
                            cSqlStmt = "insert into di36770(empid, periodid, `date`, shiftid, `description`, time1, time2, `remark`)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & oRecordSet("periodid") & cQuote & "," & _
                                       cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & oRecordSet("shiftid") & cQuote & "," & _
                                       cQuote & oRecordSet("description") & cQuote & "," & _
                                       cQuote & oRecordSet("time1") & cQuote & "," & _
                                       cQuote & oRecordSet("time2") & cQuote & "," & _
                                       cQuote & oRecordSet("remark") & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            oRecordSet.MoveNext
                        Wend
                        
                        ShowProgress 4
                        
                    End If
                    ' --> end of saving shifting sched...
                End If
                
            Else
                OpenQueryDNS EditField(Me, "DI3670", "DI3670.EMPID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "DI3670", "DI3670.EMPID=" & cQuote & Text1.Text & cQuote)       ' --> added 20050311
                
                Log2Audit Name, "EDIT Employee ID#" & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text4.Text)))
                
'                If Combo5.ListIndex > 0 Then CreateFinish
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
            
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321

    If Text1.Text <> cSeries Then ResetSeries "EMPL", cSeries

    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    Combo6.Enabled = True
    MSHFlexGrid1.Enabled = True

    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    If (oTempADO.EOF) And (oTempADO.RecordCount > 0) Then oTempADO.MoveFirst
    GetFields Me, oTempADO
    ShowRecords

endsave:
    Set oRecordSet = Nothing
    
    Exit Sub
    
ErrEmplSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
        OpenQueryDNS "update di2340 set dtr_update=1", objdbRs, True
        Script2File "update di2340 set dtr_update=1"
    Else
    
        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Log2Audit Name, "Edit Transation to EmpID #" & cString & " fullname " & Trim(Text3.Text) & ", " & Trim(Text2.Text) & " " & left(Trim(Text4.Text), 2) & ". "
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            
            If Text1.Text <> cSeries Then ResetSeries "EMPL", cSeries
            
            ' --> delete custom deduction here if ADD mode...
            If nAdd = 1 Then
                OpenQueryDNS "DELETE FROM DI3673 WHERE EMPID=" & cQuote & Text1.Text & cQuote, objdbRs, True
                Script2File "DELETE FROM DI3673 WHERE EMPID=" & cQuote & Text1.Text & cQuote
            End If
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Command2.Enabled = False
            Command3.Enabled = False
            
            Label14.Caption = ""
            Label57.Caption = ""
            Label20.Caption = ""
            
            Combo6.Enabled = True
            MSHFlexGrid1.Enabled = True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords
            
            OpenQueryDNS "update di2340 set dtr_update=1", objdbRs, True
            Script2File "update di2340 set dtr_update=1"
            
        End If
    End If
End Sub

Private Sub Command12_Click()
    Dim cEmpID As String
    Frame2.Enabled = False
    
    cEmpID = Text1.Text
    
    frmLookup.showPopup 3, " WHERE a.ACTIVE <> 0 "

    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "select * from di3670 where empid=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
        
            GetFields Me, objdbRs
            Text24.Text = IIf(objdbRs("REF_EMPID") <> "", objdbRs("REF_EMPID"), cResult)
            Text1.Text = cEmpID
            
            DTPicker4.Value = Now
            DTPicker5.Value = Now
            
            Text11.Text = ""
            ShowRecords
            Text7.Text = "0.00"
            Text8.Text = "0.00"
            Text20.Text = "0.00"
            
            XPSpin1.Value = 0
            XPSpin2.Value = 0
            
            Text18.Text = "0"
            Text19.Text = "0"
            
            Text13.Text = "0.00"
            Text14.Text = "0.00"
            Text15.Text = "0.00"
            
            If gCompanyID = "0002" Then Text29.Text = ""
            
            Combo2_Click
            Combo5.ListIndex = 0
            Combo5_Click
        End If
    End If
    Frame2.Enabled = True
End Sub

Private Sub Command13_Click()
    frmLookup.showPopup 20
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text30.Text = cResult
            Label50.Caption = objdbRs("DESCRIPTION")
        End If
    End If
    
    Text11.SetFocus

End Sub

Private Sub Command14_Click()
    frmLookup.showPopup 21
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "select a.WORKCENTERID, a.DESCRIPTION as WORK_DESC, ifnull(b.COSTCENTERID,'') as COSTCENTERID, ifnull(b.DESCRIPTION,'') as COST_DESC," & _
                       " ifnull(c.COMPCODE,'') as COMPCODE, ifnull(c.COMPName,'') as COMPName from pa97722 a " & _
                       " left join pa37722 b on a.costcenterid=b.costcenterid left join pa2660 c on a.compcode=c.compcode " & _
                       " where a.workcenterid=" & cQuote & cResult & cQuote, objdbRs, False
        
        If objdbRs.RecordCount > 0 Then
            Text31.Text = cResult
            Label53.Caption = objdbRs("WORK_DESC")
            Text30.Text = objdbRs("COSTCENTERID")
            Label50.Caption = objdbRs("COST_DESC")
            Text32.SetFocus
        End If
        
    End If
End Sub

Private Sub Command15_Click()
    frmLookup.showPopup 21
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "select a.WORKCENTERID, a.DESCRIPTION as WORK_DESC, ifnull(b.COSTCENTERID,'') as COSTCENTERID, ifnull(b.DESCRIPTION,'') as COST_DESC," & _
                       " ifnull(c.COMPCODE,'') as COMPCODE, ifnull(c.COMPName,'') as COMPName from pa97722 a " & _
                       " left join pa37722 b on a.costcenterid=b.costcenterid left join pa2660 c on a.compcode=c.compcode " & _
                       " where a.workcenterid=" & cQuote & cResult & cQuote, objdbRs, False
        
        If objdbRs.RecordCount > 0 Then
            Text32.Text = cResult
            Label58.Caption = objdbRs("WORK_DESC")
            Text30.SetFocus
        End If
        
    End If
End Sub

Private Sub Command2_Click()
frmLookup.showPopup 2, IIf(Trim(cParam) = "", "", " where lineid in " & cParam)
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "select a.lineid,a.linename,a.costcenterid, ifnull(b.description,'') as COST_DESC, ifnull(c.workcenterid,'') as workcenterid, ifnull(c.description,'') as WORK_DESC, a.ERPPOSCODE, b.cmpid from di5463 a left join pa37722 b on a.costcenterid = b.costcenterid left join pa97722 c on a.workcenterid = c.workcenterid WHERE lineid=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text5.Text = cResult
            Label14.Caption = objdbRs("LineName")
            Text31.Text = objdbRs("WORKCENTERID")
            Label53.Caption = objdbRs("WORK_DESC")
            Text32.Text = objdbRs("WORKCENTERID")
            Label58.Caption = objdbRs("WORK_DESC")
            Text30.Text = objdbRs("COSTCENTERID")
            Label50.Caption = objdbRs("COST_DESC")
            Select Case objdbRs("ERPPOSCODE")
                Case 0 ' "A"
                    Label57.Caption = "A"
                Case 1 ' "B"
                    Label57.Caption = "B"
                Case 2 ' "C"
                    Label57.Caption = "C"
                Case 3 ' "D"
                    Label57.Caption = "D"
                Case 4 ' "E"
                    Label57.Caption = "E"
                Case 5 ' "F"
                    Label57.Caption = "F"
                Case 6 ' "G"
                    Label57.Caption = "G"
                Case 7 ' "Z"
                    Label57.Caption = "Z"
            End Select
        End If
    End If
    
    Text5.SetFocus
End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 4
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM DI7670 WHERE posid=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text6.Text = cResult
            Label20.Caption = objdbRs("posName")
'            Text20.Text = objdbRs("ALLOWANCE")
        End If
    End If
    
    Text6.SetFocus
End Sub

Private Sub Command4_Click()
    frmLookup.showPopup 8
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text16.Text = cResult
        OpenQueryDNS "SELECT TAXCODE FROM PA8290 WHERE TAXID=" & cQuote & cResult & cQuote, objdbRs, False
        Label35.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TAXCODE"), "")
    End If
    Text16.SetFocus
End Sub

Private Sub Command5_Click()
    Frame2.Enabled = False
    
    frmLookup.showPopup 3, IIf(Trim(cParam) <> "", " WHERE a.DEPID IN " & cParam, "") & _
               IIf(Combo6.ListIndex = 2, "", IIf(Trim(cParam) <> "", " AND ", " WHERE ") & " a.ACTIVE IN (" & IIf(Combo6.ListIndex = 0, "0", "1,2") & ")")

    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "EMPID='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
        
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
    
    Frame2.Enabled = True
End Sub

Private Sub Command6_Click()
    frmLookup.showPopup 9
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text11.Text = cResult
        OpenQueryDNS "SELECT `DESCRIPTION`,CONCAT(TIME_FORMAT(TIME1,'%h:%i %p'),' - ',TIME_FORMAT(TIME2,'%h:%i %p')) AS `TIME` FROM PA74380 WHERE SHIFTID=" & cQuote & cResult & cQuote, objdbRs, False
        Label32.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
        Label33.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TIME"), "")
    End If
End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command6.Enabled = True
    Command12.Enabled = True
    Command14.Enabled = True
    Command15.Enabled = True
    Command13.Enabled = True
    
    Label14.Caption = ""
    Label57.Caption = ""
    Label20.Caption = ""

    Label53.Caption = ""
    Label54.Caption = ""
    Label58.Caption = ""



    If gCompanyID <> "0019" Then
        OpenQueryDNS "select * from pa2360 where userid =" & cQuote & gUserID & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            If objdbRs("GroupID") <> 1 Then
                ClearAll Me, True, True
                Command2.Enabled = True
                Command3.Enabled = True
                Command4.Enabled = True
                Command6.Enabled = True
               
                Text29.Enabled = False
                Text20.Enabled = False
                Text1.SetFocus
            End If
        Else
            ClearAll Me, True, True
            Text29.Enabled = False
            Text20.Enabled = False
            Text1.SetFocus
        End If
    Else
        Combo2.ListIndex = 1
        Combo2.Enabled = False
        
        Combo1.ListIndex = 0
        Combo1.Enabled = False
        
        Combo4.ListIndex = 1
        Combo4.Enabled = False

    End If
    
    ' --> assigned default shift...
    OpenQueryDNS "select * from pa74380 where (`default`=1) limit 1", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        Text11.Text = objdbRs("shiftid")
        Text11_KeyDown vbKeyReturn, 0
    Else
        Label32.Caption = ""
        Label33.Caption = ""
    End If
    
    Combo5.ListIndex = 0
    Combo5_Click
     
    Combo6.ListIndex = 0
    Combo6.Enabled = False
    If gCompanyID = "0019" Then
        Text7.Text = gBasicRate
        Text8.Text = gColaAmt
        
        Text7.Enabled = False
        Text8.Enabled = False
    End If
    Text26.Text = "Mariveles"
    
    If gCompanyID = "0019" Then
        DTPicker2.Enabled = False
        'DTPicker2.Value = DateAdd(2, 2, Now)
        DTPicker2.Value = DateAdd("yyyy", 2, DTPicker4.Value)
    End If
    
    MSHFlexGrid1.Enabled = False
    
    cSeries = GenerateSeries("EMPL")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("DI3670", "DI3670.EMPID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("EMPL")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    cEmpiId = Text1.Text
    Text1.SetFocus
End Sub

Private Sub Command8_Click()

    Log2Audit Name, "Edit Transation to EmpID #" & Text1.Text & " fullname " & Trim(Text3.Text) & ", " & Trim(Text2.Text) & " " & left(Trim(Text4.Text), 2) & ". "
    
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        cSeries = Text1.Text
        If gCompanyID = "0019" Then
'        ClearAll Me, True, False
            ClearAll Me, True, False
        End If
        CtrlPanel Me, nAdd
        
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command6.Enabled = True
        Command14.Enabled = True
        Command15.Enabled = True
        Command13.Enabled = True
        
        Combo6.Enabled = False
        MSHFlexGrid1.Enabled = False
        
        cEmpiId = Text1.Text
        
        If gCompanyID <> "0019" Then
            OpenQueryDNS "select * from pa2360 where userid =" & cQuote & gUserID & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                If objdbRs("GroupID") = 1 Then
                    ClearAll Me, False, False
                    Text29.Enabled = True
                    Text20.Enabled = True
                    Text29.SetFocus
                    
                    Command2.Enabled = False
                    Command3.Enabled = False
                    Command4.Enabled = False
                    Command6.Enabled = False
                    Command14.Enabled = False
                    Command15.Enabled = False
                    Command13.Enabled = False
             
                Else
                    ClearAll Me, True, False
                    Command2.Enabled = True
                    Command3.Enabled = True
                    Command4.Enabled = True
                    Command6.Enabled = True
                    Command14.Enabled = True
                    Command15.Enabled = True
                    Command13.Enabled = True
                    
                    Text29.Enabled = False
                    Text20.Enabled = False
                    Text2.SetFocus
                End If
            Else
                ClearAll Me, True, False
                Text29.Enabled = False
                Text20.Enabled = False
                Text2.SetFocus
            End If
        Else
            Combo2.ListIndex = 1
            Combo2.Enabled = False
            
            Combo1.ListIndex = 0
            Combo1.Enabled = False
            
            Combo4.ListIndex = 1
            Combo4.Enabled = False
            
            Text7.Enabled = False
            Text8.Enabled = False
            
            DTPicker2.Enabled = False
            'DTPicker2.Value = "2012-09-15"
            DTPicker2.Value = DateAdd("yyyy", 2, DTPicker4.Value)
        End If
    End If
End Sub

Sub Emp_BTTN(nMode As Boolean)

End Sub

Private Sub Command9_Click()
    On Error GoTo ErrAdminDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Employee Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM DI3670 WHERE EMPID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM DI3670 WHERE EMPID=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text3.Text)))
        
'        OpenQueryDNS "DELETE FROM DI36770 WHERE EMPID=" & cQuote & Text1.Text & cQuote, oTempADO, True
'        Script2File "DELETE FROM DI36770 WHERE EMPID=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        
'        OpenQueryDNS "SELECT * FROM DI3670 ORDER BY EMPID", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
    
    Exit Sub
    
ErrAdminDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim cSqlStmt As String
    Log2Audit Name, "OPEN"
    
'    myArray = Array("TXT:[Ded ID]:3:False", _
'                    "TXT:[DedName]:30:True", _
'                    "NUM:[Amount]:12:True", _
'                    "NUM:[Accrued Amt]:14:True", _
'                    "NUM:[Cut Off Amt]:14:True", _
'                    "NUM:[Period 1]:8:True", _
'                    "NUM:[Period 2]:8:True", _
'                    "NUM:[Auto-Compute]:1:False", _
'                    "NUM:[Fix]:1:False", _
'                    "TXT:[CMPID]:5:False")

    myArray = Array("TXT:[Dep ID]:3:False", _
                    "TXT:[Department]:25:True")

    Tag = nAccess_Tag
    nAdd = 0
    
'    XPFrame2.Visible = gUserGroup > 0
    
    OpenQueryDNS "SELECT LINEID, LINENAME FROM DI5463 ORDER BY LINEID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
    
    Combo6.ListIndex = 0
    Combo6_Click
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

Private Sub MSHFlexGrid1_EnterCell()
    Dim nCtr As Integer
    If nAdd <> 0 Then Exit Sub
    
    cParam = ""
    
    If Check1.Value <> vbChecked Then
        cParam = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)
    End If
    
    If Trim(cParam) <> "" Then
        cParam = "(" & cParam & ")"
    End If
    
    Combo6_Click
End Sub

Private Sub MSHFlexGrid1_GotFocus()
    MSHFlexGrid1_EnterCell
End Sub


Private Sub Text1_Validate(Cancel As Boolean)
    Dim cSqlStmt As String
    
    If nAdd <> 0 Then
        cSqlStmt = "select * from di3670 Where empid = " & cQuote & Text1.Text & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            cSqlStmt = "Warning!!!" & vbCrLf & _
                       "Employee ID already Exist!!!"
            MsgBox cSqlStmt, vbCritical, "System Advisory!!!"
            If Text1.Text = cEmpiId Then
                Cancel = False
            Else
                Cancel = True
                Text1.Text = cEmpiId
            End If
        End If
    End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text11.Text) = "" Then
            Command6_Click
        Else
            OpenQueryDNS "SELECT `DESCRIPTION`,CONCAT(TIME_FORMAT(TIME1,'%h:%i %p'),' - ',TIME_FORMAT(TIME2,'%h:%i %p')) AS `TIME` FROM PA74380 WHERE SHIFTID=" & cQuote & Text11.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label32.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
                Label33.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TIME"), "")
            Else
                Label32.Caption = ""
                Label33.Caption = ""
                MsgBox "Shift ID not found!", vbCritical, App.Title
                Text11.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text16.Text) = "" Then
            Command4_Click
        Else
            OpenQueryDNS "SELECT TAXCODE FROM PA8290 WHERE TAXID=" & cQuote & Text16.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label35.Caption = objdbRs("TAXCODE")
            Else
                Label35.Caption = ""
                MsgBox "Tax Code not found!", vbCritical, App.Title
                Text16.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
    Dim cSqlStmt As String
    If (nAdd = 0) Or (Trim(Text17.Text) = "") Then Exit Sub
    cSqlStmt = "select * from di3670 where (tcid=" & cQuote & Text17.Text & cQuote & ")" & _
               " and (empid<>" & cQuote & Text1.Text & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cSqlStmt = "Warning!!!" & vbCrLf & _
                   "Time Card ID already belongs to Employee #" & objdbRs("empid")
        MsgBox cSqlStmt, vbCritical, "System Advisory!!!"
        Cancel = True
    End If
End Sub

Private Sub Text30_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text30.Text) = "" Then
            Command13_Click
        Else
            OpenQueryDNS "SELECT * FROM PA37722 WHERE COSTCENTERID=" & cQuote & Text30.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label50.Caption = objdbRs("DESCRIPTION")
                Text11.SetFocus
            Else
                Label50.Caption = ""
                MsgBox "Cost Center ID Not Found!", vbCritical, App.Title
                Text30.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Text31_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text31.Text) = "" Then
            Command14_Click
        Else
            
            OpenQueryDNS "select a.WORKCENTERID, a.DESCRIPTION as WORK_DESC, ifnull(b.COSTCENTERID,'') as COSTCENTERID, ifnull(b.DESCRIPTION,'') as COST_DESC," & _
                           " ifnull(c.COMPCODE,'') as COMPCODE, ifnull(c.COMPName,'') as COMPName from pa97722 a " & _
                           " left join pa37722 b on a.costcenterid=b.costcenterid left join pa2660 c on a.compcode=c.compcode " & _
                           " where a.workcenterid=" & cQuote & Text31.Text & cQuote, objdbRs, False
            
            If objdbRs.RecordCount > 0 Then
                Text31.Text = objdbRs("WORKCENTERID")
                Label53.Caption = objdbRs("WORK_DESC")
                Text30.Text = objdbRs("COSTCENTERID")
                Label50.Caption = objdbRs("COST_DESC")

                Text32.SetFocus
            Else
                Label53.Caption = ""
                Text30.Text = ""
                Label50.Caption = ""
                
                MsgBox "Work Center ID Not Found!", vbCritical, App.Title
                
                Text31.SetFocus
            End If
            
            
        End If
    End If
End Sub

Private Sub Text32_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text32.Text) = "" Then
            Command15_Click
        Else
            
            OpenQueryDNS " select a.WORKCENTERID, a.DESCRIPTION as WORK_DESC from pa97722 a left join pa2660 c on a.compcode=c.compcode " & _
                           " where a.workcenterid=" & cQuote & Text32.Text & cQuote, objdbRs, False
            
            
            If objdbRs.RecordCount > 0 Then
                Text32.Text = objdbRs("workcenterid")
                Label58.Caption = objdbRs("WORK_DESC")
                Text30.SetFocus
            Else
                Text32.Text = ""
                Label58.Caption = ""
                
                MsgBox "BEP Work Center ID Not Found!", vbCritical, App.Title
                
                Text30.SetFocus
            End If
            
            
        End If
    End If

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text5.Text) = "" Then
            Command2_Click
        Else
            OpenQueryDNS "SELECT * FROM DI5463 WHERE lineid=" & cQuote & Text5.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label14.Caption = objdbRs("LineName")
                Select Case objdbRs("ERPPOSCODE")
                    Case 0 ' "A"
                        Label57.Caption = "A"
                    Case 1 ' "B"
                        Label57.Caption = "B"
                    Case 2 ' "C"
                        Label57.Caption = "C"
                    Case 3 ' "D"
                        Label57.Caption = "D"
                    Case 4 ' "E"
                        Label57.Caption = "E"
                    Case 5 ' "F"
                        Label57.Caption = "F"
                    Case 6 ' "G"
                        Label57.Caption = "G"
                    Case 7 ' "Z"
                        Label57.Caption = "Z"
                End Select
            Else
                Label14.Caption = ""
                Label57.Caption = ""
                MsgBox "Department ID not found!", vbCritical, App.Title
                Text5.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text6.Text) = "" Then
            Command3_Click
        Else
            OpenQueryDNS "SELECT * FROM DI7670 WHERE posid=" & cQuote & Text6.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
'                Text20.Text = objdbRs("ALLOWANCE")
                Label20.Caption = objdbRs("posName")
            Else
                Label20.Caption = ""
                MsgBox "Position ID not found!", vbCritical, App.Title
                Text6.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd = 0 Then Exit Sub
    
    ' --> apply WAP or Contractual rate...
    If (Shift And vbCtrlMask) > 0 Then
        If (KeyCode = vbKeySpace) And (Combo2.ListIndex = 1) Then
            If MsgBox("Apply " & IIf(Val(Text21.Text) <> 1, "WAP", "Contractual") & " rate?", vbYesNo, "Confirm Employment Status...") = vbYes Then
                Text7.Text = gBasicRate * IIf(Val(Text21.Text) <> 1, 0.75, 1)
                Text8.Text = gColaAmt * IIf(Val(Text21.Text) <> 1, 0.75, 1)
                Text21.Text = IIf(Val(Text21.Text) <> 1, 1, 0)
                Combo2_Click
            End If
        End If
    End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
    If Not IsNumeric(Text7.Text) Then
        Text7.Text = "0.0000"
        Cancel = False
    End If
End Sub


Private Sub XPButton1_Click()
'    GetUserRights PadStr(frmMain.mnuEmployee.Name, " ", 100, PadRight), gUserID
'    frmEmpDed.Text1 = Text1.Text
'    frmEmpDed.Text2 = Trim(Text2.Text) & " " & IIf(Trim(Text4.Text) = "", "", Trim(Text4.Text) & " ") & Trim(Text3.Text)
'    frmEmpDed.ShowRecords
'    frmEmpDed.Show
End Sub

Private Sub XPButton4_Click()
    With frmEmpShift
        .Text1.Text = Text1.Text
        .Text2(0).Text = Trim(Text2.Text) & " " & Trim(Text3.Text)
        .Text2(2).Text = Label14.Caption
        .SetFilter Text1.Text, Text11.Text
        .Show
    End With
End Sub
