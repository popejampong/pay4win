VERSION 5.00
Begin VB.Form frmRefPoint 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRefPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Point
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' Change region of a window:
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


' Precanned region creation functions:
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
' Polygon type:
Private Const WINDING = 2

' Region combination:
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

    ' Region combination types:
    Private Const RGN_AND = 1
    Private Const RGN_COPY = 5
    Private Const RGN_DIFF = 4
    Private Const RGN_MAX = RGN_COPY
    Private Const RGN_MIN = RGN_AND
    Private Const RGN_OR = 2
    Private Const RGN_XOR = 3
    ' Region combination return values:
    Private Const COMPLEXREGION = 3
    Private Const SIMPLEREGION = 2
    Private Const NULLREGION = 1

' GDI Clear up:
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private m_bHorizontal As Boolean
Private m_bVertical As Boolean

Public Sub Init(ByVal bHorizontal As Boolean, ByVal bVertical As Boolean)
   m_bHorizontal = bHorizontal
   m_bVertical = bVertical
End Sub

Private Sub Form_Load()
Dim hRgn As Long
Dim hRgnCircle As Long
Dim hRgnTri As Long
Dim tP(0 To 3) As POINTAPI
   
   hRgn = CreateRectRgn(0, 0, 0, 0)
   
   hRgnCircle = CreateEllipticRgn(10, 10, 15, 15)
   CombineRgn hRgn, hRgnCircle, hRgn, RGN_OR
   DeleteObject hRgnCircle
   
   If (m_bVertical) Then
      ' Up
      tP(0).X = 11
      tP(0).Y = 1
      tP(1).X = 6
      tP(1).Y = 6
      tP(2).X = 18
      tP(2).Y = 6
      tP(3).X = 12
      tP(3).Y = 0
      hRgnTri = CreatePolygonRgn(tP(0), 4, WINDING)
      CombineRgn hRgn, hRgn, hRgnTri, RGN_OR
      DeleteObject hRgnTri
      
      ' Down
      tP(0).X = 12
      tP(0).Y = 23
      tP(1).X = 7
      tP(1).Y = 18
      tP(2).X = 17
      tP(2).Y = 18
      tP(3).X = 12
      tP(3).Y = 23
      hRgnTri = CreatePolygonRgn(tP(0), 4, WINDING)
      CombineRgn hRgn, hRgn, hRgnTri, RGN_OR
      DeleteObject hRgnTri
   End If
   
   If (m_bHorizontal) Then
      ' Left
      tP(0).X = 1
      tP(0).Y = 11
      tP(1).X = 6
      tP(1).Y = 6
      tP(2).X = 6
      tP(2).Y = 17
      tP(3).X = 1
      tP(3).Y = 12
      hRgnTri = CreatePolygonRgn(tP(0), 4, WINDING)
      CombineRgn hRgn, hRgn, hRgnTri, RGN_OR
      DeleteObject hRgnTri
      
      ' Right
      tP(0).X = 23
      tP(0).Y = 11
      tP(1).X = 18
      tP(1).Y = 6
      tP(2).X = 18
      tP(2).Y = 17
      tP(3).X = 23
      tP(3).Y = 12
      hRgnTri = CreatePolygonRgn(tP(0), 4, WINDING)
      CombineRgn hRgn, hRgn, hRgnTri, RGN_OR
      DeleteObject hRgnTri
   End If
   
   SetWindowRgn Me.hwnd, hRgn, True
   
End Sub


