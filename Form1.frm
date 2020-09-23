VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory DC Test"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
      Height          =   315
      Left            =   3060
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh Form"
      Height          =   315
      Left            =   3060
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Form"
      Height          =   315
      Left            =   3060
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdfreeexit 
      Caption         =   "Free Memory and exit"
      Height          =   615
      Left            =   3060
      TabIndex        =   6
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdcopyd 
      Caption         =   "Copy DC D"
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdcopyc 
      Caption         =   "Copy DC C"
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdcopyb 
      Caption         =   "Copy DC B"
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdcopya 
      Caption         =   "Copy DC A"
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load pictures into memory"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   1455
   End
   Begin VB.PictureBox picsource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1020
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblme 
      Caption         =   "By Nick Thompson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   1500
      X2              =   1500
      Y1              =   1500
      Y2              =   -180
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   1500
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label lblautoredraw 
      Caption         =   "Autoredraw = False"
      Height          =   195
      Left            =   3060
      TabIndex        =   9
      Top             =   1500
      Width           =   1455
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim abmp As BITMAP
'Dim bbmp As BITMAP
'Dim cbmp As BITMAP
'Dim dbmp As BITMAP
'Dim adc As Long
'Dim bdc As Long
'Dim cdc As Long
'Dim ddc As Long
'Dim ahbmp As Long
'Dim bhbmp As Long
'Dim chbmp As Long
'Dim dhbmp As Long
'Dim aprevbmp As Integer
'Dim bprevbmp As Long
'Dim cprevbmp As Long
'Dim dprevbmp As Long
  Dim aDC As Long  'source bitmap (color)
  Dim abmp As BITMAP 'description of the source bitmap
  Dim aPrevBmp As Long  'Holds previous bitmap in source DC
  Dim bDC As Long  'source bitmap (color)
  Dim bbmp As BITMAP 'description of the source bitmap
  Dim bPrevBmp As Long  'Holds previous bitmap in source DC
  Dim cDC As Long  'source bitmap (color)
  Dim cbmp As BITMAP 'description of the source bitmap
  Dim cPrevBmp As Long  'Holds previous bitmap in source DC
  Dim dDC As Long  'source bitmap (color)
  Dim dbmp As BITMAP 'description of the source bitmap
  Dim dPrevBmp As Long  'Holds previous bitmap in source DC
  Dim hPrevBmp As Integer 'Bitmap holds previous bitmap selected in DC
  Dim Success As Long 'Stores result of call to Windows API







Private Sub cmdchange_Click()
If frmmain.AutoRedraw = True Then
  frmmain.AutoRedraw = False
  lblautoredraw.Caption = "Autoredraw = False"
Else
  frmmain.AutoRedraw = True
  lblautoredraw.Caption = "Autoredraw = True"
End If
End Sub

Private Sub cmdclear_Click()
frmmain.Cls
End Sub

Private Sub cmdcopya_Click()
  Success = BitBlt(frmmain.hdc, 0, 0, 50, 50, aDC, 0, 0, SRCCOPY)
End Sub

Private Sub cmdcopyb_Click()
  Success = BitBlt(frmmain.hdc, 50, 0, 50, 50, bDC, 0, 0, SRCCOPY)
End Sub

Private Sub cmdcopyc_Click()
  Success = BitBlt(frmmain.hdc, 0, 50, 50, 50, cDC, 0, 0, SRCCOPY)
End Sub

Private Sub cmdcopyd_Click()
  Success = BitBlt(frmmain.hdc, 50, 50, 50, 50, dDC, 0, 0, SRCCOPY)
End Sub

Private Sub cmdfreeexit_Click()
  hPrevBmp = SelectObject(aDC, aPrevBmp) 'Select orig object
  Success = DeleteDC(aDC)          'Deallocate system resources.
  hPrevBmp = SelectObject(bDC, bPrevBmp) 'Select orig object
  Success = DeleteDC(bDC)          'Deallocate system resources.
End
End Sub

Private Sub cmdload_Click()
  picsource.Picture = LoadPicture(App.Path & "\1.bmp")
  Success = newGetObject(picsource, Len(abmp), abmp)
  aDC = CreateCompatibleDC(picsource.hdc)    'Create DC to hold stage
  aPrevBmp = SelectObject(aDC, picsource)     'Select bitmap in DC

  picsource.Picture = LoadPicture(App.Path & "\2.bmp")
  Success = newGetObject(picsource, Len(bbmp), bbmp)
  bDC = CreateCompatibleDC(picsource.hdc)    'Create DC to hold stage
  bPrevBmp = SelectObject(bDC, picsource)     'Select bitmap in DC

  picsource.Picture = LoadPicture(App.Path & "\3.bmp")
  Success = newGetObject(picsource, Len(cbmp), cbmp)
  cDC = CreateCompatibleDC(picsource.hdc)    'Create DC to hold stage
  cPrevBmp = SelectObject(cDC, picsource)     'Select bitmap in DC

  picsource.Picture = LoadPicture(App.Path & "\4.bmp")
  Success = newGetObject(picsource, Len(dbmp), dbmp)
  dDC = CreateCompatibleDC(picsource.hdc)   'Create DC to hold stage
  dPrevBmp = SelectObject(dDC, picsource)     'Select bitmap in DC

End Sub


Private Sub cmdrefresh_Click()
frmmain.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
  hPrevBmp = SelectObject(aDC, aPrevBmp) 'Select orig object
  Success = DeleteDC(aDC)          'Deallocate system resources.
  hPrevBmp = SelectObject(bDC, bPrevBmp) 'Select orig object
  Success = DeleteDC(bDC)          'Deallocate system resources.
End

End Sub
