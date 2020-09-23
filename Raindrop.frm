VERSION 5.00
Begin VB.Form Raindrop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Raindrops  by Scythe       Press ESC to QUIT"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8970
   ControlBox      =   0   'False
   ForeColor       =   &H000100FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton Option3 
      Caption         =   "Let it Rain"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   6840
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Move Over Picture"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Set Drops"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000100FF&
      Height          =   6750
      Left            =   0
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9000
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000100FF&
      Height          =   6750
      Left            =   0
      Picture         =   "Raindrop.frx":0000
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   3120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Move over the Picture to disturb the water"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Klick on the Picture to set a Raindrop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   6840
      Width           =   4815
   End
End
Attribute VB_Name = "Raindrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Raindrops Demo
' Compile for real speed

'Not all inhere is from me
'I found an old pascal source somewhere
'in the net and converted it to VB


'Version 2
'now runs smooth on my old 333Mhz with Matrox MGA G100

'Changed the Circle routine for more speed
'From:  For h = 0 To 360
'To:     For h = 0 To 360 Step Drop(i).Size
'Gives about 50% of speed

'Changeg The get Picture to draw on
'From GetDibBits to CopyMemory
'Speeds up the thing a second time


Option Explicit

'Needed for speed optimation
'Private Declare Function GetTickCount Lib "kernel32" () As Long

'To copy our pic real fast
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Use DIB for fast GFX
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Type Ring
 x As Long
 y As Long
 Size As Integer
 Radius As Integer
End Type

Private Type SinCos
 Cos As Double
 Sin As Double
End Type

Private Type PointApi
 x As Long
 y As Long
End Type

Dim LookUp(360) As SinCos     'Table for faster calculations

Dim Drop()    As Ring         'Hold the snow
Dim Max       As Integer      'Maximal Size for an Drop
Dim SetDrops  As Boolean      'Set Drops or disturb water
Dim PicNew()  As RGBQUAD      'Hold our New Picture
Dim PicOrg()  As RGBQUAD      'Hold our Original Picute
Dim Binfo     As BITMAPINFO   'The GetDIBits API needs some Infos
Dim OrgLng    As Long         'Holds the Picsize for Copy Memory



Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Timer1.Enabled = False
  Unload Me
  End
 End If
End Sub

Private Sub Form_Load()
 Dim i As Integer
 Dim f As Integer

 'Fill our Sinus & Cosinus Table
 For i = 0 To 360
  LookUp(i).Cos = Cos(i * 3.14159265358979 / 180)
  LookUp(i).Sin = Sin(i * 3.14159265358979 / 180)
 Next i

 'Create a buffer that holds our picture
 ReDim PicNew(0 To Pic.ScaleWidth - 1, 0 To Pic.ScaleHeight - 1)
 ReDim PicOrg(0 To Pic.ScaleWidth - 1, 0 To Pic.ScaleHeight - 1)

 'Get the Picturesize in Memory for CopyMemory
 'X*Y*4 (4 for the 4 Bytes of RGBQUAD)
 OrgLng = (UBound(PicOrg, 1) + 1) * (UBound(PicOrg, 2) + 1) * 4

 'Set the infos for our apicall
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = Pic.ScaleWidth
 .biHeight = Pic.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = Pic.ScaleWidth * Pic.ScaleHeight
 End With

 'If we start in ide show a message
 If InIde = True Then
  PicBack.CurrentX = 100
  PicBack.CurrentY = 50
  PicBack.Print "Please compile to get full SPEED"
 End If

 'Now get the Original Picture
 GetDIBits PicBack.hdc, PicBack.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS

 ShowPic

 'Set Starting Parameters
 Option1_Click

End Sub


Private Sub Option1_Click()
 If Option1.Value = True Then
  SetDrops = True
  Label1.Visible = True
  Label2.Visible = False
  Max = 200
  ReDim Drop(0)
  ShowPic
 End If
End Sub

Private Sub Option2_Click()
 If Option2.Value = True Then
  SetDrops = False
  Label1.Visible = False
  Label2.Visible = True
  Max = 50
  ReDim Drop(0)
  ShowPic
 End If
End Sub
Private Sub Option3_Click()
 If Option3.Value = True Then
  SetDrops = True
  Label1.Visible = False
  Label2.Visible = False
  Max = 50
  ReDim Drop(0)
  ShowPic
 End If
End Sub
Private Sub DrawDrop()
 Dim i As Long
 Dim f As Integer
 Dim g As Integer
 Dim h As Integer
 Dim r As Integer
 Dim x As Long
 Dim y As Long
On Error Resume Next
'Get the picture to paint on
CopyMemory PicNew(0, 0), PicOrg(0, 0), OrgLng

'Move to all our drops
For i = 1 To UBound(Drop)
 f = PicNew(100, 100).rgbBlue
 'how thik is our drop
 r = Drop(i).Size + CInt(Rnd * 2)
 'Calculate and draw the drop using a precalculated table for sinus and cosinus
 For h = 0 To 360 Step Drop(i).Size
  x = Drop(i).Radius * LookUp(h).Cos + Drop(i).x
  y = Drop(i).Radius * LookUp(h).Sin + Drop(i).y

  'Set the pixels with a offset of 3 to get the waterhight effect
  'u could make it better by calculating a round wave but
  'this routine is faster and looks good to
  For f = 0 To r
   For g = 0 To r
    PicNew(x + f, y + g) = PicOrg(x + 3 + f, y + 3 + g)
   Next g
  Next f
 Next h
 DoEvents
 'Increase the radius for next time
 Drop(i).Radius = Drop(i).Radius + 3
 'to get the bigger point if we newly hit the water we
 'need a bigger size that gets fast smaller
 If Drop(i).Size > 1 Then Drop(i).Size = Drop(i).Size - 1
Next i

'Is there a raindrop to big
'if yes remove him
For i = 1 To UBound(Drop)
 If Drop(i).Radius > Max Then
  x = 1
  If UBound(Drop) > 1 Then
   For f = i + 1 To UBound(Drop)
    Drop(x) = Drop(f)
    x = x + 1
   Next f
   ReDim Preserve Drop(f - 2)
  Else
   'No Drops
   ReDim Drop(0)
  End If
  If UBound(Drop) = 0 Then
   ShowPic
   Exit Sub
  End If
  Exit For
 End If
Next i

'Show our Raindrops
SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicNew(0, 0), Binfo, DIB_RGB_COLORS
Pic.Refresh
DoEvents
End Sub



Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If SetDrops = False Then Exit Sub
 Dim i As Integer
 'Increase the number of drops
 i = UBound(Drop) + 1
 ReDim Preserve Drop(i)
 'Set new drop
 Drop(i).x = x
 Drop(i).y = Pic.Height - y
 Drop(i).Size = 8
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If SetDrops = True Then Exit Sub
 Dim i As Integer
 Dim f As Integer
 i = UBound(Drop)
 'We can place all 5 Pixels
 If Abs(x - Drop(i).x) > 5 Then
  If Abs(Pic.Height - Drop(i).y - y) > 5 Then
   i = i + 1
   ReDim Preserve Drop(i)
   Drop(i).x = x
   Drop(i).y = Pic.Height - y
   Drop(i).Size = 6

   'Oly 60 drops allowed
   If i > 60 Then
    For f = 1 To 60
     Drop(f) = Drop(i - 60 + f)
    Next f
    ReDim Preserve Drop(60)
   End If

  End If
 End If
End Sub

'Call our Paint sub
Private Sub Timer1_Timer()
 Dim i As Integer
 If Option3.Value = True And UBound(Drop) < 50 Then
  If Drop(UBound(Drop)).Radius > 1 Or UBound(Drop) = 0 Then
   i = UBound(Drop) + 1
   ReDim Preserve Drop(i)
   Drop(i).x = Pic.Width * Rnd
   Drop(i).y = Pic.Height * Rnd
   Drop(i).Size = 8
  End If
 End If
 If UBound(Drop) > 0 Then
  DrawDrop
 End If
 DoEvents
End Sub

'Test if we are in ide or compiled mode
Private Function InIde() As Boolean
 On Error GoTo DivideError
 Debug.Print 1 / 0
 Exit Function
DivideError:
 InIde = True
End Function

'Show the Startpicture
Private Sub ShowPic()
 'Clear Picture
 SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS
 Pic.Refresh
End Sub

