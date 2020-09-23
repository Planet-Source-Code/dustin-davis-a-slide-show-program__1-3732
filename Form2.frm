VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3840
      Top             =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'Slide Show v1.0
'Author: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'Slide Show has alot of things it can do, Add as many images as u want,
'add Bitmaps, Jpeg Files and GIF files, mix em up even
'You can add ALL of the image file listed with one click of a button
'One click remove item from list
'empty list with one click
'Auto start show on load
'Pause the show
'timer settings and so on.
'
'Please do not steal this code! It was not coded by you. If you use it,
'please give me proper credit for it!
'ENJOY!!
'This is the form for the show!
'*********************************************************************************

Dim TotalImages As Long
Dim temp As String
Dim CurrentImg As Long
Public PauseShow As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
'exit the show
Unload Me
End Sub

Private Sub Form_Load()
Form1.Visible = False
PauseShow = False
Label1.Top = 0
Label1.Left = 0
'see if showfile is enabled
If Form1.Check1.Value = Unchecked Then
    Label1.Visible = False
End If
'set timer interval
With Timer1
    .Enabled = False
    .Interval = GetSetting("SlideShow", "settings", "delay", "5")
    .Interval = Timer1.Interval * 1000
    .Enabled = True
End With
'get some settings
TotalImages = GetSetting("SlideShow", "pictures", "HowMany", "0")
CurrentImg = 1
'set positions
With Form2
    .Top = 0
    .Left = 0
End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'on mouse click, see if the show should exit, or pause or resume!
If Button = 1 Then 'exit
    Unload Me
ElseIf Button = 2 Then 'pause / resume
    If PauseShow = False Then 'pause
        Label1.Caption = "Paused on " & temp
        PauseShow = True
        Exit Sub
    Else 'resume
        Label1.Caption = "Resumeing on " & temp
        PauseShow = False
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'show form1 again
Form1.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'on mouse click, see if the show should exit, or pause or resume!
If Button = 1 Then 'exit
    Unload Me
ElseIf Button = 2 Then 'pause / resume
    If PauseShow = False Then 'pause
        Label1.Caption = "Paused on " & temp
        PauseShow = True
        Exit Sub
    Else 'resume
        Label1.Caption = "Resumeing on " & temp
        PauseShow = False
        Exit Sub
    End If
End If
End Sub

Private Sub Timer1_Timer()
'this shows all the pictures, etc.
Dim ImgWidth As Long
Dim ImgHalfWidth As Long
Dim ImgHeight As Long
Dim ImgHalfHeight As Long
'see if the show is paused or not
If PauseShow = True Then
        Exit Sub
    End If
Image1.Visible = False 'turns it off so when it resizes it wont flicker
'get image location
    temp = GetSetting("SlideShow", "pictures", CurrentImg, "0")
    'load picture
    Image1.Picture = LoadPicture(temp)
    'show image location
    Label1.Caption = temp
    'start centering proccess
    ImgWidth = Image1.Width
    ImgHalfWidth = ImgWidth / 2
    ImgHeight = Image1.Height
    ImgHalfHeight = ImgHeight / 2
    
    Image1.Top = (Form2.Height / 2) - ImgHalfHeight
    Image1.Left = (Form2.Width / 2) - ImgHalfWidth
    'end centering proccess - now the image is centered perfectly
    Image1.Visible = True 'show the picture now
    CurrentImg = CurrentImg + 1
    'check if it should loop now
    If CurrentImg >= TotalImages Then
        CurrentImg = 1
    Else
        DoEvents
    End If
    
    DoEvents
End Sub
