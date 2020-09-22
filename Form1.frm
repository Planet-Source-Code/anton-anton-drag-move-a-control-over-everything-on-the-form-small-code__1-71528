VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1680
      ScaleHeight     =   1305
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   600
      Width           =   2055
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         FillColor       =   &H80000007&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   1995
         TabIndex        =   7
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drag This Window!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   3720
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   1455
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   4200
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   1800
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Made By Anton Khakhalin

Dim xX, Yy, xXx, yYy, PrEsS As Integer

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xXx = xX
yYy = Yy
PrEsS = 1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PrEsS = 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xX = X
Yy = Y
If PrEsS = 1 Then
Picture3.Left = Picture3.Left + xX - xXx
Picture3.Top = Picture3.Top + Yy - yYy
End If
End Sub
