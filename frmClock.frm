VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED Dot Matrix Clock"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   945
   End
   Begin DotMatrix_control.ucDotMatrix ucDotMatrix1 
      Height          =   1575
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   2778
      Size            =   8.25
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Switch As Integer

Private Sub Form_Load()

    ucDotMatrix1.Text = "My Clock :-)"
    ucDotMatrix1.Left = 0
    ucDotMatrix1.Top = 0
    Me.Height = (ucDotMatrix1.Height * Screen.TwipsPerPixelY) + (Me.Height - (Me.ScaleHeight * Screen.TwipsPerPixelY))
    Me.Width = (ucDotMatrix1.Width * Screen.TwipsPerPixelX) + (Me.Width - (Me.ScaleWidth * Screen.TwipsPerPixelX))

End Sub

Private Sub Timer1_Timer()

    If Switch < 4 Then
        ucDotMatrix1.Text = Time
    Else
        ucDotMatrix1.Text = Format(Date, "mmm d, yyyy")
    End If
    Switch = Switch + 1
    If Switch > 6 Then Switch = 0

End Sub
