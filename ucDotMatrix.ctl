VERSION 5.00
Begin VB.UserControl ucDotMatrix 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   264
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   90
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Shape shpDot 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   0
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   45
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "ucDotMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ucDotMatrix coded by Frédéric Côté
'idea from Michael Hammond@PSC (Dot Matrix clock)

Option Explicit

Public Enum MatrixColor
    Red = 0
    Blue = 1
    Green = 2
    Yellow = 3
End Enum

Private mstrText As String
Private mblnEnabled As Boolean
Private mudtColor As MatrixColor
Private mblnBuilding As Boolean

Private Sub BuildMatrix()

Dim x As Integer, y As Integer
Dim lngFillLit As Long, lngFillNotLit As Long

    Select Case mudtColor
    Case 0
        lngFillLit = RGB(255, 0, 0)
    Case 1
        lngFillLit = RGB(0, 0, 255)
    Case 2
        lngFillLit = RGB(0, 190, 0)
    Case 3
        lngFillLit = RGB(255, 255, 0)
    End Select
    Select Case mudtColor
    Case 0
        lngFillNotLit = RGB(96, 0, 0)
    Case 1
        lngFillNotLit = RGB(0, 0, 96)
    Case 2
        lngFillNotLit = RGB(0, 96, 0)
    Case 3
        lngFillNotLit = RGB(96, 96, 0)
    End Select
    With picText
        .Height = .TextHeight("Xj")
        If mblnEnabled Then
            .CurrentX = (.Width - .TextWidth(mstrText)) / 2
            picText.Print mstrText
        End If
        UserControl.Cls
        UserControl.DrawWidth = 1
        UserControl.ForeColor = RGB(128, 0, 0)
        UserControl.FillStyle = vbFSSolid
        mblnBuilding = True 'modifications of the sizes here would cause other calls of this sub
        UserControl.Width = ((.Width + 1) * 7 * Screen.TwipsPerPixelX) 'width of picture(+1 more column)
        UserControl.Height = ((.Height + 1) * 7 * Screen.TwipsPerPixelY) 'height of picture(+1 more line)
        mblnBuilding = False
        For x = 0 To .Width '1 more column than the actual picture
            For y = 0 To .Height '1 more line than the actual picture
                If .Point(x, y) = 0 Then
                    UserControl.FillColor = lngFillLit
                Else
                    UserControl.FillColor = lngFillNotLit
                End If
                'Addition is supposed to be faster than multiplication
                UserControl.Circle (x + x + x + x + x + x + x + 3, y + y + y + y + y + y + y + 3), 2
            Next y
        Next x
        .Cls
    End With

End Sub

Public Property Get Color() As MatrixColor

    Color = mudtColor

End Property

Public Property Let Color(ByVal uNewValue As MatrixColor)
Attribute Color.VB_ProcData.VB_Invoke_PropertyPut = ";Liste"
Attribute Color.VB_UserMemId = -513

    mudtColor = uNewValue
    PropertyChanged "Color"
    BuildMatrix

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "is the control enabled?"
Attribute Enabled.VB_UserMemId = -514

    Enabled = mblnEnabled

End Property

Public Property Let Enabled(ByVal bNewValue As Boolean)

    mblnEnabled = bNewValue
    PropertyChanged "Enabled"
    BuildMatrix

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Police"
Attribute Font.VB_UserMemId = -512

    Set Font = picText.Font

End Property

Public Property Set Font(ByVal fNewValue As StdFont)

    With picText
        Set .Font = fNewValue
        .Font.Bold = fNewValue.Bold
        .Font.Size = fNewValue.Size
        .Font.Italic = fNewValue.Italic
        .Font.Underline = fNewValue.Underline
        .Font.Strikethrough = fNewValue.Strikethrough
    End With
    PropertyChanged "Font"
    PropertyChanged "Bold"
    PropertyChanged "Size"
    PropertyChanged "Italic"
    PropertyChanged "Underline"
    PropertyChanged "Strike"
    BuildMatrix

End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Change the text to display on the LED"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Texte"
Attribute Text.VB_UserMemId = -517

    Text = mstrText

End Property

Public Property Let Text(ByVal sNewValue As String)

    mstrText = sNewValue
    PropertyChanged "Text"
    BuildMatrix

End Property

Private Sub UserControl_Initialize()

    mblnBuilding = False

End Sub

Private Sub UserControl_InitProperties()

    mstrText = "DotMatrix"
    mblnEnabled = True
    mudtColor = Red

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mstrText = PropBag.ReadProperty("Text", "DotMatrix")
    With picText
        .Font = PropBag.ReadProperty("Font", .Font)
        .Font.Size = PropBag.ReadProperty("Size", 8)
        .Font.Bold = PropBag.ReadProperty("Bold", False)
        .Font.Italic = PropBag.ReadProperty("Italic", False)
        .Font.Strikethrough = PropBag.ReadProperty("Strike", False)
        .Font.Underline = PropBag.ReadProperty("Underline", False)
    End With
    mblnEnabled = PropBag.ReadProperty("Enabled", True)
    mudtColor = PropBag.ReadProperty("Color", 0) '0 is red

End Sub

Private Sub UserControl_Resize()

    If mblnBuilding Then Exit Sub 'see comment in BuildMatrix
    picText.Width = (UserControl.Width \ (7 * Screen.TwipsPerPixelX)) - 1
    BuildMatrix

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Text", mstrText, "DotMatrix"
    With picText
        PropBag.WriteProperty "Font", .Font, .Font
        PropBag.WriteProperty "Bold", .Font.Bold, False
        PropBag.WriteProperty "Italic", .Font.Italic, False
        PropBag.WriteProperty "Strike", .Font.Strikethrough, False
        PropBag.WriteProperty "Underline", .Font.Underline, False
        PropBag.WriteProperty "Size", .Font.Size, 8
    End With
    PropBag.WriteProperty "Enabled", mblnEnabled, True
    PropBag.WriteProperty "Color", mudtColor, 0

End Sub
