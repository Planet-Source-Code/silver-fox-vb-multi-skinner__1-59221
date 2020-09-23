VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl skn 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ClipBehavior    =   0  'None
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   825
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox default_pic 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1680
      Picture         =   "skn.ctx":0000
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox garbage 
      Height          =   15
      Left            =   2880
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   360
      Width           =   15
      Begin VB.PictureBox title_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   11
         Top             =   240
         Width           =   135
      End
      Begin VB.PictureBox bottom_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   10
         Top             =   600
         Width           =   135
      End
      Begin VB.PictureBox right_bottom_pic1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   9
         Top             =   480
         Width           =   135
      End
      Begin VB.PictureBox left_bottom_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   8
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox connect_right_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   7
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox connect_left_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         MousePointer    =   6  'Size NE SW
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   6
         Top             =   720
         Width           =   135
      End
      Begin VB.PictureBox left_pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   5
         Top             =   120
         Width           =   135
      End
      Begin VB.PictureBox right_pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   4
         Top             =   960
         Width           =   135
      End
      Begin VB.PictureBox middle_pic1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   3
         Top             =   360
         Width           =   135
         Begin VB.Image Image1 
            Height          =   255
            Left            =   360
            Top             =   2520
            Width           =   255
         End
         Begin VB.Image Image2 
            Height          =   255
            Left            =   840
            Top             =   2520
            Width           =   255
         End
         Begin VB.Image Image3 
            Height          =   255
            Left            =   1920
            Top             =   1440
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox pic_buffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   600
      ScaleHeight     =   615
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox pic_source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   595
      Left            =   840
      ScaleHeight     =   600
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin MSComctlLib.ImageList skin 
      Left            =   3840
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
End
Attribute VB_Name = "skn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim skn_pro As vbskinnerpro
Public Enum Styles
Vbskinpro = 0
'WinAmp = 1 Coming soon
'Skinable = 2 Coming soon
'Windows_blinds = 3 Coming soon
End Enum
'Default Property Values:
Const m_def_Style = 0
'Property Variables:
Dim m_Style As Styles
Dim applied As Boolean


Private Sub UserControl_InitProperties()

    m_Style = m_def_Style
End Sub
Public Property Get Style() As Styles
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As Styles)
    m_Style = New_Style

    PropertyChanged "Style"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set pic_source.Picture = PropBag.ReadProperty("skinnn", Nothing)
 m_Style = PropBag.ReadProperty("Style", m_def_Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("skinnn", pic_source.Picture, Nothing)
End Sub

Public Property Get skn() As StdPicture
    Set skn = pic_source.Picture
End Property

Public Property Set skn(ByVal New_skn As StdPicture)
    Set pic_source.Picture = New_skn
    PropertyChanged "skn"
End Property

Public Function Apply_skn()
If Ambient.UserMode = True Then
 Select Case m_Style
 Case Vbskinpro
 Set skn_pro = New vbskinnerpro
skn_pro.intilazing Parent, skin, pic_buffer, pic_source, Image1, Timer1
applied = True
'skn_pro.apply_skin pic_buffer, pic_source, Image1
End Select
End If
End Function
