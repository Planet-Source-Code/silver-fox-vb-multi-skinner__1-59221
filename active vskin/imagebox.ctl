VERSION 5.00
Begin VB.UserControl imagebox 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ControlContainer=   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   1275
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "imagebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Dim busy As Boolean
'Default Property Values:
Const m_def_Zorder = 0
Const m_def_visible = 0
Const m_def_Heigt = 0
Const m_def_wdth = 0
'Property Variables:
Dim m_Zorder As Variant
Dim m_visible As Variant
Dim m_Heigt As Variant
Dim m_wdth As Variant


Private Sub Image1_Click()
RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Resize()

If busy = True Then Exit Sub

Image1.width = UserControl.width
Image1.height = UserControl.height

End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub



Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   
    Set Image1.Picture = New_Picture
    
     If Me.Stretch = False Then
       busy = True
    UserControl.width = Image1.width
    UserControl.height = Image1.height
    busy = False
    End If
    PropertyChanged "Picture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Image1.Stretch = PropBag.ReadProperty("Stretch", False)
    UserControl.height = PropBag.ReadProperty("Height", m_def_Heigt)
   UserControl.width = PropBag.ReadProperty("width", m_def_wdth)
     Extender.Top = PropBag.ReadProperty("Top", 0)
   Extender.Left = PropBag.ReadProperty("Left", 0)
    
    Extender.visible = PropBag.ReadProperty("visible", True)
   Extender.Zorder = PropBag.ReadProperty("Zorder", m_def_Zorder)
    Image1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Stretch", Image1.Stretch, False)
    Call PropBag.WriteProperty("Height", UserControl.height, m_def_Heigt)
    Call PropBag.WriteProperty("width", UserControl.width, m_def_wdth)
     Call PropBag.WriteProperty("Top", Extender.Top, 0)
    Call PropBag.WriteProperty("Left", Extender.Left, 0)
    Call PropBag.WriteProperty("visible", Extender.visible, True)
    
    Call PropBag.WriteProperty("Zorder", Extender.Zorder, m_def_Zorder)
    Call PropBag.WriteProperty("MousePointer", Image1.MousePointer, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Stretch
Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control."
    Stretch = Image1.Stretch
End Property

Public Property Let Stretch(ByVal New_Stretch As Boolean)
    Image1.Stretch() = New_Stretch
       If New_Stretch = False Then
          busy = True
    UserControl.width = Image1.width
    UserControl.height = Image1.height
    busy = False
    End If
    PropertyChanged "Stretch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get height() As Variant
    height = UserControl.height

End Property

Public Property Let height(ByVal New_Height As Variant)
    UserControl.height = New_Heigt
    PropertyChanged "Height"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get width() As Variant
    width = UserControl.width
End Property

Public Property Let width(ByVal New_wdth As Variant)
    UserControl.width = New_wdth
    PropertyChanged "width"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Heigt = m_def_Heigt
    m_wdth = m_def_wdth
    m_visible = m_def_visible
    m_Zorder = m_def_Zorder
End Sub


Public Property Get Top() As Variant
Top = Extender.Top
End Property

Public Property Let Top(ByVal vNewValue As Variant)
Extender.Top = vNewValue
PropertyChanged "Top"
End Property
Public Property Get Left() As Variant
Left = Extender.Left

End Property
Public Property Let Left(ByVal vNewValue As Variant)
Extender.Left = vNewValue
PropertyChanged "Left"
End Property

Public Function Move(Optional Left As Single = vbNull, Optional Top = vbNull, Optional width = vbNull, Optional height = vbNull)
Extender.Move Left, Top, width, height
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get visible() As Boolean
    visible = Extender.visible
End Property

Public Property Let visible(ByVal New_visible As Boolean)
    Extender.visible = New_visible
    PropertyChanged "visible"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Zorder() As Variant
    Zorder = Extender.Zorder
    
End Property

Public Property Let Zorder(ByVal New_Zorder As Variant)
    Extender.Zorder New_Zorder
    PropertyChanged "Zorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Image1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Image1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

