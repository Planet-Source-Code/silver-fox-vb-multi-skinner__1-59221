VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbskinnerpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents forma As Form
Attribute forma.VB_VarHelpID = -1
Public Event MouseMove(X As Single, Y As Single, Button As Integer, Style)

Private Enum places
vright
vleft
End Enum
Private WithEvents left_up_pic As imagebox
Attribute left_up_pic.VB_VarHelpID = -1
Private WithEvents left_middle_pic As Image
Attribute left_middle_pic.VB_VarHelpID = -1
Private WithEvents left_connect_pic As imagebox
Attribute left_connect_pic.VB_VarHelpID = -1
Private WithEvents down_pic As imagebox
Attribute down_pic.VB_VarHelpID = -1
Private WithEvents title_pic As imagebox
Attribute title_pic.VB_VarHelpID = -1
Private WithEvents middle_pic As Shape
Attribute middle_pic.VB_VarHelpID = -1
Private WithEvents right_up_pic As imagebox
Attribute right_up_pic.VB_VarHelpID = -1
Private WithEvents right_middle_pic As Image
Attribute right_middle_pic.VB_VarHelpID = -1
Private WithEvents right_connect_pic As imagebox
Attribute right_connect_pic.VB_VarHelpID = -1
Private WithEvents close_but As imagebox
Attribute close_but.VB_VarHelpID = -1
Private WithEvents min_but As imagebox
Attribute min_but.VB_VarHelpID = -1
Private WithEvents max_but As imagebox
Attribute max_but.VB_VarHelpID = -1
Private WithEvents mgc_but As imagebox
Attribute mgc_but.VB_VarHelpID = -1
Private drawonfrm As Boolean
Dim shaperes As Object
Private WithEvents skin As ImageList
Attribute skin.VB_VarHelpID = -1
Dim bod_col As Double
Dim WithEvents timers As Timer
Attribute timers.VB_VarHelpID = -1
Dim ctlcount As Double
Dim cap As String
Dim WithEvents lbl_title As Label
Attribute lbl_title.VB_VarHelpID = -1
Dim frst As Boolean
Dim cap_col As Double
Dim icn_title As Image
Dim trans_col As Double
Dim com_color As Double
Private Sub Class_Terminate()
forma.Controls.Remove ("P111")
forma.Controls.Remove ("P121")
forma.Controls.Remove ("P131")
forma.Controls.Remove ("P151")
forma.Controls.Remove ("P11")
forma.Controls.Remove ("P22")
forma.Controls.Remove ("P77")
forma.Controls.Remove ("P88")
forma.Controls.Remove ("P99")
forma.Controls.Remove ("P55")
forma.Controls.Remove ("P355")
forma.Controls.Remove ("P33")
forma.Controls.Remove ("P66")
forma.Controls.Remove ("P44")
'forma.Controls.Remove ("P869")
End Sub

Private Sub close_but_Click()
End

End Sub

Private Sub close_but_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set close_but.Picture = skin.ListImages("close_active").Picture
End Sub

Private Sub close_but_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set close_but.Picture = skin.ListImages("close_idle").Picture
End Sub

Private Sub down_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
If Button = 1 Or Button = 3 Then
checksize 0, Y, vright
End If
End If

End Sub

Private Sub forma_Load()
forma.visible = False

End Sub

Private Sub forma_Resize()
If forma.BorderStyle <> 0 Then
forma.visible = False
resizeit True, True, True
forma.BorderStyle = 0
lbl_title.Caption = forma.Caption
forma.Caption = ""
forma.Caption = lbl_title.Caption
If frst = False Then
resizeit True, True, True
forma.Refresh
forma.visible = True
frst = True
End If
End If


End Sub

Private Sub lbl_title_DblClick()
If forma.WindowState = 0 Then
forma.WindowState = 2
Else
forma.WindowState = 0
End If
resizeit True, True, True

End Sub

Private Sub lbl_title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
moveit forma.hwnd
End If

End Sub

Private Sub left_connect_pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
If Button = 1 Or Button = 3 Then
checksize X, Y, vleft
End If
End If
End Sub

Private Sub left_middle_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
If Button = 1 Then
checksize X, 0, vleft
End If
End If

End Sub

Private Sub max_but_Click()
If forma.WindowState = 2 Then
Set max_but.Picture = skin.ListImages("maxi_ideal").Picture
forma.WindowState = 0
resizeit True, True, True
Else
Set max_but.Picture = skin.ListImages("max_ideal").Picture
forma.WindowState = 2
resizeit True, True, True
End If

End Sub

Private Sub max_but_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Set max_but.Picture = skin.ListImages("max_active").Picture
Else
Set max_but.Picture = skin.ListImages("maxi_active").Picture
End If

End Sub

Private Sub max_but_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set max_but.Picture = skin.ListImages("maxi_ideal").Picture
End Sub

Private Sub mgc_but_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set mgc_but.Picture = skin.ListImages("wzrd_active").Picture
End Sub

Private Sub mgc_but_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set mgc_but.Picture = skin.ListImages("wzrd_idle").Picture
End Sub

Private Sub min_but_Click()
forma.WindowState = 1


End Sub

Private Sub min_but_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set min_but.Picture = skin.ListImages("min_active").Picture
End Sub

Private Sub min_but_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set min_but.Picture = skin.ListImages("min_ideal").Picture
End Sub

Private Sub right_connect_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
If Button = 1 Then
checksize X, Y, vright
End If
End If

End Sub

Private Sub right_middle_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
If Button = 1 Then
checksize X, 0, vright
End If

End If
End Sub

Private Sub timers_Timer()
If ctlcount <> forma.Controls.Count Then
title_pic.Zorder = 0
right_up_pic.Zorder = 0
max_but.Zorder = 0
min_but.Zorder = 0
mgc_but.Zorder = 0
close_but.Zorder = 0
lbl_title.Zorder 0

ctlcount = forma.Controls.Count
End If
If lbl_title.Caption <> forma.Caption Then
lbl_title.Caption = forma.Caption
End If
If icn_title.Picture <> forma.Icon Then
icn_title.Picture = forma.Icon
End If

'putontop title_pic.hwnd

End Sub

Private Sub title_pic_DblClick()
If forma.WindowState = 0 Then
forma.WindowState = 2
Else
forma.WindowState = 0
End If
resizeit True, True, True

End Sub

Private Sub title_pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If forma.WindowState = 2 Then
Else
moveit forma.hwnd
End If

End Sub
Private Sub title_pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
moveit forma.hwnd
End Sub
Private Sub resizeit(vertical_left As Boolean, vertical_right As Boolean, horizintal As Boolean)
'Set title_pic.Picture = Nothing
If vertical_left = True Then
title_pic.width = forma.width - right_up_pic.width
right_up_pic.Move (forma.width - right_up_pic.width), 0, right_up_pic.width, right_up_pic.height
right_middle_pic.Move forma.width - right_middle_pic.width, right_middle_pic.Top, right_middle_pic.width, right_middle_pic.height
close_but.Left = (forma.width - close_but.width) - 150
right_connect_pic.Move forma.width - right_connect_pic.width, right_connect_pic.Top, right_connect_pic.width, right_connect_pic.height
AttachControl left_middle_pic, left_up_pic, vsBottomLeft
End If

If vertical_right = True Then

title_pic.width = forma.width - right_up_pic.width

right_up_pic.Move (forma.width - right_up_pic.width), 0, right_up_pic.width, right_up_pic.height
right_connect_pic.Move forma.width - right_connect_pic.width, right_connect_pic.Top, right_connect_pic.width, right_connect_pic.height

right_middle_pic.Move forma.width - right_middle_pic.width, right_middle_pic.Top, right_middle_pic.width, right_middle_pic.height
close_but.Left = (forma.width - close_but.width) - 150
End If

If horizintal = True Then

'left_middle_pic.height = (forma.height) - (left_connect_pic.height) * 3 + 10
left_connect_pic.Top = forma.height - left_connect_pic.height
'right_middle_pic.height = (forma.height) - (right_connect_pic.height * 3) - 30
'right_middle_pic.Top = right_up_pic.height
right_connect_pic.Top = forma.height - right_connect_pic.height
down_pic.Top = forma.height - down_pic.height
End If

 'strech the left side
' ********* pins the controls *********



'down_pic.Left = left_connect_pic.width
middle_pic.Left = 0
middle_pic.Top = 0
middle_pic.width = forma.width
middle_pic.height = forma.height

max_but.Left = (forma.width - close_but.width) - (70 * 5) - 90
min_but.Left = (forma.width - close_but.width) - (70 * 9) - 70
mgc_but.Left = (forma.width - close_but.width) - (70 * 13) - 70

lbl_title.Top = 0
lbl_title.Left = title_pic.width \ 2 - 400
lbl_title.Top = 100
icn_title.Left = 0
icn_title.Top = 90

icn_title.width = 240
icn_title.height = 240

close_but.Top = 120
max_but.Top = 120
min_but.Top = 120
mgc_but.Top = 120


End Sub
Private Sub checksize(Xx, Yy, place As places)
'On Error Resume Next
Dim gh As Object
Dim zeyadax, zeyaday As Boolean
Select Case place
Case vright

If forma.width + Xx < 2005 Then
zeyadax = True
forma.width = 2005
resizeit False, True, False
End If

If forma.height + Yy < 2005 Then
zeyaday = True
forma.height = 2005
forma.height = 2005
resizeit False, False, True
End If

If forma.height + Yy > Screen.height Then
zeyaday = True
forma.height = Screen.height
forma.height = Screen.height
resizeit False, False, True
End If

If forma.width + Xx > Screen.width Then
zeyadax = True
forma.width = Screen.width
resizeit False, True, False
End If

If zeyadax = False Then
If Xx <> 0 Then
forma.width = forma.width + Xx
resizeit False, True, False
End If

End If

If zeyaday = False Then
If Yy <> 0 Then
forma.height = forma.height + Yy
resizeit False, False, True
End If
End If


Case vleft

If forma.height + Yy < 2005 Then
zeyaday = True
forma.height = 2005
resizeit False, False, True
End If

If forma.height + Yy > Screen.height Then
zeyaday = True
forma.height = Screen.height
resizeit False, False, True
End If

If forma.width - Xx > Screen.width Then
zeyadax = True
forma.width = Screen.width
resizeit True, False, False
End If

If forma.width - Xx < 2005 Then
zeyadax = True
forma.Left = forma.Left + forma.width - 2005
forma.width = 2005
resizeit True, False, False
End If

If zeyadax = False Then
If Xx <> 0 Then

forma.Move forma.Left + Xx, forma.Top, forma.width - Xx, forma.height
resizeit True, False, False
End If
End If

If zeyaday = False Then
If Yy <> 0 Then
forma.height = forma.height + Yy
resizeit False, False, True
End If
End If
End Select


End Sub
Public Sub apply_skin(pic_buffer As PictureBox, pic_source As PictureBox, Image1 As Image)
   'On Error GoTo apply_skin_Error

If pic_source.Picture = Empty Then
pic_source.Picture = LoadResPicture(101, 0)
End If

skin.ListImages.Clear


'******************************* Now The Real Work Starts**********************************************
pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, 0, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "close_idle", Image1.Picture


pic_buffer.width = 13.5 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, 13 + 1, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "close_active", Image1.Picture

pic_buffer.width = 13.5 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 2) + 2, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "close_dis", Image1.Picture

pic_buffer.width = 13.5 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13.5, 13, pic_source.hdc, (13 * 3) + 3, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "max_ideal", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 4) + 4.5, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "max_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 5) + 5.5, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "max_dis", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 6) + 6, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "maxi_ideal", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 7) + 7, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "maxi_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 8) + 8, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "maxi_dis", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 9) + 9, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "min_ideal", Image1.Picture

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (13 * 10) + 10, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "min_active", Image1.Picture

pic_buffer.width = 14 * Screen.TwipsPerPixelX
pic_buffer.height = 14 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 14, 14, pic_source.hdc, (13 * 11) + 15, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "left_up", Image1.Picture

pic_buffer.width = 19 * Screen.TwipsPerPixelX
pic_buffer.height = 9 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 19, 9, pic_source.hdc, (13 * 12) + 17, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "up", Image1.Picture

pic_buffer.width = 14 * Screen.TwipsPerPixelX
pic_buffer.height = 14 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 14, 14, pic_source.hdc, (13 * 13) + 24, 0, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "right_up", Image1.Picture

pic_buffer.width = 19 * Screen.TwipsPerPixelX
pic_buffer.height = 30 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 19, 30, pic_source.hdc, 0, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "left_shded", Image1.Picture


pic_buffer.width = 19 * Screen.TwipsPerPixelX
pic_buffer.height = 30 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 19, 30, pic_source.hdc, 20, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "middle_shded", Image1.Picture



pic_buffer.width = 19 * Screen.TwipsPerPixelX
pic_buffer.height = 30 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 19, 30, pic_source.hdc, 20 * 2, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "right_shded", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, 20 * 3, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "min_dis", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 14, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "upp_idle", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 28, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "upp_active", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 42, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "upp_dis", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 56, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "downn_idel", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 70, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "downn_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 84, 14, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "arr_idle", Image1.Picture


pic_buffer.width = 9 * Screen.TwipsPerPixelX
pic_buffer.height = 19 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 9, 19, pic_source.hdc, (20 * 3) + 98, 15, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "left_middle", Image1.Picture

pic_buffer.width = 7 * Screen.TwipsPerPixelX
pic_buffer.height = 7 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 7, 7, pic_source.hdc, (20 * 3) + 119, 27, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "body_color", Image1.Picture

pic_buffer.width = 9 * Screen.TwipsPerPixelX
pic_buffer.height = 19 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 9, 19, pic_source.hdc, (20 * 3) + 138, 15, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "right_middle", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, 20 * 3, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "arr_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 14, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "wzrd_idle", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 28, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "wzrd_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 42, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "mgc_idle", Image1.Picture


pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 56, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "mgc_active", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 70, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "qs_idle", Image1.Picture

pic_buffer.width = 13 * Screen.TwipsPerPixelX
pic_buffer.height = 13 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 13, 13, pic_source.hdc, (20 * 3) + 84, 28, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "qs_active", Image1.Picture


pic_buffer.width = 14 * Screen.TwipsPerPixelX
pic_buffer.height = 14 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 14, 14, pic_source.hdc, (20 * 3) + 98, 35, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "left_bottom", Image1.Picture

pic_buffer.width = 19 * Screen.TwipsPerPixelX
pic_buffer.height = 9 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 19, 9, pic_source.hdc, (20 * 3) + 113, 40, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "middle_bottom", Image1.Picture

pic_buffer.width = 14 * Screen.TwipsPerPixelX
pic_buffer.height = 14 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 14, 14, pic_source.hdc, (20 * 3) + 133, 35, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "right_bottom", Image1.Picture

pic_buffer.width = 6 * Screen.TwipsPerPixelX
pic_buffer.height = 6 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 6, 6, pic_source.hdc, 32, 45, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "bar_text", Image1.Picture

pic_buffer.width = 5 * Screen.TwipsPerPixelX
pic_buffer.height = 5 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 6, 6, pic_source.hdc, 80, 46, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "body_text", Image1.Picture

pic_buffer.width = 5 * Screen.TwipsPerPixelX
pic_buffer.height = 5 * Screen.TwipsPerPixelY

pic_buffer.Cls: BitBlt pic_buffer.hdc, 0, 0, 6, 6, pic_source.hdc, 141, 46, SRCCOPY
Image1.Picture = pic_buffer.Image: skin.ListImages.Add , "trans_col", Image1.Picture

pic_buffer.Picture = skin.ListImages("body_color").Picture
pic_buffer.ScaleMode = vbPixels
bod_col = GetPixel(pic_buffer.hdc, pic_buffer.ScaleWidth \ 2, pic_buffer.ScaleHeight \ 2)
pic_buffer.Picture = skin.ListImages("bar_text").Picture
 cap_col = GetPixel(pic_buffer.hdc, pic_buffer.ScaleWidth \ 2, pic_buffer.ScaleHeight \ 2)
pic_buffer.Picture = skin.ListImages("trans_col").Picture
trans_col = GetPixel(pic_buffer.hdc, pic_buffer.ScaleWidth \ 2, pic_buffer.ScaleHeight \ 2)

pic_buffer.Picture = skin.ListImages("body_text").Picture
cap_col = GetPixel(pic_buffer.hdc, pic_buffer.ScaleWidth \ 2, pic_buffer.ScaleHeight \ 2)

addcontrols

            '****************************** THE End *********************************

   On Error GoTo 0
   Exit Sub

apply_skin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure apply_skin of User Control skn"

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic_source,pic_source,-1,Picture

Private Sub addcontrols()
'makeroundrect forma, 0
'Set forma.ControlBox = 0
'forma.BorderStyle = 0

'If skin.ListImages.Count = 0 Then pic_source = default_pic.Picture
'apply_skin

forma.BackColor = bod_col
'
'forma.BackColor = trans_col
MakeTrans forma, trans_col
forma.Refresh
For Each object In forma.Controls
If TypeOf object Is CommandButton Then
object.BackColor = bod_col
End If
Next
Set min_but = forma.Controls.Add("project1.imagebox", "P111", forma)
Set max_but = forma.Controls.Add("project1.imagebox", "P121", forma)
Set close_but = forma.Controls.Add("project1.imagebox", "P131", forma)
Set mgc_but = forma.Controls.Add("project1.imagebox", "P151", forma)
Set left_up_pic = forma.Controls.Add("project1.imagebox", "P11", forma)
Set left_middle_pic = forma.Controls.Add("vb.image", "P22", forma)
Set right_up_pic = forma.Controls.Add("project1.imagebox", "p77", forma)
Set right_middle_pic = forma.Controls.Add("vb.image", "p88", forma)
Set right_connect_pic = forma.Controls.Add("project1.imagebox", "p99", forma)
Set title_pic = forma.Controls.Add("project1.imagebox", "p55", forma)
Set right_connect_pic = forma.Controls.Add("project1.imagebox", "P355", forma)
Set left_connect_pic = forma.Controls.Add("Project1.imagebox", "P33", forma)
Set down_pic = forma.Controls.Add("Project1.imagebox", "p44", forma)
Set middle_pic = forma.Controls.Add("VB.shape", "p66", forma)
Set lbl_title = forma.Controls.Add("vb.label", "p868", title_pic)
Set icn_title = forma.Controls.Add("vb.image", "p869", title_pic)


lbl_title.visible = True
lbl_title.BackStyle = 0

left_up_pic.visible = True
left_middle_pic.visible = True
left_connect_pic.visible = True
down_pic.visible = True
title_pic.visible = True
middle_pic.visible = False
right_up_pic.visible = True
right_middle_pic.visible = True
right_connect_pic.visible = True

min_but.visible = True
max_but.visible = True
close_but.visible = True
mgc_but.visible = True

left_up_pic.height = 495
left_up_pic.width = 255
left_middle_pic.width = 125
left_connect_pic.height = 255
left_connect_pic.width = 255
down_pic.height = 135
title_pic.height = 495
right_up_pic.height = 495
right_up_pic.width = 255
right_middle_pic.width = 125
right_connect_pic.height = 255
right_connect_pic.width = 255
left_up_pic.Left = 0
left_up_pic.Top = 0
title_pic.Top = 0
title_pic.Left = left_up_pic.width
down_pic.Left = left_connect_pic.width - 100
middle_pic.Left = left_middle_pic.width
'middle_pic.BackStyle = 1





Set min_but.Picture = skin.ListImages("min_ideal").Picture

Set max_but.Picture = skin.ListImages("maxi_ideal").Picture

Set close_but.Picture = skin.ListImages("close_idle").Picture
Set mgc_but.Picture = skin.ListImages("wzrd_idle").Picture

Set title_pic.Picture = skin.ListImages("middle_shded").Picture

Set right_up_pic.Picture = skin.ListImages("right_shded").Picture
Set left_up_pic.Picture = skin.ListImages("left_shded").Picture

Set right_middle_pic.Picture = skin.ListImages("right_middle").Picture

Set right_connect_pic.Picture = skin.ListImages("right_bottom").Picture
'left_up_pic.Picture = Picture1.Picture
Set left_middle_pic.Picture = skin.ListImages("left_middle").Picture
Set left_connect_pic.Picture = skin.ListImages("left_bottom").Picture
Set down_pic.Picture = skin.ListImages("middle_bottom").Picture
'middle_pic.Picture = Image1.Picture


title_pic.Stretch = True
right_up_pic.Stretch = True
right_middle_pic.Stretch = True
'right_connect_pic.Stretch = True
left_up_pic.Stretch = True
left_middle_pic.Stretch = True

'left_connect_pic.Stretch = True
down_pic.Stretch = True
'middle_pic.Stretch = True
'title_pic.width = Screen.width
down_pic.width = Screen.width 'stretch the bottom
middle_pic.Top = 0
middle_pic.BackColor = bod_col
'forma.BackColor = bod_col
'middle_pic.width = Screen.width
'middle_pic.height = Screen.height
middle_pic.BackStyle = 1


timers.Enabled = True
lbl_title.ForeColor = cap_col
'lbl_title.AutoSize = True
lbl_title.Font = forma.Font
lbl_title.FontBold = True
icn_title.Stretch = True
right_middle_pic.height = Screen.height
right_middle_pic.Top = right_up_pic.height
left_middle_pic.height = Screen.height
left_middle_pic.Top = left_up_pic.height
icn_title.visible = True
left_connect_pic.MousePointer = 6
left_middle_pic.MousePointer = 9
right_middle_pic.MousePointer = 9
right_connect_pic.MousePointer = 8
down_pic.MousePointer = 7


resizeit True, True, True
End Sub
Public Function intilazing(forma_To_Intiliaze As Form, skinn As ImageList, pic_buffer As PictureBox, pic_source As PictureBox, Image1 As Image, Timer1 As Timer)
Set skin = skinn

Set timers = Timer1
Set forma = forma_To_Intiliaze
apply_skin pic_buffer, pic_source, Image1
End Function
