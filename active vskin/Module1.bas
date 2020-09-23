Attribute VB_Name = "function"

    Public Enum enumAlignControl
  vsBottom
  vsTop
  vsRight
  vsLeft
  vsTopLeft
  vsTopRight
  vsBottomLeft
  vsBottomright
End Enum
Public Enum anchorss
vbleft
vbright
vbDown
vbup
vbLeft_Right
vbUp_Down
End Enum

Public Function filtertxt(Text As String, filtered As String) As String
Dim filt1(30)
Dim u As Long
Y = filtered
u = Len(Y)
For g = 1 To u
filt1(g) = Mid(Y, g, 1)
Next g
For i = 0 To u
Text = Replace(Text, filt1(i), "", 1, Len(Text))
Next i
filtertxt = Text

End Function
Public Sub enabletabs(stat As Boolean)

End Sub

Public Sub login()
With Form1
.Enabled = False
.Command1.Enabled = False
.chameleonButton3.Enabled = False
.Check1.Enabled = False
.Winsock2.LocalPort = lport
.Winsock2.Listen
.Winsock1.Connect cip, cport
End With

End Sub
Public Sub encode()
Open App.Path & "\users.txt" For Random As #444
Put #444, 1, Form1.Text1.Text
Put #444, 2, Form1.Text2.Text
Close #444
EncodeFile App.Path & "\users.txt", App.Path & "\users", "sawsana", "G"
Kill App.Path & "\users.txt"
End Sub
Public Function reverse(strr) As String
X = Len(strr)

Do Until X = 0
Y = Mid(strr, X, 1)
h = h & Y
X = X - 1
Loop
If X = 0 Then
reverse = h
End If
End Function

Public Function alg(a3lg, fontnam, fontsizn, fontstri, fontund, fontcolo, cha, use)
Dim full As String
Dim fontname As String
Dim fontsize As String
Dim fontstrik As String
Dim fontundr As String
Dim fontcolor As Long


Dim chat As String
Dim usr As String
Dim find1 As String

full = reverse(a3lg)
For ay = 1 To 5
find1 = InStr(full, ")\(")
Select Case ay
Case 1
fontname = Mid(full, 1, find1 - 1)
fontname = reverse(fontname)
Case 2
fontstrik = Mid(full, 1, find1 - 1)
fontstrik = reverse(fontstrik)
Case 3
fontundr = Mid(full, 1, find1 - 1)
fontundr = reverse(fontundr)
Case 4
fontsize = Mid(full, 1, find1 - 1)
fontsize = reverse(fontsize)
Case 5
fontcolor = Mid(full, 1, find1 - 1)
fontcolor = reverse(fontcolor)

End Select
full = Mid(full, find1 + 3, Len(full))
Next ay

chat = reverse(full)
find1 = InStr(a3lg, ":")
usr = Mid(a3lg, 1, find1 - 1)
chat = Mid(chat, find1 + 1, Len(chat))



fontnam = fontname
fontsizn = fontsize
fontstri = fontstrik
fontund = fontundr
fontcolo = fontcolor


use = usr
cha = chat

End Function
Public Sub disconnected()
With Form1
.Winsock1.Close
.Winsock2.Close
.Enabled = True
.Command1.Enabled = True
.chameleonButton3.Enabled = True
.Check1.Enabled = True
.Label5 = "Disconnected"

End With

End Sub
Public Sub connected()
With Form1
.Label5 = UCase("connected")
.Enabled = True
cuser = .Text1
.Timer1.Enabled = False
If .Check1.Value = 1 Then encode
.Text1.Enabled = False
.Text2.Enabled = False
.Text3.Enabled = False
End With

End Sub
Public Sub hothere(toput)


alg toput, fontname1, fontsize1, fontstri1, fontund1, fontcolor1, chat1, use1



webcol fontcolor1, fontcolor1


With Form1
With .WebBrowser1
.Document.BODY.INNERHTML = .Document.BODY.INNERHTML & "<font face='" & fontname1 & "'><font color=" & fontcolor1 & ">" & chat1 & "</font>" & vbCrLf





'.SelStart = Len(.Text)
'.SelText = vbNewLine
'.SelStart = Len(.Text)
'.SelColor = fontcolor1
'.SelFontSize = fontsize1
'.SelFontName = "verdana"
'.SelText = usr & ":"
'.SelStrikeThru = fontstrik1
'.SelUnderline = fontundr1
'.SelFontName = fontname1
'.SelText = chat1
     
End With
End With

End Sub

Public Function isinside(obj As Object) As Boolean
Dim place As POINTAPI
GetCursorPos place
If WindowFromPoint(place.X, place.Y) <> obj.hWnd Then
isinside = False
Else
isinside = True
End If

End Function

Public Function webcol(lngColour, Result) As String
    Dim strColour As String

    
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    
    'Add leading zero's


    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    'Reverse the bgr string pairs to rgb
    Result = "#" & Mid(strColour, Len(strColour) - 1, 2) & _
    Mid(strColour, 3, 2) & _
    Mid(strColour, 2)
    
End Function



Option Explicit
Dim doit As Boolean
Private Const DT_EDITCONTROL = &H2000&
Private Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type



Public Sub TextTrans(MyTB As Object)
Dim TempDC As Long
Dim temp As String

Dim MyLoc As RECT
temp = MyTB.Text

MyLoc.left = MyTB.left
MyLoc.top = MyTB.top
MyLoc.Right = MyLoc.left + MyTB.width
MyLoc.Bottom = MyLoc.top + MyTB.height
MyTB.Parent.Cls
MyTB.Parent.ForeColor = MyTB.ForeColor


Set MyTB.Parent.Font = MyTB.Font
'DrawText MyTB.parent.hdc, temp, Len(temp), MyLoc, DT_EDITCONTROL
'TempDC = GetDC(MyTB.hwnd)
'BitBlt TempDC, 0, 0, MyTB.Width, MyTB.Height, MyTB.parent.hdc, MyTB.Left, MyTB.Top, vbSrcCopy
End Sub

Public Sub CreateEllipse(obj As Object)
SetWindowRgn obj.hWnd, CreateEllipticRgn(0, 0, obj.width \ 15, obj.height \ 15), True
End Sub
Public Sub makeroundrect(obj As Object, degree)
SetWindowRgn obj.hWnd, CreateRoundRectRgn(0, 0, obj.width \ 15, obj.height \ 15, degree, degree), True
End Sub

Public Function MakeTrans(zForm As Object, Optional ByVal TransColor As Long = &HFF00FF) As Boolean

  Dim Msg As Double
  
  
  Msg = GetWindowLong(zForm.hWnd, GWL_EXSTYLE)
  
  
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.hWnd, TransColor, 0, LWA_COLORKEY
 
End Function
Public Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long


  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hdc, 0, 0)
  End If
  lngHeight& = picSource.height / Screen.TwipsPerPixelY
  lngWidth& = picSource.width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function

Public Function TileImage(TheImage As Object, TheDestination As Object, DestX, DestY, SrcY, SrcX, WidthToStartFrom, HeightToStartFrom)
'For D = 0 To Me.ScaleWidth
'BitBlt TheDestination.hdc, DestX, DestY, SrcX, SrcY, TheImage.hdc, WidthToStartFrom, HeightToStartFrom, SRCCOPY
'Next
End Function

Public Function moveit(hnd)
Call ReleaseCapture
SendMessage hnd, &HA1, 2, 0&
End Function




Public Sub AttachControl(ByRef ctlControl2Move As Object, ByVal ctlExistingControl As Object, ByVal lngAlignLocation As enumAlignControl, Optional ByVal lngHSpaceInTwips As Long = 0, Optional ByVal lngVSpaceInTwips As Long = 0)


  With ctlControl2Move
    Select Case lngAlignLocation
      Case vsBottom
        .top = ctlExistingControl.top + ctlExistingControl.height + lngVSpaceInTwips
      Case vsTop
        .top = ctlExistingControl.top - .height - lngVSpaceInTwips
      Case vsRight
        .left = ctlExistingControl.left + ctlExistingControl.width '- .Width '- lngHSpaceInTwips
      Case vsLeft
            .left = ctlExistingControl.left - .height
      Case vsTopLeft
        .Move ctlExistingControl.left + lngHSpaceInTwips, ctlExistingControl.top - .height - lngVSpaceInTwips
      Case vsTopRight
       .Move ctlExistingControl.left + ctlExistingControl.width - .width - lngHSpaceInTwips, ctlExistingControl.top - .height - lngVSpaceInTwips
'ctlControl2Move.Move ctlExistingControl.Left + ctlExistingControl.width, ctlExistingControl.Top  '- lngVSpaceInTwips
      Case vsBottomLeft
        .Move ctlExistingControl.left + lngHSpaceInTwips, ctlExistingControl.top + ctlExistingControl.height + lngVSpaceInTwips
      Case vsBottomright
        .Move ctlExistingControl.left + ctlExistingControl.width - .width - lngHSpaceInTwips, ctlExistingControl.top + ctlExistingControl.height + lngVSpaceInTwips
    End Select
  End With
  




End Sub
Public Sub resizeit(obj As Object, options As anchorss)
With obj
        marginLeft = .left
        marginTop = .top
        marginRight = .Container.width - .left - .width
        marginBottom = .Container.height - .top - .height
    End With
    
  
Select Case options

Case vbleft
obj.left = marginLeft

Case vbLeft_Right
i = obj.Container.width - marginLeft - marginRight
If i > 0 Then obj.width = i

Case vbright
i = obj.Container.width - obj.width - marginRight
If i > 0 Then obj.left = i

Case vbup
Container.top = marginTop

Case vbUp_Down
i = obj.Container.height - marginTop - marginBottom
If i > 0 Then obj.height = i

Case vbDown
i = obj.Container.height - obj.height - marginBottom
If i > 0 Then obj.top = i
End Select
    
End Sub
Public Sub Freeze(ByVal ohWnd As Long)
    'Freezing the Application
    SendMessage ohWnd, WM_SETREDRAW, False, 0
End Sub

Public Sub UnFreeze(ByVal ohWnd As Long, Optional ByVal ForceUpdate As Boolean = True)
    'Unfreezing the Application
    SendMessage ohWnd, WM_SETREDRAW, 1&, 0
    'SendMessage ohWnd, WM_PAINT, 1&, 0
    If ForceUpdate Then
        RedrawWindow ohWnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
    End If
End Sub



Public Sub putontop(hnd)
    SetWindowPos hnd, -1, 0, 0, 0, 0, 3

End Sub
 
