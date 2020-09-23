Attribute VB_Name = "WindowPosSize"
Option Explicit

Private Type POINTAPI
 x As Long
 y As Long
End Type

Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Private Type WINDOWPLACEMENT
 Length As Long
 flags As Long
 showCmd As Long
 ptMinPosition As POINTAPI
 ptMaxPosition As POINTAPI
 rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Function GetPosAndSize(frmMe As Form) As Collection
 Set GetPosAndSize = New Collection
 
 Dim WinPlace As WINDOWPLACEMENT
 Dim rctTemp As RECT
 
 WinPlace.Length = Len(WinPlace)
 
 'Get the current Window's placement:
 GetWindowPlacement frmMe.hwnd, WinPlace
 rctTemp = WinPlace.rcNormalPosition
 
 With GetPosAndSize
  .Add "" & frmMe.ScaleX(rctTemp.Left, vbPixels, vbTwips), "Left"
  .Add "" & frmMe.ScaleY(rctTemp.Top, vbPixels, vbTwips), "Top"
  .Add "" & frmMe.ScaleX(rctTemp.Right - rctTemp.Left, vbPixels, vbTwips), "Width"
  .Add "" & frmMe.ScaleY(rctTemp.Bottom - rctTemp.Top, vbPixels, vbTwips), "Height"
 End With
End Function
