Attribute VB_Name = "modReadOnlyWnd"
Option Explicit
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'The IsWindowEnabled function determines whether the specified window
'is enabled for mouse and keyboard input.
'If the window is enabled, the return value is nonzero.
'If the window is not enabled, the return value is zero.

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'The LockWindowUpdate function disables or enables drawing in the specified window.
'Only one window can be locked at a time.

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'The EnableWindow function enables or disables mouse and keyboard input
'to the specified window or control (what??). When input is disabled,
'the window does not receive input such as mouse clicks and key presses.
'When input is enabled, the window receives all input blah blah blah.
'


Public Function IsEnabledW(chWnd As Long) As Boolean
       
       If IsWindowEnabled(chWnd) <> 0 Then
           IsEnabledW = True
          Else
           IsEnabledW = False
       End If
       
End Function


Public Function SwitchReadOnly(objForm As Form, cExclude As Control)
'cExclude covers the command buttons. We dont want to exlcude a button that says Close
'on a custom made form.

Dim oControl As Control

LockWindowUpdate objForm.hwnd 'Use this func for a "Heavy Loaded" form to avoid flicker

  For Each oControl In objForm.Controls
       Debug.Print TypeName(oControl)
       
    
            
            Select Case TypeName(oControl)
                   Case "Label", "Line", "Menu", "Image", "Shape"
                         GoTo ncon
                   Case Else
                         
                         If TypeName(oControl) = TypeName(cExclude) Then
                            EnableWindow oControl.hwnd, True
                            GoTo ncon
                         End If
                         
                         
                         EnableWindow oControl.hwnd, Not IsEnabledW(oControl.hwnd)
                                                    'Reverse the Enable bit
            End Select


ncon:
  
  Next oControl



LockWindowUpdate ByVal 0
Set oControl = Nothing

End Function

