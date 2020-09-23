<div align="center">

## SendKeys for Vista


</div>

### Description

Replaces the Visual Basic SendKeys statement which will not work in Vista. Uses the API call to the keybd_event to send characters and control codes. I have not extensively tested this routine, but it works for all the programs I have written to control external programs. The "WAIT" option is handled for compatibility, but is not implemented.
 
### More Info
 
Since this subroutine is named SendKeys, in most cases there is no need to re-code. Just place the new SendKeys subroutine in a public section of your project. However, you can no longer depend on the "WAIT" option. It is ignored in the new SendKeys.

I did not bother to re-create the triple code SendKeys function such as +(AC). My replacement sends only the control key (+) key and the "A" as a double key, but it sends the "C" as just a capitol "C".

I have not tested this routine in anything other than Visual Basic 5.0. I did not bother with some of the control fuctions like {BREAK}, {PRTSCR}, etc. You may want to.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ray B Dutton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ray-b-dutton.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ray-b-dutton-sendkeys-for-vista__1-69994/archive/master.zip)

### API Declarations

```
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
```


### Source Code

```
Public Sub SendKeys(St As String, Optional Wait As Boolean)
  '*****************************************************************************
  'Replacement for the Visual Basic SendKeys function. The optional Wait parameter
  'is included for compatibility only, but is ignored. The multiple key
  'function indicated by parentheses is handled but only the control key and next
  'key are treated as a multiple key stroke, not three. The next character(s)
  'is treated as a separate keystroke. The control keys +^% will be recognized
  'as standard characters unless they appear as the first character in the
  'SendKeys string.
  '
  'This new subroutine requires the following declarations in your project's form or
  'bas module:
  '
  'Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  '  ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
  '
  'Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
  '
  '*******************************************************************************
  Dim vbKCode As Variant
  Dim ShiftCtrlAlt As Variant
  Dim CapsLockState As Variant
  Dim keys(0 To 255) As Byte
  'Check the state of the CapsLock button to determine whether to
  'send or not send the SHIFT KEY
  GetKeyboardState keys(0)
  CapsLockState = keys(vbKeyCapital)
start:
  'Check for Shift, Ctrl, and Alt
  If InStr("+^%", Left$(St$, 1)) > 0 Then
    Select Case Left$(St$, 1)
      Case "+"
        ShiftCtrlAlt = vbKeyShift
      Case "^"
        ShiftCtrlAlt = vbKeyControl
      Case "%"
        ShiftCtrlAlt = vbKeyMenu
      Case Else
        ShiftCtrlAlt = ""
    End Select
  End If
  'Check for Special Keys
  If InStr(St$, "{") > 0 Then
    P1 = InStr(St$, "{")
    P2 = InStr(St$, "}")
    SpecialKey$ = Mid$(St$, P1, P2 - P1 + 1)
    Select Case SpecialKey$
      Case "{BACKSPACE}", "{BS}", "{BKSP}"
        vbKCode = vbKeyBack
      Case "{DELETE}", "{DEL}"
        vbKCode = vbKeyDelete
      Case "{DOWN}"
        vbKCode = vbKeyDown
      Case "{END}"
        vbKCode = vbKeyEnd
      Case "{ENTER}"
      Case "{ESC}"
        vbKCode = vbKeyEscape
      Case "{HELP}"
        vbKCode = vbKeyHelp
      Case "{HOME}"
        vbKCode = vbKeyHome
      Case "{INSERT}", "{INS}"
        vbKCode = vbKeyInsert
      Case "{LEFT}"
        vbKCode = vbKeyLeft
      Case "{NUMLOCK}"
        vbKCode = vbKeyNumlock
      Case "{PGDN}"
        vbKCode = vbKeyPageDown
      Case "{PGUP}"
        vbKCode = vbKeyPageUp
      Case "{RIGHT}"
        vbKCode = vbKeyRight
      Case "{SCROLLLOCK}"
        vbKCode = vbKeyScrollLock
      Case "{TAB}"
        vbKCode = vbKeyTab
      Case "{UP}"
        vbKCode = vbKeyUp
      Case "{F1}"
        vbKCode = vbKeyF1
      Case "{F2}"
        vbKCode = vbKeyF2
      Case "{F3}"
        vbKCode = vbKeyF3
      Case "{F4}"
        vbKCode = vbKeyF4
      Case "{F5}"
        vbKCode = vbKeyF5
      Case "{F6}"
        vbKCode = vbKeyF6
      Case "{F7}"
        vbKCode = vbKeyF7
      Case "{F8}"
        vbKCode = vbKeyF8
      Case "{F9}"
        vbKCode = vbKeyF9
      Case "{F10}"
        vbKCode = vbKeyF10
      Case "{F11}"
        vbKCode = vbKeyF11
      Case "{F12}"
        vbKCode = vbKeyF12
      Case "{F13}"
        vbKCode = vbKeyF13
      Case "{F14}"
        vbKCode = vbKeyF14
      Case "{F15}"
        vbKCode = vbKeyF15
      Case "{F16}"
        vbKCode = vbKeyF16
      Case Else
        vbKCode = ""
        Exit Sub
    End Select
    If ShiftCtrlAlt > 0 Then
      GoSub SendWithControl
    Else
      GoSub SendWithoutControl
    End If
    If Len(St$) > P2 Then
      'If there are more characters in the string,
      'remove those keys sent and start over.
      St$ = Mid$(St$, P2 + 1)
      GoTo start
    End If
    Exit Sub
  End If
  'Section to send a Control Key and a Character
  Set1$ = ")!@#$%^&*(" 'Characters above the numbers requiring SHIFT KEY
  Set2$ = "`-=[]\;',./" 'Other miscellaneous characters
  Set3$ = "~_+{}|:" & Chr(34) & "<>?" 'Miscellaneous characters requiring SHIFT KEY
  If ShiftCtrlAlt > 0 Then
    'Handle the three key problem which uses parentheses
    If InStr(St$, "(") > 0 Then
      'Strip the Parentheses.
      St$ = Mid$(St$, 1, 1) & Mid$(St$, 3, InStr(St$, ")") - 3)
    End If
    vbKCode = Asc(UCase(Mid$(St$, 2, 1)))
    'Check for characters 0 to 9, and A to Z. Scan codes same as ASCII
    If (vbKCode >= 48 And vbKCode <= 57) Or (vbKCode >= 65 And vbKCode <= 90) Then
      If ShiftCtrlAlt = vbKeyShift Then
        'Handle the problem of the CAPSLOCK
        If CapsLockState = False Then
          GoSub SendWithControl
        Else
          GoSub SendWithoutControl
        End If
      Else
        GoSub SendWithControl
      End If
    Else
      'Set the scan code for each miscellaneous character
      If InStr(Set1$, Mid$(St$, 2, 1)) > 0 Then
        vbKCode = 47 + InStr(Set1$, Mid$(St$, 2, 1))
      ElseIf InStr(Set2$, Mid$(St$, 2, 1)) > 0 Then
        vbKCode = Choose(InStr(Set2$, Mid$(St$, 2, 1)), 192, 189, 187, 219, _
        221, 220, 186, 222, 188, 190, 191)
      ElseIf InStr(Set3$, Mid$(St$, 2, 1)) > 0 Then
        vbKCode = Choose(InStr(Set3$, Mid$(St$, i, 1)), 192, 189, 187, 219, _
        221, 220, 186, 222, 188, 190, 191)
      End If
    End If
    'If there are more characters to print, remove the control key
    'and the first character and go to the next section. No control characters
    'processed beyond this point.
    If Len(St$) > 2 Then
      St$ = Mid$(St$, 3)
    Else
      Exit Sub
    End If
  End If
  '********* SEND CHARACTER STRING **********
  'Send all remaining characters as text, including control type characters
  'such as (+^%{[) etc.
  ShiftCtrlAlt = vbKeyShift 'Prepare to send the SHIFT KEY when needed
  'Set the scan code for each character and then send it
  For i = 1 To Len(St$)
    vbKCode = Asc(UCase(Mid$(St$, i, 1)))
    If InStr(Set1$, Mid$(St$, i, 1)) > 0 Then
      vbKCode = 47 + InStr(Set1$, Mid$(St$, i, 1))
      GoSub SendWithControl
    ElseIf InStr(Set2$, Mid$(St$, i, 1)) > 0 Then
      vbKCode = Choose(InStr(Set2$, Mid$(St$, i, 1)), 192, 189, 187, 219, 221, _
      220, 186, 222, 188, 190, 191)
      GoSub SendWithoutControl
    ElseIf InStr(Set3$, Mid$(St$, i, 1)) > 0 Then
      vbKCode = Choose(InStr(Set3$, Mid$(St$, i, 1)), 192, 189, 187, 219, 221, _
      220, 186, 222, 188, 190, 191)
      GoSub SendWithControl
    Else
      'Check to see if the character is upper or lower case and then
      'determine whether to send the SHIFT KEY based upon whether or not
      'the CAPSLOCK is set.
      If Asc(Mid$(St$, i, 1)) > vbKCode Then 'If true character is to be lowercase
        If CapsLockState = False Then
          GoSub SendWithoutControl
        Else
          GoSub SendWithControl
        End If
      Else
        If CapsLockState = False Then
          GoSub SendWithControl
        Else
          GoSub SendWithoutControl
        End If
      End If
    End If
  Next i
  Exit Sub
'API call to send a Control Code and a Character
SendWithControl:
  keybd_event ShiftCtrlAlt, 0, 0, 0 'Control Key Down
  keybd_event vbKCode, 0, 0, 0 'Character Key Down
  keybd_event ShiftCtrlAlt, 0, &H2, 0 'Control Key Up
  keybd_event vbKCode, 0, &H2, 0 'Character Key Up
Return
'API call to send only one Character
SendWithoutControl:
  keybd_event vbKCode, 0, 0, 0 'Character Key Down
  keybd_event vbKCode, 0, &H2, 0 'Character Key Up
Return
End Sub
```

