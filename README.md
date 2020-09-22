<div align="center">

## text object line info


</div>

### Description

The first function returns usefull information about text box objects. these include :

[Line count] = 0

[Cursor Position] = 1

[Current Line Number] = 2

[Current Line Start] = 3

[Current Line End] = 4

[Current Line Length] = 5

[Current Line Cursor Position] = 6

[Line Start] = 7

[Line End] = 8

[Line Length] = 9

The next function returns the text of a given line of a text box object.
 
### More Info
 
Public Enum LineInfo

[Line count] = 0

[Cursor Position] = 1

[Current Line Number] = 2

[Current Line Start] = 3

[Current Line End] = 4

[Current Line Length] = 5

[Current Line Cursor Position] = 6

[Line Start] = 7

[Line End] = 8

[Line Length] = 9

End Enum

Public Function getLineInfo(txtObj As Object, info As LineInfo, Optional lineNumber As Long) As Long

Public Function GetLineText(txtObj As Object, lineNumber As Long) As String

'// If lineNumber = 0 then current line's text is given


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Damien McGivern](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/damien-mcgivern.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/damien-mcgivern-text-object-line-info__1-3323/archive/master.zip)

### API Declarations

```
Public Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const EM_GETSEL As Long = &HB0
Public Const EM_SETSEL As Long = &HB1
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_LINEINDEX As Long = &HBB
Public Const EM_LINELENGTH As Long = &HC1
Public Const EM_LINEFROMCHAR As Long = &HC9
Public Const EM_SCROLLCARET As Long = &HB7
Public Const WM_SETREDRAW As Long = &HB
```


### Source Code

```
'Author : Damien McGivern
'E-Mail : D_McGivern@Yahoo.Com
'Date : 30 Aug 1999
Option Explicit
Public Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const EM_GETSEL As Long = &HB0
Public Const EM_SETSEL As Long = &HB1
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_LINEINDEX As Long = &HBB
Public Const EM_LINELENGTH As Long = &HC1
Public Const EM_LINEFROMCHAR As Long = &HC9
Public Const EM_SCROLLCARET As Long = &HB7
Public Const WM_SETREDRAW As Long = &HB
Public Enum LineInfo
 [Line count] = 0
 [Cursor Position] = 1
 [Current Line Number] = 2
 [Current Line Start] = 3
 [Current Line End] = 4
 [Current Line Length] = 5
 [Current Line Cursor Position] = 6
 [Line Start] = 7
 [Line End] = 8
 [Line Length] = 9
End Enum
Public Function getLineInfo(txtObj As Object, info As LineInfo, Optional lineNumber As Long) As Long
 Dim cursorPoint As Long
 '//Record where the cursor is
 cursorPoint = txtObj.SelStart
 Select Case info
  Case Is = 0 ' = "lineCount"
   getLineInfo = SendMessageLong(txtObj.hWnd, EM_GETLINECOUNT, 0, 0&)
  Case Is = 1 ' = "cursorPosition"
   getLineInfo = (SendMessageLong(txtObj.hWnd, EM_GETSEL, 0, 0&) \ &H10000) + 1
  Case Is = 2 ' = "currentLineNumber"
   getLineInfo = (SendMessageLong(txtObj.hWnd, EM_LINEFROMCHAR, -1, 0&)) + 1
  Case Is = 3 ' = "currentLineStart"
   getLineInfo = SendMessageLong(txtObj.hWnd, EM_LINEINDEX, -1, 0&) + 1
  Case Is = 4 ' = "currentLineEnd"
   getLineInfo = SendMessageLong(txtObj.hWnd, EM_LINEINDEX, -1, 0&) + 1 + SendMessageLong(txtObj.hWnd, EM_LINELENGTH, -1, 0&)
  Case Is = 5 ' = "currentLineLength"
   getLineInfo = SendMessageLong(txtObj.hWnd, EM_LINELENGTH, -1, 0&)
  Case Is = 6 ' = "currentLineCursorPosition"
   getLineInfo = (SendMessageLong(txtObj.hWnd, EM_GETSEL, 0, 0&) \ &H10000) + 1 - SendMessageLong(txtObj.hWnd, EM_LINEINDEX, getLineInfo(txtObj, [Current Line Number]) - 1, 0&)
  Case Is = 7 ' = "lineStart"
   getLineInfo = (SendMessageLong(txtObj.hWnd, EM_LINEINDEX, (lineNumber - 1), 0&)) + 1
  Case Is = 8 ' = "lineEnd"
   getLineInfo = SendMessageLong(txtObj.hWnd, EM_LINEINDEX, (lineNumber - 1), 0&) + 1 + SendMessageLong(txtObj.hWnd, EM_LINELENGTH, (lineNumber - 1), 0&)
  Case Is = 9 ' = "lineLength"
   getLineInfo = (SendMessageLong(txtObj.hWnd, EM_LINEINDEX, lineNumber, 0&)) + 1 - (SendMessageLong(txtObj.hWnd, EM_LINEINDEX, (lineNumber - 1), 0&)) - 3
 End Select
End Function
Public Function GetLineText(txtObj As Object, lineNumber As Long) As String
'// If lineNumber = 0 then current line's text is given
 If lineNumber = 0 Then lineNumber = getLineInfo(txtObj, [Current Line Number])
 '// Select text
 Call SendMessageLong(txtObj.hWnd, EM_SETSEL, ((getLineInfo(txtObj, [Line Start], lineNumber)) - 1), ((getLineInfo(txtObj, [Line Start], lineNumber + 1)) - 1))
 GetLineText = txtObj.SelText
End Function
```

