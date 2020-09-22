Attribute VB_Name = "Colorize"
'########################################################
'for the most part this module came from PlanetSourceCode
'thanks to the author (whoever that is)
'########################################################


Option Explicit
Option Compare Binary
Private regEx           As New VBScript_RegExp_55.RegExp
Private replStr         As String

Private arrColorize()   As String
Private lngNoColors     As Long
'Private objOutput       As New cAppendString
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Function SysAllocStringByteLen& Lib "oleaut32" (ByVal olestr&, ByVal BLen&)

Private plngStringLen   As Long
Private plngBufferLen   As Long
Private pstrBuffer      As String
Private strColorizedRTFHeader As String

Function DoColorize(strHTML As String) As String
    Clear
    fcnColorize strHTML, DoColorize
End Function

Public Function fcnColorize(strToColorize As String, ByRef strColorizedRTF As String) As Boolean

    arrColorize = fcnGetColorizeArray()
    '***  build the rtf header
    strColorizedRTFHeader = "{\rtf1\ansi\deff0\deflang2057{\fonttbl{\f0\fmodern\fprq1\fcharset0 Arial;}}"
    strColorizedRTFHeader = strColorizedRTFHeader & "{\colortbl\red0\green0\blue0;"


    InitRegEX "( |\n|\r|\t)(\w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+)", "$1\cf3 $2\cf2 $3\cf4 "



    '********************************

    fcnColorize = False
    Dim strOverLap       As String
    Dim lngFile          As Long
    Dim lngLen           As Long
    Dim lngLOF           As Long
    Dim lngCounter       As Long
    Dim lngNoBlocks      As Long
    Dim lngExtra         As Long
    Dim lngMain          As Long
    Dim strBlockIn       As String
    Dim lngDataPos       As Long
    Dim lngCount         As Long
    Dim strTemp          As String
    Dim strTemp2         As String
    Dim blnDoString      As Boolean

    '***  add the rtf colors to the header
    For lngCount = LBound(arrColorize, 2) To UBound(arrColorize, 2)
        strColorizedRTFHeader = strColorizedRTFHeader & "\red" & CStr(arrColorize(7, lngCount)) & "\green" & CStr(arrColorize(8, lngCount)) & "\blue" & CStr(arrColorize(9, lngCount)) & ";"
        '***  add spaces to the color defs
        arrColorize(2, lngCount) = arrColorize(2, lngCount) & Chr$(32&)
        arrColorize(3, lngCount) = arrColorize(3, lngCount) & Chr$(32&)
    Next lngCount
    '***  add some extra tabs to the header
    strColorizedRTFHeader = strColorizedRTFHeader & "}\deflang2057\pard\tx120\tx240\tx360\tx480\tx600\tx720\tx840\plain\f0\fs20\cf0 "

    '***  set the number of colors
    lngNoColors = UBound(arrColorize, 2)

    Append strColorizedRTFHeader

    '***  this is the main loop
    strTemp = strToColorize

    strTemp2 = fcnEscapeRTF(strTemp)

    strBlockIn = strTemp2

    fcnColorizeBlock strBlockIn, 0, ""



    Append ("}")
    strColorizedRTF = Value
    fcnColorize = True



End Function

Private Function fcnColorizeBlock(strBlock As String, Optional lngState As Long, Optional strOverLap As String, Optional blnRoot = True) As Boolean

    '***  this function replaces the text in the html file with rtf color codes.

    Dim strColEnd        As String
    Dim lngCounter       As Long
    Dim lngPosS          As Long
    Dim lngPosE          As Long
    Dim lngFoundS        As Long
    Dim lngFoundE        As Long
    Dim strLeft          As String
    Dim strMid           As String
    Dim strRight         As String
    Dim strTemp          As String

    '***  all was searched, nothing found. Add to buffer unmodified
    If lngState > lngNoColors Or LenB(strBlock) = 0 Then
        '***  checked all
        Append strBlock

        Exit Function
    End If
    lngPosS = 1
    lngCounter = lngState

    lngFoundS = InStr(lngPosS, strBlock, arrColorize(0, lngCounter), arrColorize(6, lngCounter))

    If lngFoundS <> 0 And Not (blnRoot = True And arrColorize(5, lngCounter) = "False") Then
        '***  found a start, now search the left part for other tags

        '***  process left string
        strLeft = Left$(strBlock, lngFoundS - 1)
        lngState = lngCounter + 1
        fcnColorizeBlock strLeft, lngState, "", blnRoot


        If LenB(strOverLap) = 0 Then
            '***  search mid part

            strColEnd = arrColorize(1, lngCounter)
            lngPosE = lngFoundS + 1
            lngFoundE = InStr(lngPosE, strBlock, strColEnd, arrColorize(6, lngCounter))

            If lngFoundE <> 0 Then

                strMid = Mid$(strBlock, lngFoundS, lngFoundE - lngFoundS + Len(strColEnd))

                If arrColorize(4, lngCounter) = "True" Then
                    '***  regular expression to colorize the inside of the html tag (4)
                    strMid = ReplaceText(strMid)

                End If

                '***  found something, add colors and string to buffer
                Append arrColorize(2, lngCounter)
                Append strMid
                Append arrColorize(3&, lngCounter)


                strRight = Mid$(strBlock, lngFoundE + Len(strColEnd))
                lngState = lngCounter


                fcnColorizeBlock strRight, lngState, strOverLap, blnRoot

            ElseIf blnRoot Then                            'And lngFoundE = 0

                strOverLap = Mid$(strBlock, lngFoundS)
                fcnColorizeBlock strOverLap, lngState, "", blnRoot


                lngState = lngCounter

            End If                                         'lngFoundE <> 0

        Else                                               'NOT LenB(strOverLap) = 0

            strOverLap = vbNullString
            fcnColorizeBlock strBlock, lngState, strOverLap, blnRoot
            strBlock = strLeft

        End If                                             'LenB(strOverLap) = 0

    Else

        '***  nothing found search next delimiter
        lngState = lngCounter + 1
        fcnColorizeBlock strBlock, lngState, strOverLap, blnRoot

    End If

    fcnColorizeBlock = lngState
End Function

Public Function fcnGetColorizeArray() As String()
    '***  the array in this function is customizable
    '***  it contains a number of start - end definitions and colors
    '***  items are replaced ordered by index

    'start     start text
    'End       end text
    'startrtf  color code for this item
    'endrtf    color code after this item
    'fill      true if the inside of this item must be colored
    'root      true if this item cannot be inside other items. for example a remark "<!-- ... -->"
    'compare   compare mode used to find start and end in the text. vbTextCompare or vbBinaryCompare
    'red       single byte color value for this item
    'green     single byte color value for this item
    'blue      single byte color value for this item

    '***  new definitions can be made to colorize other file types than html
    '***  feel free to change/extend the array
    '***  also change the redim!

    Dim arrCol()         As String

    ReDim arrCol(10, 3)

    arrCol(0, 0) = "<!--"                                  'start of the tag
    arrCol(1, 0) = "-->"                                   'end of the tag
    arrCol(2, 0) = "\cf1"                                  'start rtf color code
    arrCol(3, 0) = "\cf0"                                  'end rtf color code
    arrCol(4, 0) = "False"                                 'search internally?
    arrCol(5, 0) = "True"                                  'root level
    arrCol(6, 0) = "0"                                     '0=binary compare, 1=textcompare
    arrCol(7, 0) = "255"                                   'red value in \cf1
    arrCol(8, 0) = "25"                                    'green
    arrCol(9, 0) = "0"                                     'blue
    arrCol(10, 0) = ""                                     'reserved for regular expression

    arrCol(0, 1) = "<script"
    arrCol(1, 1) = "</script>"
    arrCol(2, 1) = "\cf2"
    arrCol(3, 1) = "\cf0"
    arrCol(4, 1) = "False"
    arrCol(5, 1) = "True"
    arrCol(6, 1) = "1"
    arrCol(7, 1) = "40"
    arrCol(8, 1) = "40"
    arrCol(9, 1) = "180"
    arrCol(10, 1) = ""

    arrCol(0, 2) = "<%"
    arrCol(1, 2) = "%>"
    arrCol(2, 2) = "\cf3"
    arrCol(3, 2) = "\cf0"
    arrCol(4, 2) = "False"
    arrCol(5, 2) = "True"
    arrCol(6, 2) = "0"
    arrCol(7, 2) = "0"
    arrCol(8, 2) = "120"
    arrCol(9, 2) = "25"
    arrCol(10, 2) = ""

    arrCol(0, 3) = "<"
    arrCol(1, 3) = ">"
    arrCol(2, 3) = "\cf4"
    arrCol(3, 3) = "\cf0"
    arrCol(4, 3) = "True"
    arrCol(5, 3) = "True"
    arrCol(6, 3) = "0"
    arrCol(7, 3) = "20"
    arrCol(8, 3) = "20"
    arrCol(9, 3) = "255"
    arrCol(10, 3) = ""

    fcnGetColorizeArray = arrCol
End Function

Private Function fcnEscapeRTF(strText As String) As String

    Dim strText2         As String
    '***  escape the rtf special chars
    strText2 = Replace(strText, "\", "\\", , , vbBinaryCompare)
    strText2 = Replace(strText2, "{", "\{", , , vbBinaryCompare)
    strText2 = Replace(strText2, "}", "\}", , , vbBinaryCompare)
    '***  put \par after each vbLf. Do not use vbCrLf, because a vbCrLf might be broken
    '***  in two by a block border.
    fcnEscapeRTF = Replace(strText2, vbLf, vbLf & "\par ", , , vbBinaryCompare)
End Function
Public Function ReplaceText(textStr As String) As String
    ReplaceText = regEx.Replace(textStr, replStr)          ' Make replacement.
End Function

Public Sub InitRegEX(patrn As String, newrepl As String)
    regEx.Pattern = patrn                                  ' Set pattern.
    regEx.IgnoreCase = False                               ' Make case insensitive.
    regEx.Global = True
    replStr = newrepl
End Sub
Public Sub Clear()
    '***  do not clear the buffer to save allocation time
    '***  if you use the function multiple times
    plngStringLen = 0&

    plngBufferLen = 0&                                     'clear the buffer
    pstrBuffer = vbNullString                              'clear the buffer
End Sub


Public Sub Append(Text As String)
    Dim lngText          As Long
    Dim strTemp          As String
    Dim lngVPointr       As Long

    lngText = Len(Text)

    If lngText > 0 Then
        If (plngStringLen + lngText) > plngBufferLen Then
            plngBufferLen = (plngStringLen + lngText) * 2&
            strTemp = AllocString04(plngBufferLen)

            '***  copymemory might be faster than this
            Mid$(strTemp, 1&) = pstrBuffer

            '***  Alternate pstrBuffer = strTemp
            '***  switch pointers instead of slow =
            lngVPointr = StrPtr(pstrBuffer)
            RtlMoveMemory ByVal VarPtr(pstrBuffer), ByVal VarPtr(strTemp), 4&
            RtlMoveMemory ByVal VarPtr(strTemp), lngVPointr, 4&

            'Debug.Print "plngBufferLen: " & plngBufferLen
        End If

        Mid$(pstrBuffer, plngStringLen + 1&) = Text
        plngStringLen = plngStringLen + lngText
    End If


End Sub

Public Function Value() As String
    Value = Left$(pstrBuffer, plngStringLen)
End Function

Private Function AllocString04(ByVal lSize As Long) As String
    ' http://www.xbeat.net/vbspeed/
    ' by Jory, jory@joryanick.com, 20011023
    RtlMoveMemory ByVal VarPtr(AllocString04), SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function


