Attribute VB_Name = "Module1"
'Dim Tidy Options
Public tidy As New TidyCOM.TidyObject

Public CleanedSource As String
Public optClean As Boolean
Public optWord2000 As Boolean
Public optBreakBeforeBr As Boolean
Public optCharEncoding
Public optDropEmptyPar As Boolean
Public optFixBackslash As Boolean
Public optShowWarnings As Boolean
Public optIndent As Integer
Public optDoctype As String
Public optQuoteAmpersand As Boolean
Public optIndentSpaces As Integer
Public optWrap As Boolean
Public optDropFontTags As Boolean
Public optEncloseBlockText As Boolean
Public optEncloseText As Boolean
Public optTidyMark As Boolean
Public optTabSize As Integer
Public optDoubleLineBreaks As Boolean
Public optOutputXml As Boolean
Public optCharacterEncoding As Integer
Public optNumericEntities As Boolean

Public Function Clean(strToClean) As String
    ' On Error GoTo errs

    'this string must be present for the Word cleaner to function
    strNS = "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:w=""urn:schemas-microsoft-com:office:word"" xmlns=""http://www.w3.org/TR/REC-html40"">" & vbNewLine

    tidy.Options.Clean = optClean
    tidy.Options.Word2000 = optWord2000
    tidy.Options.BreakBeforeBr = optBreakBeforeBr
    tidy.Options.CharEncoding = optCharacterEncoding
    tidy.Options.DropEmptyParas = optDropEmptyParas
    tidy.Options.FixBackslash = optFixBackslash
    tidy.Options.ShowWarnings = optShowWarnings
    tidy.Options.Indent = optIndent
    tidy.Options.NumericEntities = optNumericEntities
    tidy.Options.Doctype = optDoctype
    tidy.Options.QuoteAmpersand = optQuoteAmpersand
    tidy.Options.IndentSpaces = optIndentSpaces
    tidy.Options.Wrap = optWrap
    tidy.Options.DropFontTags = optDropFontTags
    tidy.Options.EncloseBlockText = optEncloseBlockText
    tidy.Options.EncloseText = optEncloseText
    tidy.Options.TidyMark = optTidyMark
    tidy.Options.TabSize = optTabSize
    tidy.Options.OutputXml = optOutputXml

    If tidy.Options.Word2000 Then
        If InStr(1, strToClean, "xmlns", vbTextCompare) = 0 Then
            NameSpaceAdded = True
            strToClean = strNS & strToClean
        End If
    End If

    Clean = tidy.TidyMemToMem(strToClean)
    Clean = Replace(Clean, strNS, "")

    If CBool(optDoubleLineBreaks) Then
        'strip double linefeeds
        Do Until InStr(Clean, vbNewLine & vbNewLine) = 0
            Clean = Replace(Clean, vbNewLine & vbNewLine, vbNewLine)
        Loop
    End If

    CleanedSource = Clean

    Exit Function
errs:
    MsgBox "An error has occured"
End Function
