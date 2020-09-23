Attribute VB_Name = "mdlParseCode"
'Based on code written by Jameson Schriber

Option Explicit

Public Sub ParseCode(Code As String)
    
    Dim UserCodeArray() As String
    Dim ArgumentsArray() As String
    Dim UserCode As String
    Dim LineFunction As String
    Dim LineArguments As String
    Dim UserLines As Integer
    Dim LeftParenthesisPos As Integer
    Dim RightParenthesisPos  As Integer
    Dim LineLength As Integer
    Dim i, n As Integer
    'Dim doc as a MS HTML Object Library
    Dim doc As HTMLDocument
       
    'Set doc as browser doc
    Set doc = frmBrowser.wbBrowser.Document
    
    'Preparing the code
    UserCode = BackslashEscape(Code)

    'Splitting the code, an array element for each command and its arguments
    UserCodeArray() = Split(UserCode, ";")

    'Finding number of commands
    UserLines = UBound(UserCodeArray)
    
    ReDim Preserve UserCodeArray(UserLines - 1)
    For i = 0 To UserLines - 1
        'The guts of the parsing/processing routine, pretty self-explanatory
        LineLength = Len(UserCodeArray(i))
        LeftParenthesisPos = InStr(UserCodeArray(i), "(")
        RightParenthesisPos = InStrRev(UserCodeArray(i), ")")
        LineFunction = Left(UserCodeArray(i), LeftParenthesisPos - 1)
        LineArguments = Mid(UserCodeArray(i), LeftParenthesisPos + 1, RightParenthesisPos - (LeftParenthesisPos + 1))
        
        'THE COMMAND SELECT-CASE BLOCK
        'Each command has it's own case statement
        'Arguments are accessible through LineArguments string
        Select Case UCase(LineFunction)
        
        Case "BROWSE"
            'Split arguments
            ArgumentsArray = Split(LineArguments, ",")
            'Check for correct arguments
            If UBound(ArgumentsArray) <> 0 Then
                MsgBox "Command: '" & LineFunction & "' needs 1 arguments", vbOKOnly, "Syntax Error"
            Else
                'Convert characters back and remove quotes
                ArgumentsArray(0) = ConvertEscapeCharsBack(ArgumentsArray(0))
                ArgumentsArray(0) = Replace(ArgumentsArray(0), """", "")
            
                'Browse to URL
                frmBrowser.wbBrowser.Navigate ArgumentsArray(0)
            End If
        Case "SETINPUTFIELD"
            'Split arguments
            ArgumentsArray = Split(LineArguments, ",")
            'Check for correct arguments
            If UBound(ArgumentsArray) <> 2 Then
                MsgBox "Command: '" & LineFunction & "' needs 3 arguments", vbOKOnly, "Syntax Error"
            Else
                'Convert characters back and remove quotes
                For n = LBound(ArgumentsArray) To UBound(ArgumentsArray)
                    ArgumentsArray(n) = ConvertEscapeCharsBack(ArgumentsArray(n))
                    ArgumentsArray(n) = Replace(ArgumentsArray(n), """", "")
                Next n
                'Set input fields
                SetInputField doc, CInt(ArgumentsArray(0)), ArgumentsArray(1), ArgumentsArray(2)
            End If
        Case "SUBMIT"
            'Spilt arguments
            ArgumentsArray = Split(LineArguments, ",")
            'Check for correct arguments
            If UBound(ArgumentsArray) <> 0 Then
                MsgBox "Command: '" & LineFunction & "' needs 1 arguments", vbOKOnly, "Syntax Error"
            Else
                'Submit the form (same result as click the search button)
                doc.Forms(0).submit
            End If
        Case "MSG"
            'Split arguments
            ArgumentsArray = Split(LineArguments, ",")
            'Check for correct arguments
            If UBound(ArgumentsArray) <> 1 Then
                MsgBox "Command: '" & LineFunction & "' needs 2 arguments", vbOKOnly, "Syntax Error"
            Else
                'Send message box
                MsgBox ArgumentsArray(0), vbOKOnly, ArgumentsArray(1)
            End If
        Case "PRINT"
            'Split arguments
            ArgumentsArray = Split(LineArguments, ",")
            'Check for correct arguments
            If UBound(ArgumentsArray) <> 0 Then
                MsgBox "Command: '" & LineFunction & "' needs 1 arguments", vbOKOnly, "Syntax Error"
            Else
                'Print web page, true argument show print dialog, false do not show dialog
                frmBrowser.PrintWebPage CBool(ArgumentsArray(0))
            End If
        Case Else
            'Message box with error
            MsgBox "Command: '" & LineFunction & "' not a valid command", vbOKOnly, "Script Syntax Error"
        'Case Add more commands here
        End Select
    Next

End Sub

Public Function BackslashEscape(Code As String) As String

    Dim buffer As String
    
    'This function is also a good place to kill all tabs and newlines before we actually start processing the code
    buffer = Replace(Code, vbCrLf, "")
    buffer = Replace(buffer, vbTab, "")
    'Replace all backslash escape characters so that we can process the agruments and convert them back later on
    buffer = Replace(buffer, "\n", Chr(0) & "Newline")
    buffer = Replace(buffer, "\t", Chr(0) & "Tab")
    buffer = Replace(buffer, "\\", Chr(0) & "Backslash")
    buffer = Replace(buffer, "\""", Chr(0) & "Quote")
    buffer = Replace(buffer, "\;", Chr(0) & "Colon")
    BackslashEscape = Replace(buffer, "\,", Chr(0) & "Comma")
    
End Function

Public Function ConvertEscapeCharsBack(Code As String) As String

    Dim buffer As String

    'Convert the "intermediate" escape chars to the actual characters
    buffer = Replace(Code, Chr(0) & "Newline", vbCrLf)
    buffer = Replace(buffer, Chr(0) & "Tab", vbTab)
    buffer = Replace(buffer, Chr(0) & "Backslash", "\")
    buffer = Replace(buffer, Chr(0) & "Quote", """")
    buffer = Replace(buffer, Chr(0) & "Colon", ";")
    buffer = Replace(buffer, Chr(0) & "Comma", ",")
    ConvertEscapeCharsBack = Trim(buffer)
    
End Function
