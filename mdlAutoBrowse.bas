Attribute VB_Name = "mdlAutoBrowse"
'Based on source from http://vbpoint.cjb.net

Option Explicit

Public Sub SetInputField(doc As HTMLDocument, Form As Integer, Name As String, Value As String)
'doc = HTMLDocument, can be retrieved from webbrowser --> webbrowser.document
'Form = number of the form (if only one form in the doc --> Form = 0)
'Name = Name of the field you would like to fill
'Value = The new value for the input field called name
'PRE: Legal parameters entered
'POST: Input field with name Name on form Form in document doc will be filled with Value
    
    Dim q As Integer
    
    For q = 0 To doc.Forms(Form).length - 1
        If doc.Forms(Form)(q).Name = Name Then
            doc.Forms(Form)(q).Value = Value
            Exit For
        End If
    Next q
  
End Sub

'Sub to get the contents from a textbox:
Public Function GetInputField(doc As HTMLDocument, Form As Integer, Name As String) As String
  
    Dim q As Integer
    
    For q = 0 To doc.Forms(Form).length - 1
        If doc.Forms(Form)(q).Name = Name Then
            GetInputField = doc.Forms(Form)(q).Value
            Exit For
        End If
    Next q

End Function

'Sub to set a Checkbox:
Public Sub SetCheckBox(doc As HTMLDocument, Form As Integer, Name As String, Value As Boolean)
    
    Dim q As Integer
    
    For q = 0 To doc.Forms(Form).length - 1
        If doc.Forms(Form)(q).Name = Name Then
            doc.Forms(Form)(q).Checked = Value
            Exit For
        End If
    Next q
  
End Sub

'Sub set a radio button:
Public Sub SetRadioButton(doc As HTMLDocument, Form As Integer, Name As String, Name2 As String)
  
    Dim q As Integer
    For q = 0 To doc.Forms(Form).length - 1
        If (doc.Forms(Form)(q).Name = Name) And (doc.Forms(Form)(q).Value = Name2) Then
            doc.Forms(Form)(q).Checked = True
        Exit For
        End If
    Next q

End Sub

'Sub set a combo box:
Public Function SetComboBoxValue(ByVal doc As IHTMLDocument3, Form As Integer, Name As String, Name2 As String)
'****  This one bases it's selection on the Value of the - <option value =
'value'> - Tag.
    Dim q, i As Integer

    For q = 0 To doc.Forms(Form).length - 1
        If (doc.Forms(Form)(q).Name = Name) Then
            For i = 0 To doc.Forms(Form)(q).length - 1
                If doc.Forms(Form)(q).Options(i).Value = Name2 Then
                    doc.Forms(Form)(q).Options(i).Selected = True
                Exit For
                End If
            Next i
        End If
    Next q

End Function

Public Function SetComboValue(ByVal doc As IHTMLDocument3, Form As Integer, Name As String, Name2 As String)
'**** This one bases it's selection on the Value of the Text after the -
'<option>Text - Tag.
    Dim q, i As Integer

    For q = 0 To doc.Forms(Form).length - 1
        If (doc.Forms(Form)(q).Name = Name) Then
            For i = 0 To doc.Forms(Form)(q).length - 1
                If doc.Forms(Form)(q).Options(i).Text = Name2 Then
                    doc.Forms(Form)(q).Options(i).Selected = True
                    Exit For
                End If
            Next
        End If
    Next q

End Function
