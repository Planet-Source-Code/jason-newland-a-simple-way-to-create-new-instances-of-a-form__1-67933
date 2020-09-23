Attribute VB_Name = "modNewInstance"
'basic module for creating multiple instances of a form

Public sNewForm() As frmClient
Public sNewFormMax As Long

Public Sub CreateNewInstance()
    On Error Resume Next
    Dim i As Long
    'this sub creates a new instance of a form
    'easily and effectively
    '
    'first check that the maximum count to redim the
    'form array is not 0
    If sNewFormMax = 0 Then
        'increase the array initally to 5
        sNewFormMax = 5
        ReDim Preserve sNewForm(0 To sNewFormMax) As frmClient
    End If
    '
    'ok, next we check if we need to redim the
    'array if it is going to be greater than sNewFormMax
    If GetFormCount > sNewFormMax Then
        'increase array by 3, can be any number you want
        'but we are trying to minimize resource head-room
        Debug.Print "increasing array by 3.."
        sNewFormMax = sNewFormMax + 10
        ReDim Preserve sNewForm(0 To sNewFormMax) As frmClient
    End If
    '
    'now search for an empty (Nothing) array to set the
    'new form to
    For i = LBound(sNewForm) To UBound(sNewForm)
        'this will check for an empty array to fill rather
        'than create a new array and end up running out of
        'arrays, especially when the form is closed
        If sNewForm(i) Is Nothing Then
            Debug.Print "We are using form array index: " & i
            Set sNewForm(i) = New frmClient
            '
            With sNewForm(i)
                .Caption = "Hello I'm Form Array # " & i
                'you could do other things here, like
                'set controls on the form to a specific
                'setting, etc
                '
                'show the damn thing!
                .Show vbModeless, frmMain
                Exit For
            End With
        End If
    Next i
End Sub

Private Function GetFormCount() As Long
    On Error Resume Next
    'function to the current number of frmClient forms
    'will always be one higher than actual sNewFormMax
    '
    'it can always be modified to except a string input
    'so it can search for other forms for other parts of
    'your program, like private message forms, etc
    Dim i As Long
    Dim total As Long
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "frmClient" Then total = total + 1
    Next i
    GetFormCount = total
End Function
