'This series of functions examined a particular document for errors or omissions and auto-completed some fields. 
'Most subs were called upon tabbing in or out of a field in the document. Some were called as part of
'an error check. 

Public Sub CbxMedYes()
'If the treatment summary did not list a date for the last medical appointment, check a box to say that
'a medical appointment is needed.
Dim TSumMedDate As FormField
Dim CbxMedical As CheckBox

If ThisDocument.FormFields("TSumMedDate").Result <> "" Then
    ThisDocument.FormFields("CbxMedical").CheckBox.Value = True
End If

End Sub

Public Sub CbxOptYes()
'If the treatment summary did not list a date for the last eye appointment, check a box to say that
'an eye appointment is needed.

Dim TSumOptDate As FormField
Dim CbxOptical As CheckBox

If ThisDocument.FormFields("TSumOptDate").Result <> "" Then
    ThisDocument.FormFields("CbxOptical").CheckBox.Value = True
End If

End Sub

Public Sub CbxDenYes()
'If the treatment summary did not list a date for the last dental appointment, check a box to say that
'a dental appointment is needed.

Dim TSumDentalDate As FormField
Dim CbxDental As CheckBox

If ThisDocument.FormFields("TSumDentalDate").Result <> "" Then
    ThisDocument.FormFields("CbxDental").CheckBox.Value = True
End If

End Sub

Public Sub InsertHelperName()
'Auto-enter the author's name, if indicated. 

Dim HelperName As String
Dim Response As Variant

If ThisDocument.FormFields("ContactPersName").Result = "" Then
Response = MsgBox("Insert your name?", vbYesNo, "Are you the contact person?")
    If Response = vbYes Then
        Call AuthorNameFill(HelperName)
        ThisDocument.FormFields("ContactPersName").Result = HelperName
        If HelperName = "" Then
            MsgBox "Sorry, your name was not in the lookup table."
        End If
    End If
End If



End Sub

Public Sub InsertNameAsAuth()
'Auto-enter the author's name, if indicated. 

Dim DocAuthor As String
Dim Response As Variant

If ThisDocument.FormFields("TSumAuthorName").Result = "" Then
Response = MsgBox("Insert your name?", vbYesNo, "Are you the author?")
    If Response = vbYes Then
        Call AuthorNameFill(DocAuthor)
        ThisDocument.FormFields("TSumAuthorName").Result = DocAuthor
        If DocAuthor = "" Then
            MsgBox "Sorry, your name was not in the lookup table."
        End If
    End If
End If


End Sub


Public Sub InsertToday(FieldChoiceC As Long)
'Inserts today's date into a field when called.
Dim Response As String
Dim TodaysDate As Date

TodaysDate = Format(Now, "mm/dd/yyyy")


    If FieldChoiceC = 5 Then 'called from Discharge Date field
        ActiveDocument.FormFields("CtDischargeDate").Result = TodaysDate
        Exit Sub
    End If
    
    If FieldChoiceC = 6 Then 'called from Date Mailed field
        ActiveDocument.FormFields("TSumDateGivenMailed").Result = TodaysDate
        Exit Sub
    End If


End Sub

Public Sub TsnDateAuto()
'inserts today's date into field if desired.
Dim TransitionFieldNo As Long
Dim CtDischargeDate As FormField
Dim Result As Variant

If ActiveDocument.FormFields("CtDischargeDate").Result = "" Then
    Result = MsgBox("Insert today's date?", vbYesNo)
        If Result = vbYes Then
            TransitionFieldNo = 5
            Call InsertToday(TransitionFieldNo)
        End If
End If


End Sub

Public Sub MailingDateAuto()
'inserts today's date into field if desired.
Dim MailingFldNo As Long
Dim CtDischargeDate As FormField
Dim Result As Variant

If ActiveDocument.FormFields("TSumDateGivenMailed").Result = "" Then
    Result = MsgBox("Insert today's date?", vbYesNo)
        If Result = vbYes Then
            MailingFldNo = 6
            Call InsertToday(MailingFldNo)
        End If
End If


End Sub
