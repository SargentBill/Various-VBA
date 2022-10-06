Public Sub CallLogsToTable3()
' ConvertCallLogToTable Macro
' Takes the Android call log and makes a formatted table out of the data.
' Future development:
'    add error check in case i run it on Mileage.
'    add a way to combine call log w mileage to make daily notes and times.
' Updates:
'

Dim RowLoopVar As Long
Dim LastRowColF As Long
Dim IntFormat As Long
Dim ExtrDateString As String 'string from cell orig date/time
Dim ExtrDate As String 'Extracted Date String
Dim DateAsDate As Date 'Date string converted to Date
Dim TheWeekday As Variant 'vbNumber of the day of week
Dim ExtrStartTime As String 'Extracted Time String
Dim ExtrStartTimeAsTime As Date 'Extracted Time as date
Dim DurTime As Variant
Dim EndTime0 As Date
Dim EndTime1 As Date
Dim ExtrStartTimeWithTimeValue As Variant
Dim StartedDurTime As String 'Duration unformatted and unanalyzed
Dim HourDigits As Integer 'hour digit of duration
Dim MinDigits As Integer 'minute digits of duration
Dim SecDigits As Integer 'seconds digits of duration

LastRowColA = Range("A65536").End(xlUp).Row 'Last row in Phone Number column
LastRowColB = Range("B65536").End(xlUp).Row 'Last row in Contact column.
LastRowColC = Range("C65536").End(xlUp).Row 'Last row in In/Out Type column.
LastRowColD = Range("D65536").End(xlUp).Row 'Last row in orig Date/Time column.
LastRowColE = Range("E65536").End(xlUp).Row 'Last row in final Day/initial Duration column.
LastRowColF = Range("F65536").End(xlUp).Row 'Last row in final Date/initial Subject column.
LastRowColG = Range("G65536").End(xlUp).Row 'Last row in final Start time/intial Note column.
LastRowColH = Range("H65536").End(xlUp).Row 'Last row in final stop time/ intial Next Event column.
LastRowColI = Range("I65536").End(xlUp).Row 'Last row in final duration/initial Company column.
LastRowColJ = Range("J65536").End(xlUp).Row 'Last row in final Subject/initial Matter column.
LastRowColK = Range("K65536").End(xlUp).Row 'Last row in final Note column.
LastRowColL = Range("L65536").End(xlUp).Row 'Last row in final Next Event column.
LastRowColM = Range("M65536").End(xlUp).Row 'Last row in final Company column.
LastRowColN = Range("N65536").End(xlUp).Row 'Last row in final Matter column.
LastRowColO = Range("O65536").End(xlUp).Row 'Last row in Blank column.

'If ActiveSheet.Range("G1").Value = "Note" Then
'    MsgBox "You accidentally ran this on CallNotes instead of CallLog."
'    Exit Sub
'End If

    'copy all the stuff over
        'LastRowColA is used b/c a phone # is required -no blanks
        'Copy the stuff on the right side of sheet first, erase original data, copy on top of old.
    'Copy Matter (From J to N)
    ActiveSheet.Range("J1:J" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("N1")
    'copy Company (from I to M)
    ActiveSheet.Range("I1:I" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("M1")
    'Copy next event col (from H to L)
    ActiveSheet.Range("H1:H" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("L1")
    'copy Note Column over to the right (from G to K)
    ActiveSheet.Range("G1:G" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("K1")
    
    'clear contents and formatting of the stuff copied so far
    With ActiveSheet.Range("G1:J" & LastRowColA)
        .ClearContents 'contents
        .Clear 'formulas and formatting
    End With
    
    'copy Subject Column over to the right (From F to J)
    ActiveSheet.Range("F1:F" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("J1")
    
    'copy Duration over (From E to I)
    ActiveSheet.Range("E1:E" & LastRowColA).Copy _
    Destination:=ActiveSheet.Range("I1")
    
    'clear source of what was just copied.
    With ActiveSheet.Range("E1:F" & LastRowColA)
        .ClearContents 'contents
        .Clear 'formulas and formatting
    End With
    'Name the new column headers
    ActiveSheet.Range("E1").Value = "Day"
    ActiveSheet.Range("F1").Value = "Date"
    ActiveSheet.Range("G1").Value = "Start Time"
    ActiveSheet.Range("H1").Value = "End Time"
    ActiveSheet.Range("I1").Value = "Duration"
    ActiveSheet.Range("J1").Value = "Subject"
    ActiveSheet.Range("K1").Value = "Note"
    ActiveSheet.Range("L1").Value = "Next Event"
    ActiveSheet.Range("M1").Value = "Company"
    ActiveSheet.Range("N1").Value = "Matter"
    ActiveSheet.Range("O1").Value = "Overlap"
    ActiveSheet.Range("P1").Value = "Blank1"
    ActiveSheet.Range("Q1").Value = "Blank2"
    ActiveSheet.Range("R1").Value = "Blank3"
    
    'match formatting of new headers
    Range("D1").Select
    IntFormat = Selection.Interior.ColorIndex
    With Range("E1:R1")
        .Select
        .Interior.ColorIndex = IntFormat
    End With

'center and format columns
ActiveSheet.Range("C1:I" & LastRowColA).HorizontalAlignment = xlCenter
Columns("A:A").ColumnWidth = 12.5
Columns("B:B").ColumnWidth = 14.7
Columns("C:C").ColumnWidth = 6.29
Columns("D:D").ColumnWidth = 0
Columns("E:E").ColumnWidth = 11
Columns("F:F").ColumnWidth = 11.5
Columns("G:G").ColumnWidth = 8.7
Columns("H:H").ColumnWidth = 8.7
Columns("K:K").ColumnWidth = 30

For i = 2 To LastRowColA
    ExtrDateString = Range("D" & i).Value
    ExtrDate = Left(ExtrDateString, 10) 'pulls out the date chars e.g. 10-23-2012
    DateAsDate = CDate(ExtrDate)
    TheWeekday = Weekday(DateAsDate)
    Select Case TheWeekday
        Case 2
            ActiveSheet.Range("E" & i).Value = "Monday"
        Case 3
            ActiveSheet.Range("E" & i).Value = "Tuesday"
        Case 4
            ActiveSheet.Range("E" & i).Value = "Wednesday"
        Case 5
            ActiveSheet.Range("E" & i).Value = "Thursday"
        Case 6
            ActiveSheet.Range("E" & i).Value = "Friday"
        Case 1
            ActiveSheet.Range("E" & i).Value = "Sunday"
        Case 2
            ActiveSheet.Range("E" & i).Value = "Saturday"
        Case Else
            ActiveSheet.Range("E" & i).Value = ""
    End Select
    
ActiveSheet.Range("F" & i).Value = DateAsDate

'G is StartTime

ExtrStartTime = Right(ExtrDateString, 8) 'extracts time from D col string
ExtrStartTimeWithTimeValue = TimeValue(ExtrStartTime) 'start time ver II
StartDayAndTime = DateAsDate + ExtrStartTimeWithTimeValue
ActiveSheet.Range("G" & i).Value = StartDayAndTime 'inserts time to the column

'ActiveSheet.Range("G" & i).Value = ExtrStartTimeWithTimeValue 'inserts time to the column
StartedDurTime = ActiveSheet.Range("I" & i).Value 'VBA sees this as hh:mm, so extract the numbers

If Len(StartedDurTime) = 5 Then
    SecDigits = Right(StartedDurTime, 2)
    MinDigits = Left(StartedDurTime, 2)
    HourDigits = 0
End If

If Len(StartedDurTime) > 5 Then
    HourDigits = Left(StartedDurTime, 2)
    SecDigits = Right(StartedDurTime, 2)
    MinDigits = Mid(StartedDurTime, 4, 2)
    HourDigits = CInt(HourDigits)
End If


SecDigits = CInt(SecDigits) 'converts text to integer
MinDigits = CInt(MinDigits)

DurTime = TimeSerial(HourDigits, MinDigits, SecDigits)

EndTime0 = DateAsDate + ExtrStartTimeWithTimeValue
EndTime1 = EndTime0 + DurTime
ActiveSheet.Range("H" & i).Value = EndTime1

Next i

' set the range to military time:

ActiveSheet.Range("G2:H" & LastRowColA).NumberFormat = "hh:mm"




'Make a table with the data
    Range("A1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$R$500"), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
    With ActiveCell.Characters(Start:=1, Length:=9).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
Exit Sub


'Freeze the top row, rest can scroll.
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True


End Sub



Public Sub CombineLogs()

Dim LastRowDestination As Long
Dim LastRowPhone As Long
Dim wcount As Long

    LastRowPhone = ActiveWindow.ActiveSheet.Range("A65536").End(xlUp).Row
    ActiveWindow.ActiveSheet.Range("A2:R" & LastRowPhone).Copy
    ActiveWindow.ActivateNext
    LastRowDestination = ActiveWindow.ActiveSheet.Range("A65536").End(xlUp).Row
    LastRowDestination = LastRowDestination + 1
    'ActiveWindow.ActiveSheet.Range("A" & LastRowDestination).Select
    ActiveSheet.Paste Destination:=Worksheets(1).Range("A" & LastRowDestination)
    
Call CallLogsToTable3
    
End Sub
