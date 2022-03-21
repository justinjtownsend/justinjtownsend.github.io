Sub CustomEventReader()

'***************************************************************************
'Purpose: Identify the unique events from a calendar file / stream *.ics
'Inputs:  icsLoc - *.ics file
'Outputs: List of distinct events
'
'Author: Justin Townsend
'***************************************************************************
  
  'Variable declaration
  Dim icsLoc As String, line As String, evt_line() As String, eStr As String, dtStr As String
  Dim objStream As stream
  Dim item As Variant
  Dim u As Variant
  
  icsLoc = Application.GetOpenFilename("Calendar Files (*.ics),*.ics")
  
  If icsLoc = "False" Then Exit Sub

On Error GoTo ERROR_HANDLER

  Set objStream = CreateObject("ADODB.Stream")
    
  objStream.Charset = "utf-8"
  objStream.Open
  objStream.Type = adTypeText
  objStream.LoadFromFile (icsLoc)
  
  'Clear the active sheet / delete the existing table
  ActiveSheet.Cells.Clear

  'Set counter to track worksheet used range
  u = 0
  
  'Add a header for the events and the required columns
  ActiveSheet.Cells(1, 1).Value = "begin-event-type"
  ActiveSheet.Cells(1, 2).Value = "dtstamp"
  ActiveSheet.Cells(1, 3).Value = "dtstart"
  ActiveSheet.Cells(1, 4).Value = "dtend"
  ActiveSheet.Cells(1, 5).Value = "summary"
  ActiveSheet.Cells(1, 6).Value = "uid"
  ActiveSheet.Cells(1, 7).Value = "description"
  ActiveSheet.Cells(1, 8).Value = "organizer"
  ActiveSheet.Cells(1, 9).Value = "created"
  ActiveSheet.Cells(1, 10).Value = "last-modified"
  ActiveSheet.Cells(1, 11).Value = "attendee"
  ActiveSheet.Cells(1, 12).Value = "sequence"
  ActiveSheet.Cells(1, 13).Value = "subcalendar-type"
  ActiveSheet.Cells(1, 14).Value = "status"
  ActiveSheet.Cells(1, 15).Value = "end-event-type"
  
  Do Until objStream.EOS

    line = vbNullString
    
    If (ReadEvent(objStream, line)) Then
    
      u = ActiveSheet.UsedRange.Rows.Count + 1
    
      'Process the event
      evt_line() = Split(line, vbCrLf)
      
      For Each item In evt_line()
      
        dStr = vbNullString
        eStr = vbNullString

        Select Case True
           
          'Identify the beginning of the event, BEGIN:VEVENT
          Case Left(item, 12) = "BEGIN:VEVENT"
            eStr = Split(item, ":")(1)
            Cells(u, 1) = eStr
          'Identify the date / timestamp of the event, DTSTAMP
          Case Left(item, 7) = "DTSTAMP"
            Cells(u, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            dtStr = Split(item, ":")(1)
            Cells(u, 2) = Format(ParseDateZ(dtStr), "yyyy-mm-dd hh:mm:ss")
          'Identify the start date for the event, DTSTART
          Case Left(item, 7) = "DTSTART"
            Cells(u, 3).NumberFormat = "yyyy-mm-dd"
            dtStr = Split(item, ":")(1)
            Cells(u, 3) = Format(ParseDate(dtStr), "yyyy-mm-dd")
          'Identify the end date for the event, DTEND
          Case Left(item, 5) = "DTEND"
            Cells(u, 4).NumberFormat = "yyyy-mm-dd"
            dtStr = Split(item, ":")(1)
            Cells(u, 4) = Format(ParseDate(dtStr), "yyyy-mm-dd")
          'Identify the event summary, SUMMARY
          Case Left(item, 7) = "SUMMARY"
            eStr = Split(item, ":")(1)
            Cells(u, 5) = eStr
          'Identify the unique identifier for the event, UID
          Case Left(item, 3) = "UID"
            eStr = Split(item, ":")(1)
            Cells(u, 6) = eStr
          'Identify the description, DESCRIPTION
          Case Left(item, 11) = "DESCRIPTION"
            eStr = Split(item, ":")(1)
            Cells(u, 7) = eStr
          'Identify the organizer, ORGANIZER
          Case Left(item, 9) = "ORGANIZER"
            eStr = Split(item, ";")(3)
            eStr = Mid(eStr, InStr(eStr, "mailto:"), Len(eStr))
            Cells(u, 8) = eStr
          'Identify the date created, CREATED
          Case Left(item, 7) = "CREATED"
            Cells(u, 9).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            dtStr = Split(item, ":")(1)
            Cells(u, 9) = Format(ParseDateZ(dtStr), "yyyy-mm-dd hh:mm:ss")
          'Identify the date modified, LAST-MODIFIED
          Case Left(item, 13) = "LAST-MODIFIED"
            Cells(u, 10).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            dtStr = Split(item, ":")(1)
            Cells(u, 10) = Format(ParseDateZ(dtStr), "yyyy-mm-dd hh:mm:ss")
          'Identify the event attendee, ATTENDEE
          Case Left(item, 8) = "ATTENDEE"
            eStr = Split(item, ";")(3)
            eStr = Mid(eStr, InStr(eStr, "mailto:"), Len(eStr))
            Cells(u, 11) = eStr
          'Identify the SEQUENCE
          Case Left(item, 8) = "SEQUENCE"
            eStr = Split(item, ":")(1)
            Cells(u, 12) = eStr
          'Identify the event subtype, X-CONFLUENCE-SUBCALENDAR-TYPE
          Case Left(item, 29) = "X-CONFLUENCE-SUBCALENDAR-TYPE"
            eStr = Split(item, ":")(1)
            Cells(u, 13) = eStr
          'Identify the status, STATUS
          Case Left(item, 6) = "STATUS"
            eStr = Split(item, ":")(1)
            Cells(u, 14) = eStr
          'Identify the end of the event, END:VEVENT
          Case Left(item, 12) = "END:VEVENT"
            eStr = Split(item, ":")(1)
            Cells(u, 15) = eStr
        End Select
      
      Next item

    End If
    
  Loop

ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes).Name = "PROG_ABS"

FUNCTION_EXIT:
    MsgBox Message, vbOKOnly
Exit Sub

ERROR_HANDLER:
    Debug.Print Err.Description
    Message = "Error in this line: " & line & " Error text: " & Err.Description
  GoTo FUNCTION_EXIT

End Sub

Private Function ReadEvent(ByRef stream As ADODB.stream, ByRef line As String) As Boolean

'***************************************************************************
'Purpose: Parses an *.ics object stream and extracts the events ONLY
'Inputs: Object stream and empty line string
'Outputs: Object stream and line string with events details
'***************************************************************************

  Dim result As Boolean
  Dim s As String
On Error GoTo ERROR_HANDLER
  s = stream.ReadText(adReadLine)
  If s Like ("BEGIN:VEVENT") Then
    Do Until s Like ("END:VEVENT")
      If Len(s) >= 73 Then
        s = s + stream.ReadText(adReadLine)
      Else
        s = s
      End If
      line = line + vbCrLf + s
      s = stream.ReadText(adReadLine)
    Loop
    line = line + vbCrLf + "END:VEVENT"
    result = True
  Else
    result = ReadEvent(stream, line)
  End If

FUNCTION_EXIT:

  ReadEvent = result

Exit Function

ERROR_HANDLER:
  Debug.Print Err.Description
  result = False
  GoTo FUNCTION_EXIT
End Function

Function ParseDateZ(dtStr As String)

'***************************************************************************
'Purpose: Parses a datetime string and returns a date
'Input: String containing datetime information
'Output: Date serial number
'***************************************************************************

    Dim dtArr() As String
    Dim dt As Date
    dtArr = Split(Replace(dtStr, "Z", ""), "T")
    dt = DateSerial(Left(dtArr(0), 4), Mid(dtArr(0), 5, 2), Right(dtArr(0), 2))
    If UBound(dtArr) > 1 Then
        dt = dt + TimeSerial(Left(dtArr(1), 2), Mid(dtArr(1), 3, 2), Right(dtArr(1), 2))
    End If
    ParseDateZ = dt
End Function

Function ParseDate(dtStr As String)

'***************************************************************************
'Purpose: Parses a datet string and returns a date
'Input: String containing date information
'Output: Date serial number
'***************************************************************************

    Dim dt As Date
    dt = DateSerial(Left(dtStr, 4), Mid(dtStr, 5, 2), Right(dtStr, 2))
    ParseDate = dt
End Function
