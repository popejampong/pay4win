VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Attribute VB_Name = "clsReadTextFile"
' Text File Class for reading files
' Copyright 2000, Chuck Grimsby
' All Rights Reserved
' Use is free without restriction,
' I just keep the copyright!

' Class keeps 3 lines for the programmer to use:
'  .Text = current line to work with
'  .PreviousLine = the last line.
'  .NextLine = The next line for 'look ahead' operations
'  (occasionally useful to know what's next!)

' The class has it's own EOF property (EndOfFile).
' Use that rather then the channel's EOF.
' Since the class is doing it's own handling of channel numbers, inputs, etc.,
' using one that the class isn't handling could return erronious information.

' The BytesLeft should also be ignored.
' The class uses it internally for it's own purposes....
' The .LeftOver property is also something that the class uses
' for it's own use. Leave it alone!

' BufferSize can be set to whatever the programmer needs/wants.
' On files that have more then 4096 characters per line,
' this can be changed.  Unless that's the case, 4096 bytes
' (the default) usually works quite well.

' The .StripLeadingSpaces and .StripTrailingSpaces properties
' are used for formatting of the line within your program.
' Use as needed/desired.

' The .NoBlankLines property is for use when you'd rather
' not deal with blank lines.  Note that this setting doesn't
' apply to the .NextLine propery or the .PreviousLine property.

' The .StripNulls becomes useful in files that are "padded".
' I use this a lot on files that come off of Main Frames and
' are really ment to be printed, rather then read into a program.
' Most files =don't= have nulls (Chr$(0)) in them, so normally
' you'll leave this off.  (Turning it on slows the reads!)

' The .OnlyAlphaNumericCharacters property is also fairly useless
' except in some extremely rare conditions.  I can't remember why
' I put this in here, it changes the class so that it only
' returns the characters in the ASCII range of 32 - 127
' (Characters that can be printed), but not the so-called
' "upper-ASCII" characters that can also be printed.
' International users will probably find it =much= more
' useless then those in the US.

' The .CountOnlyNonBlankLines is for the count of lines.
' Personally, I almost always leave this set to False as
' I'd rather have a count that is more true to the format
' of the file, but I've had occurances where this is useful.

' You can set the LineDelimiter property to be whatever you
' want/need.  This becomes =really= helpful in files that are
' created on systems that don't use either Cr, Lf, or CrLf
' as line delimiters.
' The class will automatically try to figure out what the line
' Delimiter is if you don't set this.  It looks for Carridge
' Returns (Chr$(13)), Line Feeds (Chr$(10)) and the 2 of them
' together.
' The class will automatically look for FormFeeds and use those
' as additional Line Delimiters.

' This class can also be used on FixedWidth files.
' Just set the .FixedWidthLineLength property to the length
' of the line, and that's it.
' When the .FixedWidthLineLength >0, the line delimiters
' are ignored, so feel free to do so in your code.

' The methods used are:
'  .cfOpen to open the file set in the .FileName property
'  .csGetALine which is used to get the next line from the file
'  and .cfCloseFile to close the file when you're done with it.
'  Watch the .EndOfFile property to know when you're done.

' sample useage:
'Sub ReadAFile(strFileName)
'   Dim myTextFile As New clsReadTextFile       ' Create a new class internal to your program.
'   Dim myString As String ' Local String to play with in your program
'   Dim intError As Integer
'
'   myTextFile.FileName = strFileName            ' set the file name to read
'   myTextFile.NoBlankLines = True               ' Don't return blank lines!
'   myTextFile.CountOnlyNonBlankLines = False    ' Count all lines regardless of whether or not they are returned
'   myTextFile.StripLeadingSpaces = False        ' leave any leading spaces.
'   myTextFile.StripTrailingSpaces = False       ' leave the trailing spaces too!
'   myTextFile.StripNulls = True                 ' eliminate Chr$(0)'s
'   myTextFile.OnlyAlphaNumericCharacters = True ' don't send me any characters I can't use!
'
'   intError = myTextFile.cfOpenFile  ' open the file, return any errors in doing so (class doesn't handle them!)
'   If intError = 0 Then
'      ' no error in opening the file has occured
'      While Not myTextFile.EndOfFile                 ' Watch for the end of the file
'         myTextFile.csGetALine                       ' Tells the class to go to a new line
'         myString = myTextFile.Text                  ' set your string = to the current string of the class
'         Debug.Print myTextFile.LinesRead, myString  ' Show work in the Debug window. Optional!  Don't do in production!
'         ' want to see the next line?
'         If myTextFile.NextLine = <whatever> Then....
'         ' want to see the Last line?
'         If myTextFile.PreviousLine = <whatever> Then ....
'         <do whatever with the string here.>
'         ' want to find something in the file?
'         '     find a line within the file:
'         '      myTextFile.csFindLine StringToSearchFor, True
'         '          Note: Change True To False if you care about case!
'         '      myString = myTextFile.Text
'         ' on larger files, you may not want to do a line-by-line
'         ' search, so do:
'         '      myTextFile.csFindInFile StringToSearchFor, True
'         '      myString = myTextFile.Text
'      Wend
'   Else
'      ' handle the error here!
'      MsgBox Error(intError)
'   End If
'   myTextFile.cfCloseFile    ' close the file. We're done with it!
'   Set myTextFile = Nothing  ' Always set your objects to nothing when you're done with them!
'End Sub

' the code itself is below:

'local variable(s) to hold property value(s)
Private mvarBytesLeft As Currency
Private mvarText As String
Private mvarLeftOver As String
Private mvarEndOfFile As Boolean
Private mvarChannelNumber As Integer
Private mvarBufferSize As Long
Private mvarFileName As String
Private mvarNoBlankLines As Boolean
Private mvarStripLeadingSpaces As Boolean
Private mvarStripTrailingSpaces As Boolean
Private mvarLinesRead As Double
Private mvarCountOnlyNonBlankLines As Boolean
Private mvarNextLine As String
Private mvarPreviousLine As String
Private mvarLineDelimiter As String
Private mvarFixedWidthLineLength As Integer
Private mvarOnlyAlphaNumericCharacters As Boolean
Private mvarStripNulls As Boolean

Public Property Let StripNulls(ByVal vData As Boolean)
    mvarStripNulls = vData
End Property

Public Property Get StripNulls() As Boolean
    StripNulls = mvarStripNulls
End Property

Public Property Let OnlyAlphaNumericCharacters(ByVal vData As Boolean)
    mvarOnlyAlphaNumericCharacters = vData
End Property

Public Property Get OnlyAlphaNumericCharacters() As Boolean
    OnlyAlphaNumericCharacters = mvarOnlyAlphaNumericCharacters
End Property

Public Property Let FixedWidthLineLength(ByVal vData As Integer)
    mvarFixedWidthLineLength = vData
End Property

Public Property Get FixedWidthLineLength() As Integer
    FixedWidthLineLength = mvarFixedWidthLineLength
End Property

Public Property Let LineDelimiter(ByVal vData As String)
    mvarLineDelimiter = vData
End Property

Public Property Get LineDelimiter() As String
    LineDelimiter = mvarLineDelimiter
End Property

Public Property Let PreviousLine(ByVal vData As String)
    mvarPreviousLine = vData
End Property

Public Property Get PreviousLine() As String
    PreviousLine = mvarPreviousLine
End Property

Public Property Let NextLine(ByVal vData As String)
    mvarNextLine = vData
End Property

Public Property Get NextLine() As String
    NextLine = mvarNextLine
End Property

Public Property Let CountOnlyNonBlankLines(ByVal vData As Boolean)
    mvarCountOnlyNonBlankLines = vData
End Property

Public Property Get CountOnlyNonBlankLines() As Boolean
    CountOnlyNonBlankLines = mvarCountOnlyNonBlankLines
End Property

Public Property Let LinesRead(ByVal vData As Double)
    mvarLinesRead = vData
End Property

Public Property Get LinesRead() As Double
    LinesRead = mvarLinesRead
End Property

Public Property Let StripTrailingSpaces(ByVal vData As Boolean)
    mvarStripTrailingSpaces = vData
End Property

Public Property Get StripTrailingSpaces() As Boolean
    StripTrailingSpaces = mvarStripTrailingSpaces
End Property

Public Property Let StripLeadingSpaces(ByVal vData As Boolean)
    mvarStripLeadingSpaces = vData
End Property

Public Property Get StripLeadingSpaces() As Boolean
    StripLeadingSpaces = mvarStripLeadingSpaces
End Property

Public Property Let NoBlankLines(ByVal vData As Boolean)
    mvarNoBlankLines = vData
End Property

Public Property Get NoBlankLines() As Boolean
    NoBlankLines = mvarNoBlankLines
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let BufferSize(ByVal vData As Long)
    mvarBufferSize = vData
End Property

Public Property Get BufferSize() As Long
    BufferSize = mvarBufferSize
End Property

Public Property Let ChannelNumber(ByVal vData As Integer)
    mvarChannelNumber = vData
End Property

Public Property Get ChannelNumber() As Integer
    ChannelNumber = mvarChannelNumber
End Property

Public Property Let EndOfFile(ByVal vData As Boolean)
    mvarEndOfFile = vData
End Property

Public Property Get EndOfFile() As Boolean
    EndOfFile = mvarEndOfFile
End Property

Public Property Let LeftOver(ByVal vData As String)
    mvarLeftOver = vData
End Property

Public Property Get LeftOver() As String
    LeftOver = mvarLeftOver
End Property

Public Property Let Text(ByVal vData As String)
    mvarText = vData
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Public Property Let BytesLeft(ByVal vData As Currency)
    mvarBytesLeft = vData
End Property

Public Property Get BytesLeft() As Currency
    BytesLeft = mvarBytesLeft
End Property

Function cfOpenFile() As Integer
    On Error GoTo Open_File_Error
    
    ChannelNumber = FreeFile
    
    Open FileName For Binary As #ChannelNumber
    BytesLeft = LOF(ChannelNumber)
    
    ' no error, exit out:
    If Err.Number = 0 Then Exit Function
    
    ' error of somekind....
Open_File_Error:
       cfOpenFile = Err.Number
       
    ' reset error handling:
    On Error GoTo 0
End Function

Function cfCloseFile()
   Close #ChannelNumber
End Function

Private Sub Class_Initialize()
   LinesRead = 0
   BytesLeft = 0
   PreviousLine = vbNullString
   Text = vbNullString
   NextLine = vbNullString
   LeftOver = vbNullString
   LineDelimiter = vbNullString
   ' defaults:
   BufferSize = 4096
   StripTrailingSpaces = False
   StripLeadingSpaces = False
   StripNulls = False
   NoBlankLines = False
   OnlyAlphaNumericCharacters = False
   CountOnlyNonBlankLines = False
   FixedWidthLineLength = 0
End Sub

Private Sub Class_Terminate()
   Call cfCloseFile
End Sub

Public Sub csGetALine()
    Static NotFirst As Boolean
    
    ' move the last line read into the
    ' PreviousLine so it isn't lost:
    PreviousLine = Text
    
csGetALine_Start:
    
    If NotFirst = False Then
       ' load a line into the nextline property
       Call csGetNextLine
       ' set NotFirst to true so this doesn't happen again
       NotFirst = True
       ' now go through this routine the right way:
       GoTo csGetALine_Start
    Else
       ' move the next line into the text property
       Text = NextLine
       ' increment the line counter
       LinesRead = LinesRead + 1
    End If
    
    If (CountOnlyNonBlankLines = True) And _
       (Len(Trim$(Text)) = 0) Then
          ' decrement the line counter for blank lines
          ' if the user doesn't want to count those
          LinesRead = LinesRead - 1
    End If
    
    ' trim the string based on user settings:
    If StripLeadingSpaces = True Then
       Text = LTrim$(Text)
    End If
    
    If StripTrailingSpaces = True Then
       Text = RTrim$(Text)
    End If
    
    If OnlyAlphaNumericCharacters = True Then
       Text = AlphaNumOnly(Text)
    End If
    
    ' load the next line into the nextline propery:
    Call csGetNextLine
    
    ' if text is blank, loop through again if the user
    ' doesn't want blank lines:
    If (NoBlankLines = True And _
       Len(Trim$(Text)) = 0) Then
       If Not EndOfFile Then
          GoTo csGetALine_Start
       End If
    End If
End Sub

Private Sub csGetNextLine()
    Dim intFF As Integer
    Dim intX As Integer
    Dim Temp As String
    
    ' keep the buffer full:
    Call LoadBuffer
    
    If FixedWidthLineLength = 0 Then
       If LineDelimiter = vbNullString Then
          ' figure out what the line delimiter is:
          Call DetermineLineDelimiter
       End If
    
       ' see if someone stuck a form feed
       ' in the middle of the line:
       intFF = InStr(LeftOver, vbFormFeed)
       intX = InStr(LeftOver, LineDelimiter)
    
       ' figure out which is the left most:
       If intX > 0 Then
          If (intFF < intX) And (intFF > 0) Then
             intX = intFF
          End If
       Else
          If intFF > 0 Then
             intX = intFF
          End If
       End If
    
       ' trim the string to the leftmost deliminater:
       If intX > 0 Then
          NextLine = "" & left$(LeftOver, intX - 1)
          If intX = intFF Then
             LeftOver = Mid$(LeftOver, intX + 1)
          Else
             LeftOver = Mid$(LeftOver, _
                            intX + Len(LineDelimiter))
          End If
       Else
          NextLine = "" & LeftOver
          LeftOver = ""
       End If
    Else
       ' for Fixed Width files, ignore the delimiters,
       ' and use the length set by the programmer:
       NextLine = left$(LeftOver, FixedWidthLineLength)
       LeftOver = Mid$(LeftOver, FixedWidthLineLength + 1)
    End If
End Sub

Private Sub LoadBuffer()
    Dim intX As Integer
    
    If Not EndOfFile Then
       If Len(LeftOver) < BufferSize Then
          intX = BufferSize - Len(LeftOver)
          If BytesLeft < intX Then
             intX = BytesLeft
          End If
       End If
       If StripNulls = True Then
          ' it's easier/faster to do this here....
          LeftOver = LeftOver & _
                   NoNulls(Input$(intX, ChannelNumber))
       Else
          LeftOver = LeftOver & _
                   Input$(intX, ChannelNumber)
       End If
       ' update the number of bytes left in the file
       ' we're reading from:
       BytesLeft = BytesLeft - intX
    End If
    
    ' see if we're done:
    If (EOF(ChannelNumber)) Or _
       (BytesLeft = 0 And _
       LeftOver = "" And _
       NextLine = "") Then
          EndOfFile = True
    End If
    
    ' just something to free up displays....
    ' (make it nice for the users!)
    DoEvents
End Sub

Private Sub DetermineLineDelimiter()
    Dim intCRLF As Integer  ' standard CrLf
    Dim intCR As Integer    ' Carridge Return
    Dim intLF As Integer    ' Line Feed
    Dim intX As Integer
    
    'find the leftmost CR,LF or FF:
    intCRLF = InStr(LeftOver, vbCrLf)
    intLF = InStr(LeftOver, vbLf)
    intCR = InStr(LeftOver, vbCr)
    intX = Len(LeftOver)
    
    ' Use whatever is leftmost as the delimiter!
    If (intCRLF < intX) And (intCRLF > 0) Then
       LineDelimiter = vbCrLf
       intX = intCRLF
    End If
    If (intLF < intX) And (intLF > 0) Then
       LineDelimiter = vbLf
       intX = intLF
    End If
    If (intCR < intX) And (intCR > 0) Then
       LineDelimiter = vbCr
       intX = intCR
    End If
End Sub

Private Function NoNulls(strIn As String) As String
    ' removes all Chr$(0)'s (ASCII Null's) from
    ' a string.
    Dim intI As Integer
    Dim strTemp As String
    strTemp = strIn
    intI = InStr(strTemp, Chr$(0))
    While intI > 0
       strTemp = left$(strTemp, intI - 1) & _
                Mid$(strTemp, intI + 1)
       intI = InStr(strTemp, Chr$(0))
    Wend
    NoNulls = strTemp
End Function

Private Function AlphaNumOnly(strIn As String) As String
    ' removes all but the ASCII Characters on the
    ' keyboard.
    Dim intI As Integer
    Dim strTemp As String
    strTemp = strIn
    While intI < Len(strTemp)
       intI = intI + 1
       Select Case Asc(Mid$(strTemp, intI, 1))
       Case 32 To 126
       ' do nothing. we want to keep these!
       Case Else
          ' get rid of everything else:
          strTemp = left$(strTemp, intI - 1) & _
                   Mid$(strTemp, intI + 1)
          If intI > 1 Then intI = intI - 1
       End Select
    Wend
    AlphaNumOnly = strTemp
End Function

Public Sub csFindLine(strStringToFind As String, _
          Optional bolIgnoreCase As Boolean = True)
    ' finds a string within a file by reading through
    ' the file line by line.
    Dim intI As Integer
    
    If Text = "" Then csGetALine
    Do
       If bolIgnoreCase = False Then
          ' don't care what case (upper/lower) the string is):
          intI = InStr(1, Text, strStringToFind, vbBinaryCompare)
       Else
          intI = InStr(1, Text, strStringToFind, vbTextCompare)
       End If
       ' not found, try the next line:
       If intI = 0 Then Call csGetALine
    Loop While Not EndOfFile And intI = 0
End Sub

Public Sub csFindInFile(strStringToFind As String, _
      Optional bolIgnoreCase As Boolean = True)
'Faster way to find a line that's in a large file,
'but slower returning it back to the program.
    
    Dim intI As Integer
    Dim Temp As String
    
    'clear the nextline and text properties,
    ' since they no longer have value to us:
    NextLine = vbNullString
    Text = vbNullString
    
    ' makesure the buffer is full:
    Call LoadBuffer
    
    Do
       If bolIgnoreCase = False Then
          ' don't care what case (upper/lower) the string is):
          intI = InStr(1, LeftOver, _
                strStringToFind, vbBinaryCompare)
       Else
          intI = InStr(1, LeftOver, _
                strStringToFind, vbTextCompare)
       End If
       ' not found, try the next set of characters:
       If intI = 0 Then
          ' save the right most characters in case we
          ' only got part of them:
          LeftOver = right$(LeftOver, _
                      Len(strStringToFind))
          ' reload the buffer:
          Call LoadBuffer
       End If
       ' let the screen updates happen:
       DoEvents
       ' note: since we're doing some pretty weird things
       ' here, we can't rely on EndOfFile, so watch the
       ' BytesLeft property instead to find out if the class
       ' is at the end of the file.
    Loop While BytesLeft > 0 And intI = 0
    
    If intI > 0 Then
       ' found it!
       ' put everything that's before what we've searched for
       ' into the previousline:
       Temp = left$(LeftOver, intI - 1)
       LeftOver = Mid$(LeftOver, intI)
       csGetALine
       If FixedWidthLineLength = 0 Then
          ' work backwards through temp
          'to find the line delimiter:
          Do
             intI = intI - 1
             If Mid$(Temp, intI, Len(LineDelimiter)) = _
                      LineDelimiter Then Exit Do
          Loop While intI > 0 And intI > Len(LineDelimiter)
          ' move those characters from Temp into
          ' the front of text to make a complete line:
          If intI > 0 Then
             Text = Mid$(Temp, intI + Len(LineDelimiter)) _
                      & Text
             ' put the rest of temp into the previousline
             ' property so it's available to the programmer:
             PreviousLine = left$(Temp, _
                            intI - Len(LineDelimiter))
          End If
       End If
    Else
       'The string we're searching for wasn't there!
       'Text, NextLine & Leftover no longer have
       'value, so clear them out to reduce confusion:
       Text = vbNullString
       NextLine = vbNullString
       LeftOver = vbNullString
    End If
End Sub


