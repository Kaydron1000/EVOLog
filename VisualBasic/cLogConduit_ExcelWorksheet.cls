VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_ExcelWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''EVOLog - a logging platform for VBA inspired by Serilog Library for .NET.
''--------------------------------------
''https://github.com/Kaydron1000/EVOLog
''--------------------------------------
''
''Copyright (C) 2023  Brandon Neff
''
''SPDX-License-Identifier: GPL-2.0
''This program is free software; you can redistribute it and/or
''modify it under the terms of the GNU General Public License
''as published by the Free Software Foundation; either version 2
''of the License, or (at your option) any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''GNU General Public License for more details.
''
''You should have received a copy of the GNU General Public License
''along with this program; if not, write to the Free Software
''Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''This log conduit channels log messages to an Excel worksheet.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements ILogConduit

Private Const constConduitFileExtension As String = "xlsx"

Dim pParentLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pLoggingLevelNames As Collection

Dim pDefaultHeaders() As String
Dim pLoggingWB As Workbook
Dim pLoggingSht As Worksheet
Dim pLogFolderPath As String
Dim pLogFilePath As String
Dim pLogFileName As String

'*******************************************************************************
' Parent Loggers that log messages are channeled from.
'*******************************************************************************
Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLogger = pParentLogger
End Property

'*******************************************************************************
' ILogConduit Name
'*******************************************************************************
Property Get ConduitName() As String
    ConduitName = pConduitName
End Property

'*******************************************************************************
' Set the logging level that is channeled from this conduit.
'*******************************************************************************
Property Let ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = Level
End Property

'*******************************************************************************
' Get the logging level that is channeled from this conduit.
'*******************************************************************************
Property Get ConduitLoggingLevel() As LoggingLevels
    ConduitLoggingLevel = pConduitLoggingLevel
End Property

'*******************************************************************************
' Set the message template that describes the format of the log message.
'*******************************************************************************
Property Let MessageTemplate(Template As String)
    pMessageTemplate = Template
End Property

'*******************************************************************************
' Get the message template that describes the format of the log message.
'*******************************************************************************
Property Get MessageTemplate() As String
    MessageTemplate = pMessageTemplate
End Property

'*******************************************************************************
' Log File extension
'*******************************************************************************
Property Get ConduitFileExtension() As String
    ConduitFileExtension = constConduitFileExtension
End Property


Property Get LogFileOpen()
    LogFileOpen = pFileOpen
End Property

Property Get LogFolderPath() As String
    LogFolderPath = pLogFolderPath
End Property
Property Get LogFilePath() As String
    LogFilePath = pLogFilePath
End Property
Property Get LogFileName() As String
    LogFileName = pLogFileName
End Property

Private Sub InitConduit(ParentLogger As cEvoLogger)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
End Sub
Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Information, Optional LogFileName As Variant = "#DATE##TIME#_#CONDUITNAME#", _
         Optional LogFolderPath As Variant = ".\LogOutput", Optional FileAccess As FileAccesses = FileAccesses.Append, _
         Optional WorksheetName As Variant = "#DATE##TIME#_#CONDUITNAME#", Optional Headers As Variant = Nothing)
    
    Dim n As Integer
    Dim workbookOpen As Boolean
    Dim shtName As String
    Dim baseShtName As String
    
    Dim fileExists As Boolean
    Dim FileName As String
    Dim folderPath As String
    
    pConduitName = ConduitName
    Me.ConduitLoggingLevel = ConduitLoggingLevel
    If Left(LogFolderPath, 2) = ".\" Then
        folderPath = BuildPath(ThisWorkbook.Path, LogFolderPath, "\")
    Else
        folderPath = BuildPath(LogFolderPath, "\")
    End If
    
    FileName = LogFileName
    ''Checking if LogFolderPath is a valid Folder
    If Dir(folderPath, vbDirectory) <> "." Then ''Err.Raise 76
        MkDir folderPath
    End If
    
    
    FileName = FormatFilePath(FileName)
    
    fileExists = Len(Dir(BuildPath(folderPath, FileName))) > 0
    
    shtName = Left(FormatString(CStr(WorksheetName)), 27)

    ''Checking for workbook being opened
    If ContainsKey(Application.Workbooks, Me.LogFileName) Then
        If LibFileTools.GetLocalPath(Application.Workbooks(Me.LogFileName).Path) = Me.LogFolderPath Then
            workbookOpen = True
        End If
    End If

    If fileExists And FileAccess = FileAccesses.Append Then
        pLogFolderPath = folderPath
        pLogFileName = FileName
        pLogFilePath = BuildPath(folderPath, FileName)
        
        
        ''If workbook is open set logging workbook to open the workbook, else open the workbook
        Application.DisplayAlerts = False
        If workbookOpen Then
            Set pLoggingWB = Application.Workbooks(Me.LogFileName)
        Else
            Set pLoggingWB = Application.Workbooks.Open(FileName:=Me.LogFilePath, UpdateLinks:=False, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False)
        End If
        Application.DisplayAlerts = True
    ElseIf fileExists And FileAccess = FileAccesses.Overwrite Then
        pLogFolderPath = folderPath
        pLogFileName = FileName
        pLogFilePath = BuildPath(folderPath, FileName)
        
        ''Overwriting workbook
        Application.DisplayAlerts = False
        If workbookOpen Then
            Application.Workbooks(Me.LogFileName).Close SaveChanges:=False
        Else
            Set pLoggingWB = Application.Workbooks.Add
            pLoggingWB.SaveAs FileName:=Me.LogFilePath, FileFormat:=XlFileFormat.xlWorkbookDefault, _
            AccessMode:=XlSaveAsAccessMode.xlNoChange, ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
        End If
        Application.DisplayAlerts = True
    ElseIf fileExists And FileAccess = FileAccesses.CreateNew Then
        FileName = GetUniqueLogFileName(folderPath, FileName)
        pLogFolderPath = folderPath
        pLogFileName = FileName
        pLogFilePath = BuildPath(folderPath, FileName)
        
        ''Creating new workbook
        Application.DisplayRlerts = False
        Set pLoggingWB = Application.Workbooks.Add
        pLoggingWB.SaveAs FileName:=Me.LogFilePath, FileFormat:=XlFileFormat.xlWorkbookDefault, _
        AccessMode:=XlSaveAsAccessMode.xlNoChange, ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
        Application.DisplayRlerts = True
    ElseIf Not fileExists Then
        pLogFolderPath = folderPath
        pLogFileName = FileName
        pLogFilePath = BuildPath(folderPath, FileName)
        
        Application.DisplayAlerts = False
        Set pLoggingWB = Application.Workbooks.Add
        pLoggingWB.SaveAs FileName:=Me.LogFilePath, FileFormat:=XlFileFormat.xlWorkbookDefault, _
        AccessMode:=XlSaveAsAccessMode.xlNoChange, ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
        Application.DisplayRlerts = True
    End If
    
    ''Setting up worksheet
    n = 1
    baseShtName = shtName
    Do While ContainsKey(pLoggingWB.Worksheets, shtName)
        shtName = baseShtName & "_" & Format(n, "000")
        n = n + 1
    Loop
    Set pLoggingSht = pLoggingWB.Worksheets.Add
    pLoggingSht.Name = shtName
    ''Headers
    SetTitleAndHeaders pLoggingSht, "LogSheet - " & ConduitName, pDefaultHeaders
End Sub

'*******************************************************************************
' Formats the string. Replaces the variables.
'*******************************************************************************
Private Function FormatString(FilePath As String, Optional TimeStampFormatString As Variant = "yyyymmdd_hhmmss", _
                               Optional DateFormatString As Variant = "mm/dd/yy", Optional TimeFormatString As Variant = "hh:nn:ss") As String
    ''TODO: Fix milliseconds formatting
    Dim a, b, c, d
    tsts
    FormatString = UCase(FilePath)
    FormatString = Replace(FormatString, "#LOGGINGLEVEL#", pLoggingLevelNames(ConduitLoggingLevel))
    FormatString = Replace(FormatString, "#DATE#", Format(Now, DateFormatString))
    FormatString = Replace(FormatString, "#TIME#", Format(Now, TimeFormatString) & Right(Format(Timer, "0.000"), 4))
    FormatString = Replace(FormatString, "#TIMESTAMP#", Format(Now, TimeStampFormatString) & Right(Format(Timer, "0.000"), 4))
    FormatString = Replace(FormatString, "#CONDUITNAME#", Me.ConduitName)
    
'    s = Format(Now, "yyyy-mm-dd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function
Private Function tsts()
    Dim a, b, c, d, e, f, g, h, i
    Dim ab, bb, cb, db, eb, fb, gb, hb, ib
    Dim r, s, t, u, v, w, x, y, z
    Dim rb, sb, tb, ub, vb, wb, xb, yb, zb
    
    a = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "34.555")
    b = FormatTime(#2:22:55 PM#, "hh;nn;ss.ll", "34.666")
    c = FormatTime(#2:22:55 PM#, "hh;nn;ss.l", "34.777")
    
    e = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", CSng("34.777"))
    f = FormatTime(#2:22:55 PM#, "hh;nn;ss.ll", CSng("34.888"))
    g = FormatTime(#2:22:55 PM#, "hh;nn;ss.l", CSng("34.999"))
    ab = "14;22;55.555" = a
    bb = "14;22;55.66" = b
    cb = "14;22;55.7" = c
    
    eb = "14;22;55.777" = e
    fb = "14;22;55.88" = f
    gb = "14;22;55.9" = g
    
    r = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "34.5")
    s = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "34.66")
    t = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "777")
    u = FormatTime(#2:22:55 PM#, "hh;nn;ss.ll", "888")
    z = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "9")
    
    v = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", CSng("34.7"))
    w = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", CSng("34.88"))
    x = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", CSng("34"))
    y = FormatTime(#2:22:55 PM#, "hh;nn;ss.l", CSng(".00"))
    
    
    d = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "900")
    a = FormatTime(#2:22:55 PM#, "hh;nn;ss.lll", "aaa")
    rb = "14;22;55.500" = r
    sb = "14;22;55.660" = s
    tb = "14;22;55.777" = t
    ub = "14;22;55.88" = u
    
    vb = "14;22;55.700" = v
    wb = "14;22;55.880" = w
    xb = "14;22;55.000" = x
    yb = "14;22;55.0" = y
    zb = "14;22;55.009" = z
    db = "14;22;55.900" = d
End Function
Private Function FormatTime(InputTime As Date, Optional TimeFormatString As String = "hh:nn:ss", Optional Milliseconds As Variant) As String
    Dim initialTimeStrg As String
    Dim retTimeStrg As String
    Dim hrStrg As String
    Dim minStrg As String
    Dim secStrg As String
    Dim milliStrg As String
    Dim bb
    
    If Not IsMissing(Milliseconds) Then
         If TypeName(Milliseconds) = "Single" Then
            milliStrg = Right(Format(Milliseconds, "0.000"), 3)
         ElseIf TypeName(Milliseconds) = "String" Then
            If IsNumeric(Milliseconds) Then
                If InStr(Milliseconds, ".") > 0 Then
                    milliStrg = Mid(Milliseconds & "000", InStr(Milliseconds, ".") + 1, 3)
                Else
                    milliStrg = Format(Left(Milliseconds, 3), "000")
                End If
            Else
                bb = Application.VBE.ActiveCodePane.CodeModule.Parent.Name & "." & Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Application.VBE.ActiveCodePane.TopLine, 0)
                Call Err.Raise(-44, bb, "Invalid value sent for Milliseconds parameter")
            End If
        Else
               bb = Application.VBE.ActiveCodePane.CodeModule.Parent.Name & "." & Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Application.VBE.ActiveCodePane.TopLine, 0)
               Call Err.Raise(-44, bb, "Invalid value sent for Milliseconds parameter")
        End If
    End If
    Dim a As Single

    initialTimeStrg = Format(InputTime, "hh:mm:ss")

    hrStrg = Mid(initialTimeStrg, 1, 2)
    minStrg = Mid(initialTimeStrg, 4, 2)
    secStrg = Mid(initialTimeStrg, 7, 2)

    retTimeStrg = TimeFormatString
    retTimeStrg = Replace(retTimeStrg, "hh", Format(CInt(hrStrg), "00"))
    retTimeStrg = Replace(retTimeStrg, "h", Format(CInt(hrStrg), "0"))
    retTimeStrg = Replace(retTimeStrg, "nn", Format(CInt(minStrg), "00"))
    retTimeStrg = Replace(retTimeStrg, "n", Format(CInt(minStrg), "0"))
    retTimeStrg = Replace(retTimeStrg, "ss", Format(CInt(secStrg), "00"))
    retTimeStrg = Replace(retTimeStrg, "s", Format(CInt(secStrg), "0"))
    retTimeStrg = Replace(retTimeStrg, "lll", Left(milliStrg, 3))
    retTimeStrg = Replace(retTimeStrg, "ll", Left(milliStrg, 2))
    retTimeStrg = Replace(retTimeStrg, "l", Left(milliStrg, 1))
    
    FormatTime = retTimeStrg
End Function

'*******************************************************************************
' Formats the filepath. Replaces the variables and adds file extension if needed.
'*******************************************************************************
Private Function FormatFilePath(FilePath As String, Optional TimeStampFormatString As Variant = "yyyymmdd_hhmmss", _
                               Optional DateFormatString As Variant = "yymmdd", Optional TimeFormatString As Variant = "hh;nn;ss") As String
    FormatFilePath = FormatString(FilePath, TimeStampFormatString, DateFormatString, TimeFormatString)
    If Right(FormatFilePath, Len(constConduitFileExtension) + 1) <> "." & constConduitFileExtension Then
        FormatFilePath = FormatFilePath & "." & constConduitFileExtension
    End If
End Function

Sub SetTitleAndHeaders(sht As Worksheet, Title As String, Headers() As String)
    Dim n As Integer
    
    sht.Cells(2, 1).Value = "Line No"
    For n = LBound(Headers) To UBound(Headers)
        sht.Cells(2, (n - LBound(Headers)) + 1).Value = Headers(n)
    Next
    sht.Range("A1", sht.Cells(l, UBound(Headers) - LBound(Headers) + 1)).MergeCells = True
    sht.Range("A1").Value = Title
End Sub
Private Sub ChannelLogArtifact(LogMessageLevel As LoggingLevels, Message As String)
    Dim arr() As String
    Dim templateArr() As String
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    templateArr = Split(MessageTemplate, vbTab)
    ReDim arr(LBound(arr) To UBound(arr))
    
    For cnt = LBound(arr) To UBound(arr)
        arr(cnt) = FormatMessage(templatelArr(cnt), LogMessageLevel, Message)
    Next
    
    maxX = UBound(arr, 1)
    minX = LBound(arr, 1)
    maxY = UBound(arr, 2)
    minY = LBound(arr, 2)
    
    pLoggingSht.Cells(pLoggingSht.Rows.Count, 1).End(x1Up).Offset(l, 0).Resize(maxY - minY + 1, maxX - minX + 1).Value = Application.WorksheetFunction.Transpose(TransposeArray(arr))
End Sub

Private Sub ChannelLogArtifacts(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Dim cnt As Long
    Dim MessageCnt As Integer
    Dim arr() As String
    Dim templateArr() As String
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    templateArr = Split(MessageTemplate, vbTab)
    ReDim arr(LBound(LogMessageLevels) To LBound(LogMessageLevels) + EntryCount - 1, LBound(templateArr) To UBound(templateArr))
    
    For MessageCnt = LBound(LogMessageLevels) To LBound(LogMessageLevels) + EntryCount - 1
        For cnt = LBound(templateArr) To UBound(templateArr)
            arr(MessageCnt, cnt) = FormatMessage(templatelArr(cnt), LogMessageLevels(MessageCnt), Messages(MessageCnt))
        Next
    Next
    
    maxX = UBound(arr, 2)
    minX = LBound(arr, 2)
    maxY = UBound(arr, 1)
    minY = LBound(arr, 1)
    
    pLoggingSht.Cells(pLoggingSht.Rows.Count, 1).End(x1Up).Offset(l, 0).Resize(maxY - minY + 1, maxX - minX + 1).Value = Application.WorksheetFunction.Transpose(TransposeArray(arr))
End Sub

Private Function TransposeArray(MyArray As Variant) As Variant
    Dim x As Long, y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    Dim tempArr As Variant
    
    'Get Upper and Lower Bounds
    maxX = UBound(MyArray, 1)
    minX = LBound(MyArray, 1)
    maxY = UBound(MyArray, 2)
    minY = LBound(MyArray, 2)
    
    'Create New Temp Array
    ReDim tempArr(minY To max¥, minX To maxX)
    
    'Transpose the Array
    For x = minX To maxX
        For y = minY To max¥
            tempArr(y, x) = MyArray(x, y)
        Next y
    Next x

    'Output Array
    TransposeArray = tempArr
End Function
Private Function GetUniqueLogFileName(LogFolderPath As String, LogFileName As String) As String
    Dim n As Integer
    Dim folderPath As String
    Dim FileName As String
    Dim extStart As Integer
    Dim baseName As String
    Dim extName As String
    Dim tempLogFileName As String
    
    folderPath = LogFolderPath
    
    FileName = LogFileName
    
    On Error GoTo errMark
        If Dir(BuildPath(folderPath, "\"), vbDirectory) <> "." Then Err.Raise 76
        
        extStart = InStrRev(FileName, ".")
        If extStart = O Then FileName = FileName & "." & SinkFileExtension
        
        extStart = InStrRev(FileName, ".")
        baseName = Left(FileName, extStart - 1)
        extName = Right(FileName, Len(FileName) - extStart)
        
        tempLogFileName = baseName & "." & extName
        
        n = 1
        Do While Len(Dir(BuildPath(folderPath, tempLogFileName))) <> 0 And ContainsKey(Application.Workbooks, tempLogFileName) And n <= 999
            tempLogFileName = baseName & "_" & Format(n, "000") & "." & extName
            n = n + 1
        Loop
        If n >= 999 Then
            Err.Raise 999, "New File Name", "Cannot create log file. Over 1000 instances exist of " & FileName & "™ exists in " & LogFolderPath & "."
        End If
errMark:
        If Err.Number <> O Then
            Debug.Print Err.Description
            Err.Raise Err.Number
            GetUniqueLogFileName = ""
        Else
            GetUniqueLogFileName = tempLogFileName
        End If
    On Error GoTo 0
End Function

Private Function FileNameFromPath(strFullPath As String) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Private Function FolderFromPath(strFullPath As String) As String
    FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function

Private Function ContainsKey(coll As Variant, key As String)
    Dim errNum1 As Integer
    Dim errNum2 As Integer
    Dim retItm As Variant
    
    On Error Resume Next
    Set retItm = coll(key)
    errNum1 = Err.Number
    On Error GoTo 0
    
    On Error Resume Next
    retItm = coll(key)
    errNum2 = Err.Number
    On Error GoTo 0
    
    If errNum1 = 0 Or errNum2 = 0 Then
        ContainsKey = True
    Else
        ContainsKey = False
    End If
End Function

Private Function SanatizeFileName(FileName As String)
    Dim illegalCharArray() As String
    Dim replaceChar() As String
    Dim n As Integer
    illegalCharArray = Array("\", "/", "<", ">", ":", """", "|", "?", "*")
    replaceChar = Array("_", "_", "-", "-", ";", "_", "-", "-")
    
    SanatizeFileName = FileName
    For n = LBound(illegalCharArray) To UBound(illegalCharArray)
        SanatizeFileName = Replace(SanatizeFileName, illegalCharArray(n), replaceChar(n))
    Next
End Function
'*******************************************************************************
' Class initialize method
'*******************************************************************************
Private Sub Class_Initialize()
    pMessageTemplate = "[#LOGGINGLEVEL#]" & vbTab & "[#DATE# #TIME#]: " & vbTab & "#MESSAGE#"
    ReDim pDefaultHeaders(0 To 2)
    pDefaultHeaders(0) = "Log Level"
    pDefaultHeaders(1) = "Time Stamp"
    pDefaultHeaders(2) = "Log Message"
    
    Set pParentLoggers = New cEvoLoggers
    
    Set pLoggingLevelNames = New Collection
    ''pDefaultLoggingLevel = Information
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
End Sub

'*******************************************************************************
' Class terminate method
'*******************************************************************************
Private Sub Class_Terminate()
End Sub


'*******************************************************************************
' Implementing ILogConduit Interface Methods and Properties
'*******************************************************************************


Private Property Get ILogConduit_ParentEvoLoggers() As cEvoLoggers
    Set ILogConduit_ParentEvoLoggers = Me.ParentEvoLoggers
End Property

Private Property Get ILogConduit_ConduitName() As String
    ILogConduit_ConduitName = Me.ConduitName
End Property

Private Property Let ILogConduit_ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = ILogConduit_ConduitLoggingLevel
End Property

Private Property Get ILogConduit_ConduitLoggingLevel() As LoggingLevels
    ILogConduit_ConduitLoggingLevel = pConduitLoggingLevel
End Property

Private Property Let ILogConduit_MessageTemplate(Template As String)
    pMessageTemplate = ILogConduit_MessageTemplate
End Property

Private Property Get ILogConduit_MessageTemplate() As String
    ILogConduit_MessageTemplate = pMessageTemplate
End Property

Private Sub ILogConduit_InitConduit(ParentLogger As cEvoLogger)
    InitConduit ParentLogger
End Sub

Private Sub ILogConduit_ChannelLogArtifact(LogArtifact As cEvoArtifact)
    ChannelLogArtifact LogArtifact
End Sub

Private Sub ILogConduit_ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    ChannelLogArtifacts LogArtifacts
End Sub

