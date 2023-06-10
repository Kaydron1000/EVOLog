VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_ExcelWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Implements ILogChannel

Private Const constConduitFileExtension As String = "xlsx"

Dim pParentLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pDefaultLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pLoggingLevelNames As Collection

Dim pDefaultHeaders() As String
Dim pLoggingWB As Workbook
Dim pLoggingSht As Worksheet
Dim pLogFolderPath As String
Dim pLogFilePath As String
Dim pLogFileName As String

Property Get ConduitFileExtension() As String
    ConduitFileExtension = constConduitFileExtension
End Property

Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLoggers = pParentLoggers
End Property
Property Get ConduitName() As String
    ConduitName = pConduitName
End Property

Property Let ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLogginglLevel = Level
End Property
Property Get ConduitLoggingLevel() As LoggingLevels
    ConduitLoggingLevel = pConduitLoggingLevel
End Property
Property Let MessageTemplate(Template As String)
    pMessageTemplate = Template
End Property
Property Get MessageTemplate() As String
    MessageTemplate = pMessageTemplate
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

Sub InitConduit(ParentLogger As cEvoLogger, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
      
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub
Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment, Optional LogFileName As String = "{DATETIME}_{CONDUITNAME}", _
         Optional LogFolderPath As String = "LogOutput", Optional FileAccess As FileAccesses = FileAccesses.Append, _
         Optional WorksheetName As String = "{DATETIME}_{CONDUITNAME}", Optional Headers As Variant = Nothing)
    
    Dim n As Integer
    Dim fileExists As Boolean
    Dim fileName As String
    Dim folderPath As String
    Dim workbookOpen As Boolean
    Dim shtName As String
    Dim baseShtName As String
    
    folderPath = LogFolderPath
    fileName = LogFileName
    shtName = WorksheetName
    ''Checking if LogFolderPath is a valid Folder
    If Dir(BuildPath(folderPath, "\"), vbDirectory) <> "." Then Err.Raise 76
    
    If fileName = "" Then fileName = ConduitName & "." & constSinkFileExtension
        If shtName = "" Then shtName = ConduitName
            ''*TODO: need to get headers if specified
            
        fileExists = Len(Dir(BuildPath(folderPath, fileName))) > O
        If fileExists And FileAccess = FileAccesses.Append Then
            pLogFolderPath = folderPath
            pLogFileName = fileName
            pLogFilePath = BuildPath(folderPath, fileName)
            
            ''Checking for workbook being opened
            If ContainsKey(Application.Workbooks, Me.LogFileName) Then i
            
            If GetPath.GetLocalPath(Application.Workbooks(Me.LogFileName).Path) = Me.LogFolderPath Then
                workbookOpen = True
            End If
        End If
        
        ''If workbook is open set logging workbook to open workbook else open the workbook
        Application.DisplayAlerts = False
        If workbookOpen Then
            Set pLoggingWB = Application.Workbooks(Me.LogFileName)
        Else
            Set pLoggingWB = Application.Workbooks.Open(fileName:=Me.LogFilePath, UpdateLinks:=False, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False)
        End If
        Application.DisplayAlerts = True
    ElseIf fileExists And FileAccess = FileAccesses.Overwrite Then
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
        
        ''Overwriting workbook
        Application.DisplayAlerts = False
        If workbookOpen Then
            Application.Workbooks(Me.LogFileName).Close SaveChanges:=False
        Else
            Set pLoggingWB = Application.Workbooks.Add
            pLoggingWB.SaveAs fileName:=Me.LogFilePath, FileFormat:=X1FileFormat.xlWorkbookDefault, _
            AccessMode:=X1SaveAsAccessMode.x1NoChange, ConflictResolution:=X1SaveConflictResolution.xlLocalSessionChanges
        End If
        Application.DisplayAlerts = True
    ElseIf fileExists And FileAccess = FileAccesses.CreateNew Then
        fileName = GetUniqueLogFileName(folderPath, fileName)
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
        
        ''Creating new workbook
        Application.DisplayRlerts = False
        Set pLoggingWB = Application.Workbooks.Add
        pLoggingWB.SaveAs fileName:=Me.LogFilePath, FileFormat:=XlFileFormat.xlWorkbookDefault, _
        AccessMode:=X1SaveAsAccessMode.x1NoChange, ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
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
    SetTitleAndHeaders pLoggingSht, "LogSheet - " & LogSinkName, pDefaultHeaders
    
    Me.ConduitName = ConduitName
    
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub

Sub SetTitleAndHeaders(sht As Worksheet, Title As String, Headers() As String)
    Dim n As Integer
    
    For n = LBound(Headers) To UBound(Headers)
        sht.Cells(2, (n - LBound(Headers)) + 1).Value = Headers(n)
    Next
    sht.Range("A1", sht.Cells(l, UBound(Headers) - LBound(Headers) + 1)).MergeCells = True
    sht.Range("Al1").Value = Title
End Sub
Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
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

Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
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

Private Function FormatMessage(MessageTemplate As String, LoggingLevel As LoggingLevels, Message As String)
    FormatMessage = UCase(MessageTemplate)
    FormatMessage = Replace(FormatMessage, "#LOGGINGLEVEL#", ParentLogger.LoggingLevelNames(LoggingLevel))
    FormatMessage = Replace(FormatMessage, "#DATE#", DateTime.Date)
    FormatMessage = Replace(FormatMessage, "#TIME#", DateTime.Time)
    FormatMessage = Replace(FormatMessage, "#TIMESTAMP#", DateTime.Now)
    FormatMessage = Replace(FormatMessage, "#MESSAGE#", Message)
End Function
Private Function TransposeArray(MyArray As Variant) As Variant
    Dim X As Long, Y As Long
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
    For X = minX To maxX
        For Y = minY To max¥
            tempArr(Y, X) = MyArray(X, Y)
        Next Y
    Next X

    'Output Array
    TransposeArray = tempArr
End Function
Private Function GetUniqueLogFileName(LogFolderPath As String, LogFileName As String) As String
    Dim n As Integer
    Dim folderPath As String
    Dim fileName As String
    Dim extStart As Integer
    Dim baseName As String
    Dim extName As String
    Dim tempLogFileName As String
    
    folderPath = LogFolderPath
    
    fileName = LogFileName
    
    On Error GoTo errMark
        If Dir(BuildPath(folderPath, "\"), vbDirectory) <> "." Then Err.Raise 76
        
        extStart = InStrRev(fileName, ".")
        If extStart = O Then fileName = fileName & "." & SinkFileExtension
        
        extStart = InStrRev(fileName, ".")
        baseName = Left(fileName, extStart - 1)
        extName = Right(fileName, Len(fileName) - extStart)
        
        tempLogFileName = baseName & "." & extName
        
        n = 1
        Do While Len(Dir(BuildPath(folderPath, tempLogFileName))) <> 0 And ContainsKey(Application.Workbooks, tempLogFileName) And n <= 999
            tempLogFileName = baseName & "_" & Format(n, "000") & "." & extName
            n = n + 1
        Loop
        If n >= 999 Then
            Err.Raise 999, "New File Name", "Cannot create log file. Over 1000 instances exist of " & fileName & "™ exists in " & LogFolderPath & "."
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
Private Sub Class_Initialize()
    pMessageTemplate = "[#LOGGINGLEVEL#]" & vbTab & "[#DATE# #TIME#]: " & vbTab & "#MESSAGE#"
    ReDim pDefaultHeaders(0 To 2)
    pDefaultHeaders(0) = "Log Level"
    pDefaultHeaders(l) = "Time Stamp"
    pDefaultHeaders(2) = "Log Message"
    
    Set pParentLoggers = New Collection
    
    Set pLoggingLevelNames = New Collection
    pDefaultLoggingLevel = Information
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
End Sub
Private Sub Class_Terminate()
End Sub
Property Get ILogConduit_ParentEvoLoggers() As cEvoLoggers
    Set ILogConduit_ParentEvoLoggers = Me.ParentEvoLoggers
End Property

Property Get ILogConduit_ConduitName() As String
    ILogConduit_ConduitName = Me.ConduitName
End Property

Property Let ILogConduit_ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = ILogConduit_ConduitLoggingLevel
End Property

Property Get ILogConduit_ConduitLoggingLevel() As LoggingLevels
    ILogConduit_ConduitLoggingLevel = pConduitLoggingLevel
End Property

Property Let ILogConduit_MessageTemplate(Template As String)
    pMessageTemplate = ILogConduit_MessageTemplate
End Property

Property Get ILogConduit_MessageTemplate() As String
    ILogConduit_MessageTemplate = pMessageTemplate
End Property

Sub ILogConduit_InitConduit(ParentLogger As cEvoLogger, ConduitName As String, ConduitLoggingLevel As LoggingLevels)
    Me.InitConduit ParentLogger, ConduitName, ConduitLoggingLevel
End Sub

Sub ILogConduit_ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    Me.ChannelLogEntry LogMessageLevel, Message
End Sub

Sub ILogConduit_ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Me.ChannelLogEntries LogMessageLevels, Messages, EntryCount
End Sub




