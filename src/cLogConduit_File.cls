VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_File"
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
''This log conduit channels log messages to a text file.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements ILogChannel

Private Const constConduitFileExtension As String = "txt"
Dim pParentLoggers As Collection
Dim pLoggingLevelNames As Collection
Dim pDefaultLoggingLevel As LoggingLevels
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pFileNum As Integer
Dim pFileOpen As Boolean
Dim pLogFolderPath As String
Dim pLogFilePath As String
Dim pLogFileName As String

'*******************************************************************************
' Log File extension
'*******************************************************************************
Property Get ConduitFileExtension() As String
    ConduitFileExtension = constConduitFileExtension
End Property

'*******************************************************************************
' Parent Logger that is log messages are emitting from
'*******************************************************************************
Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLogger = pParentLogger
End Property

'*******************************************************************************
' Log File extension
'*******************************************************************************
Property Get ConduitName() As String
    ConduitName = pConduitName
End Property

Property Let ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = Level
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

Sub InitConduit(ParentLogger As cLogger, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
      
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub

Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment, Optional LogFileName As String = "{DATETIME}_{CONDUITNAME}", _
         Optional LogFolderPath As String = "LogOutput", Optional FileAccess As FileAccesses = FileAccesses.Append)
    
    Dim fileExists As Boolean
    Dim fileName As String
    Dim folderPath As String
    
    folderPath = LogFolderPath
    fileName = LogFileName
    ''Checking if LogFolderPath is a valid Folder
    If Dir(BuildPath(folderPath, "\"), vbDirectory) <> "." Then Err.Raise 76
    
    If fileName = "" Then fileName = LogSinkName & "." & constSinkFileExtension
    
    fileName = Format(DateTime.Now(), "yyyymmdd hhmmss_") & fileName
    fileExists = Len(Dir(BuildPath(folderPath, fileName))) > O
    
    If fileExists And FileAccess = FileAccesses.Append Then
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
    ElseIf fileExists And FileAccess = FileAccesses.Overwrite Then
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
    ElseIf fileExists And FileAccess = FileAccesses.CreateNew Then
        fileName = GetUniqueLogFileName(folderPath, fileName)
        
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
    ElseIf Not fileExists Then
        fileName = GetUniqueLogFileName(folderPath, fileName)
        pLogFolderPath = folderPath
        pLogFileName = fileName
        pLogFilePath = BuildPath(folderPath, fileName)
    End If
    
    pFileNum = FreeFile
    If FileAccess = FileAccesses.Append Then
        Open LogFilePath For Append As #pFileNum
        pFileOpen = True
    ElseIf FileAccess = FileAccesses.CreateNew Or FileAccess = FileAccesses.Overwrite Then
        Open LogFilePath For Output As #pFileNum
        pFileOpen = True
    End If
    
    
    Me.ConduitName = ConduitName
    
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub
Sub OpenLogFile()
    If Not pFileOpen And pFileNum > O Then
        Open LogFilePath For Append As #pFileNum
        pFileOpen = True
    End If
End Sub
Sub CloseLogFile()
    If pFileOpen And pFileNum > O Then
        Close #pFileNum
        pFileOpen = False
    End If
End Sub
Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    If LogMessageLevel <> LoggingLevels.NoAssignment Then
        If LogMessageLevel >= Me.ConduitLoggingLevel Then
            If pFileOpen Then
                Print #pFileNum, FormatMessage(pMessageTemplate, LogMessagelLevel, Message)
            Else
                OpenLogFile
                Print #pFileNum, FormatMessage(pMessageTemplate, LogMessagelLevel, Message)
                CloseLogFile
            End If
        End If
    End If
End Sub
Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Dim cnt As Long
    Dim logFileIsOpen As Boolean
    
    If UBound(LogMessageLevels) + 1 < LBound(LogMessageLevels) + EntryCount And UBound(Messages) + 1 < LBound(Messages) + EntryCount Then Err.Raise 440
    
    logFileIsOpen = Me.LogFileOpen
    
    If Not logFileIsOpen Then OpenLogFile
    For cnt = LBound(LogMessageLevels) To LBound(LogMessageLevels) + (EntryCount - 1)
        If LogMessageLevels(cnt) <> LoggingLevels.NoAssignment Then
            If LogMessageLevels(cnt) >= Me.ConduitLoggingLevel Then
                Print #pFileNum, FormatMessage(pMessageTemplate, LogMessageLevels(cnt), Messages(cnt))
            End If
        End If
    Next
    If Not logFileIsOpen Then CloseLogFile
End Sub
Private Function FormatMessage(MessageTemplate As String, LoggingLevel As LoggingLevels, Message As String)
    FormatMessage = UCase(MessageTemplate)
    FormatMessage = Replace(FormatMessage, "#LOGGINGLEVEL#", pLoggingLevelNames(LoggingLevel))
    FormatMessage = Replace(FormatMessage, "#DATE#", DateTime.Date)
    FormatMessage = Replace(FormatMessage, "#TIME#", DateTime.Time)
    FormatMessage = Replace(FormatMessage, "#TIMESTAMP#", DateTime.Now)
    FormatMessage = Replace(FormatMessage, "#MESSAGE#", Message)
End Function

Private Function FileNameFromPath(strFullPath As String) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Private Function FolderFromPath(strFullPath As String) As String
    FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
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
        Do While Len(Dir(BuildPath(folderPath, tempLogFileName))) <> 0 And n <= 999
            tempLogFileName = baseName & "_" & Format(n, "000") & "." & extName
            n = n + 1
        Loop
        
        If n >= 999 Then
            Err.Raise 999, "New File Name", "Cannot create log file. Over 1000 instances exist of " & fileName & " exists in " & LogFolderPath & "."
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

Private Sub Class_Initialize()
    pMessageTemplate = "[#LOGGINGLEVEL#] [#DATE# #TIME#]: #MESSAGE#"
    Set pParentLoggers = New Collection
    
    Set pLoggingLevelNames = New Collection
    pDefaultLoggingLevel = Verbose
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
End Sub

Private Sub Class_Terminate()
    If pFileOpen Then CloseLogFile
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




