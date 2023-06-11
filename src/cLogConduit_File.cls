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
Implements ILogConduit

Private Const constConduitFileExtension As String = "txt"

Dim pParentLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pLoggingLevelNames As Collection

Dim pFileNum As Integer
Dim pFileOpen As Boolean
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

'*******************************************************************************
' The status if log file is open.
'*******************************************************************************
Property Get LogFileOpen() As Boolean
    LogFileOpen = pFileOpen
End Property

'*******************************************************************************
' Folder path the log file exists in.
'*******************************************************************************
Property Get LogFolderPath() As String
    LogFolderPath = pLogFolderPath
End Property

'*******************************************************************************
' The full file path the log file exists at.
'*******************************************************************************
Property Get LogFilePath() As String
    LogFilePath = pLogFilePath
End Property

'*******************************************************************************
' The file name of the log file.
'*******************************************************************************
Property Get LogFileName() As String
    LogFileName = pLogFileName
End Property

'*******************************************************************************
' Initializes properties of the conduit. Meant for when adding to a cEvoLogger.
'*******************************************************************************
Private Sub InitConduit(ParentLogger As cLogger)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
End Sub

'*******************************************************************************
' Initializes properties of the conduit. Initializes everything for the
' channeling to a file except for the parent cEvoLogger object.
'*******************************************************************************
Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Verbose, Optional LogFileName As String = "{DATETIME}_{CONDUITNAME}", _
         Optional LogFolderPath As String = "LogOutput", Optional FileAccess As FileAccesses = FileAccesses.Append)
    
    Dim fileExists As Boolean
    Dim fileName As String
    Dim folderPath As String
    
    folderPath = LogFolderPath
    fileName = LogFileName
    ''Checking if LogFolderPath is a valid Folder
    If Dir(BuildPath(folderPath, "\"), vbDirectory) <> "." Then Err.Raise 76
    
    If fileName = "" Then fileName = LogSinkName & "." & constSinkFileExtension
    
    fileName = Format(DateTime.Now(), "yyyymmdd_hhmmss_") & fileName
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

'*******************************************************************************
' Open log file manually for append. To close after this command use CloseLogFile.
'*******************************************************************************
Sub OpenLogFile()
    If Not pFileOpen And pFileNum > O Then
        Open LogFilePath For Append As #pFileNum
        pFileOpen = True
    End If
End Sub

'*******************************************************************************
' Closes log file if it is open.
'*******************************************************************************
Sub CloseLogFile()
    If pFileOpen And pFileNum > O Then
        Close #pFileNum
        pFileOpen = False
    End If
End Sub

'*******************************************************************************
' Channels a single LogArtifact to its destination.
'*******************************************************************************
Private Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
    Dim logFileIsOpen As Boolean
    
    logFileIsOpen = Me.LogFileOpen
    
    If Not logFileIsOpen Then OpenLogFile
    If LogArtifact.LoggingLevel >= Me.ConduitLoggingLevel Then
        Print #pFileNum, LogArtifact.FormatArtifact(pMessageTemplate)
    End If
    If Not logFileIsOpen Then CloseLogFile
End Sub

'*******************************************************************************
' Channels an array of LogArtifacts to their destinations.
'*******************************************************************************
Private Sub ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    Dim cnt As Long
    Dim logFileIsOpen As Boolean
    
    logFileIsOpen = Me.LogFileOpen
    
    If Not logFileIsOpen Then OpenLogFile
    For cnt = LBound(LogArtifacts) To UBound(LogArtifacts)
        If LogArtifacts(cnt).LoggingLevel >= Me.ConduitLoggingLevel Then
            Print #pFileNum, LogArtifacts(cnt).FormatArtifact(pMessageTemplate)
        End If
    Next
    If Not logFileIsOpen Then CloseLogFile
End Sub

'*******************************************************************************
' Gets the file name from a file path.
'*******************************************************************************
Private Static Function FileNameFromPath(strFullPath As String) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

'*******************************************************************************
' Gets the directory path from a file path.
'*******************************************************************************
Private Static Function FolderFromPath(strFullPath As String) As String
    FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function

'*******************************************************************************
' Gets a unique file name in the specified folder based on the given filename.
'*******************************************************************************
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
        If extStart = 0 Then fileName = fileName & "." & SinkFileExtension
        
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

'*******************************************************************************
' Class initialize method
'*******************************************************************************
Private Sub Class_Initialize()
    pMessageTemplate = "[#LOGGINGLEVEL#] [#DATE# #TIME#]: #MESSAGE#"
    Set pParentLoggers = New cEvoLoggers
    
    Set pLoggingLevelNames = New Collection
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
    If pFileOpen Then CloseLogFile
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
