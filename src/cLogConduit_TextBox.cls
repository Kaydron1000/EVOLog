VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_TextBox"
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
''This log conduit channels log messages to a MSForms.Textbox.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements ILogConduit

Dim pParentLoggers As cEvoLoggers
Dim pLoggingLevelNames As Collection
Dim pDefaultLoggingLevel As LoggingLevels
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String
Dim pTextBox As MSForms.TextBox

Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLoggers = pParentLoggers
End Property

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

Sub InitConduit(ParentLogger As cEvoLogger, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
      
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub
Sub Init(TextBox As MSForms.TextBox, ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = NoAssignment)
    Set pTextBox = TextBox
    Me.ConduitName = ConduitName
    
    If ConduitLoggingLevel = NoAssignment Then
        Me.ConduitLoggingLevel = pDefaultLoggingLevel
    Else
        Me.ConduitLoggingLevel = ConduitLoggingLevel
    End If
End Sub

Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    If LogMessageLevel <> LoggingLevels.NoAssignment Then
        If LogMessageLevel >= Me.ConduitLoggingLevel Then
            pTextBox.Text = pTextBox.Text & FormatMessage(pMessageTemplate, LogMessagelLevel, Message) & vbNewLine
        End If
    End If
End Sub
Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Dim cnt As Long
    Dim sendStrg As String
    Dim sb As cStringBuilder
    
    Set sb = New cStringBuilder
    If UBound(LogMessageLevels) + 1 < LBound(LogMessageLevels) + EntryCount And UBound(Messages) + 1 < LBound(Messages) + EntryCount Then Err.Raise 440
    
    sb.Append pTextBox.Text
    For cnt = LBound(LogMessageLevels) To LBound(LogMessageLevels) + (EntryCount - 1)
        If LogMessageLevels(cnt) <> LoggingLevels.NoAssignment Then
            If LogMessageLevels(cnt) >= Me.ConduitLoggingLevel Then
                sb.Append FormatMessage(pMessageTemplate, LogMessageLevels(cnt), Messages(cnt)) & vbNewLine
            End If
        End If
    Next
    
    pTextBox.Text = sb.toString
End Sub
Private Function FormatMessage(MessageTemplate As String, LoggingLevel As LoggingLevels, Message As String)
    FormatMessage = UCase(MessageTemplate)
    FormatMessage = Replace(FormatMessage, "#LOGGINGLEVEL#", pLoggingLevelNames(LoggingLevel))
    FormatMessage = Replace(FormatMessage, "#DATE#", DateTime.Date)
    FormatMessage = Replace(FormatMessage, "#TIME#", DateTime.Time)
    FormatMessage = Replace(FormatMessage, "#TIMESTAMP#", DateTime.Now)
    FormatMessage = Replace(FormatMessage, "#MESSAGE#", Message)
End Function
Private Sub Class_Initialize()
    pMessageTemplate = "[$#LOGGINGLEVEL#] [#DATE# $TIME#]: $MESSAGE#"
    Set pParentLoggers = New cEvoLoggers
    
    Set pLoggingLevelNames = New Collection
    pDefaultLoggingLevel = Information
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
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


