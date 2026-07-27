VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_ListBox"
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
''This log conduit channels LogArtifacts to Immediate Window.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements ILogConduit

Dim pParentEvoLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pListBox As MSForms.ListBox
Dim pLoggingLevelNames As Collection

'*******************************************************************************
' Parent Loggers that log messages are channeled from.
'*******************************************************************************
Public Property Get ParentEvoLoggers() As cEvoLoggers
    Set ParentEvoLoggers = pParentEvoLoggers
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
' Initializes properties of the conduit. Meant for when adding to a cEvoLogger.
'*******************************************************************************
Private Sub InitConduit(ParentEvoLogger As cEvoLogger)
    pParentEvoLoggers.Add ParentEvoLogger, ParentEvoLogger.LoggerName
End Sub

'*******************************************************************************
' Initializes properties of the conduit.
'*******************************************************************************
Sub Init(Listbx As Object, ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Information)
    Set pListBox = Listbx
    pConduitName = ConduitName
    
    Me.ConduitLoggingLevel = ConduitLoggingLevel
End Sub
'*******************************************************************************
' Channels a single LogArtifact to its destination.
'*******************************************************************************
Private Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
    If LogArtifact.LoggingLevel >= Me.ConduitLoggingLevel Then
        pListBox.AddItem LogArtifact.FormatArtifact(pMessageTemplate)
    End If
    ''
End Sub

'*******************************************************************************
' Channels an array of LogArtifacts to their destinations.
'*******************************************************************************
Private Sub ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    Dim cnt As Long
    
    For cnt = LBound(LogArtifacts) To UBound(LogArtifacts)
        If LogArtifacts(cnt).LoggingLevel >= Me.ConduitLoggingLevel Then
            pListBox.AddItem LogArtifacts(cnt).FormatArtifact(pMessageTemplate)
        End If
    Next
    Call AutoSizeListBoxColumns(pListBox, 1)
End Sub

Private Sub Class_Initialize()
    pMessageTemplate = "[#LOGGINGLEVEL#] [#DATE# #TIME#]: #MESSAGE#"
    Set pParentEvoLoggers = New cEvoLoggers
    
    Set pLoggingLevelNames = New Collection
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Status", "Status"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
End Sub

Private Sub ListboxAutoColumnwidth(ListBox As MSForms.ListBox)
    Dim uf As UserForm
    Dim TextBox_Dummy As Control
    Dim colCnt As Integer
    Dim rowCnt As Integer
    Dim maxColwidth As Double
    Dim colwidthsArr() As String
    ReDim colwidthsArr(0 To ListBox.ColumnCount - 1)
    
    On Error Resume Next
    Set uf = GetParentForm(ListBox)
    Set TextBox_Dummy = uf.Controls("txtBoxDummy")
    On Error GoTo 0
    
    If TextBox_Dummy Is Nothing Then
        Set TextBox_Dummy = Me.Controls.Add("Forms.TextBox.1", "txtBoxDummy", False)
        TextBox_Dummy.AutoSize = True
    End If
    TextBox_Dummy.Font.size = ListBox.Font.size
    TextBox_Dummy.Font.Name = ListBox.Font.Name
    
    For colCnt = 0 To ListBox.ColumnCount - 1
        maxColwidth = 0
        For rowCnt = 0 To ListBox.ListCount - 1
            TextBox_Dummy.Text = ListBox.List(rowCnt, colCnt)
            If TextBox_Dummy.Width > maxColwWidth Then maxColwidth = TextBox_Dummy.Width
        Next
        colwidthsArr(colCnt) = CStr(maxColwidth + 5)
    Next
    
    ListBox.ColumnWidths = Join(colwidthsArr, ";")
End Sub
Private Function GetParentForm(ctrl As Object) As UserForm
    Dim par As Object
    Set par = ctrl
    
    Do While Not TypeOf par Is MSForms.UserForm
        Set par = par.Parent
    Loop
    
    Set GetParentForm = par
End Function
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



