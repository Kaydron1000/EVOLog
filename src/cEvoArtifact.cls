VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvoArtifact"
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

Option Explicit

Public pLoggingLevelNames As Collection

Public Charged As Boolean
Public Message As String
Public TimeStamp As Date
Public LoggingLevel As LoggingLevels


Public Static Function LoggingLevelNames() As Collection
    Me.Charged = False
End Function


Public Function FormatArtifact(MessageTemplate As String, Optional TimeStampFormatString As Variant = "yyyymmdd_hhmmss", _
                               Optional DateFormatString As Variant = "mm/dd/yy", Optional TimeFormatString As Variant = "hh:mm:ss") As String
    FormatArtifact = UCase(MessageTemplate)
    FormatArtifact = Replace(FormatArtifact, "#LOGGINGLEVEL#", pLoggingLevelNames(LoggingLevel))
    FormatArtifact = Replace(FormatArtifact, "#DATE#", Format(TimeStamp.Date, DateFormatString))
    FormatArtifact = Replace(FormatArtifact, "#TIME#", Format(TimeStamp.Time, TimeFormatString))
    FormatArtifact = Replace(FormatArtifact, "#TIMESTAMP#", Format(TimeStamp, TimeStampFormatString))
    FormatArtifact = Replace(FormatArtifact, "#MESSAGE#", Message)
    
End Function

Private Sub Class_Initialize()
    Charged = False
End Sub
