VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStringBuilder"
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

Private pBuffer() As String
Private pCurrentIndex As Long
Private pMaxIndex As Long

'*******************************************************************************
' Appends the given Text to this StringBuilder
'*******************************************************************************
Public Sub Append(Text As String)

    pCurrentIndex = pCurrentIndex + 1

    If pCurrentIndex > pMaxIndex Then
        pMaxIndex = 2 * pMaxIndex
        ReDim Preserve pBuffer(1 To pMaxIndex)
    End If
    pBuffer(pCurrentIndex) = Text

End Sub

'*******************************************************************************
' Returns the text in this StringBuilder
'*******************************************************************************
Public Function ToString(Optional Separator As String = vbNullString) As String

    If pCurrentIndex > 0 Then
        ReDim Preserve pBuffer(1 To pCurrentIndex)
        pMaxIndex = pCurrentIndex
        ToString = Join(MyBuffer, Separator)
    End If

End Function

'*******************************************************************************
' Class initialize method
'*******************************************************************************
Private Sub Class_Initialize()
    pCurrentIndex = 0
    pMaxIndex = 16
    ReDim pBuffer(1 To pMaxIndex)
End Sub

