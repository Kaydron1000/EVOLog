VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduits"
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
''Custom collection to store and return ILogConduit types.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private myCustomCollection As Collection
Private myCustomCollectionKeys As Collection


Private Sub Class_Initialize()
    Set myCustomCollection = New Collection
    Set myCustomCollectionKeys = New Collection
End Sub

Public Sub Class_Terminate()
    Set myCustomCollection = Nothing
End Sub

Public Sub Add(ByVal Item As Object, Optional Key As Variant, _
               Optional Before As Variant, Optional After As Variant)
    myCustomCollection.Add Item:=Item, Key:=Key, Before:=Before, After:=After
    myCustomCollectionKeys.Add Item:=Key, Key:=Key
End Sub

Public Sub AddItems(ParamArray arr() As Variant)
    Dim v As Variant
    For Each v In arr
        myCustomCollection.Add v
    Next
End Sub

Public Sub Remove(Index As Variant)
    myCustomCollection.Remove Index
End Sub

Public Property Get Item(ByVal Index As Variant) As ILogConduit
Attribute Item.VB_UserMemId = 0
    Set Item = myCustomCollection.Item(Index)
End Property

Public Property Get Count() As Long
    Count = myCustomCollection.Count
End Property

Public Property Get Keys() As Collection
    Set Keys = myCustomCollectionKeys
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = myCustomCollection.[_NewEnum]
End Property

