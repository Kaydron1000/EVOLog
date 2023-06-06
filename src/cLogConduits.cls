VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Sub Add(ByVal Item As Object, Optional key, Optional before, Optional after)
    If IsMissing(key) Then
        If IsMissing(before) Then
            myCustomCollection.Add Item:=Item, before:=before
        ElseIf IsMissing(after) Then
            myCustomCollection.Add Item:=Item, after:=after
        Else
            myCustomCollection.Add Item:=Item
        End If
    ElseIf Not IsMissing(key) Then
        If IsMissing(before) Then
            myCustomCollection.Add Item:=Item, key:=key, before:=before
        ElseIf IsMissing(after) Then
            myCustomCollection.Add Item:=Item, key:=key, after:=after
        Else
            myCustomCollection.Add Item:=Item, key:=key
        End If
        myCustomCollectionKeys.Add Item:=key, key:=key
    End If
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

