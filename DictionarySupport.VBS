Option Explicit

'====================================================================
' If key is not found when changing an item, a new key is created 
' with the specified newitem. If key is not found when attempting 
' to return an existing item, a new key is created and its 
' corresponding item is left empty.
' 
' С этим и боремся...
'====================================================================
Private Function GetSafeItem(ByRef Dic, ByVal Key, ByVal DefaultItem)
  If Dic.Exists(Key) Then
    GetSafeItem = Dic(Key)
  Else
    GetSafeItem = DefaultItem
  End If
End Function
