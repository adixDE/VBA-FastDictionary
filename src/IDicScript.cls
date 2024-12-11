VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDicScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Common.Dictionary")
'@IgnoreModule ProcedureNotUsed,ParameterCanBeByVal
'@NoIndent
'@Interface
Option Explicit
' *** Header ********************************************************************
' ObjectName: IDicScriptRead
' Version:    1.0
' BuildDate:  11/01/2024
' --- Description ---------------------------------------------------------------
'  (Implements) > Interface for DicScript (to offer read only property access)
'  Similar to DicScript you can toggle ToAny usage
'  Purpose was/is to have a (min) read access interface for a config class
'  Usage: setup a DicScript and assign this to a IDicScriptRead class
'  This allows access to the core "Get" properties and the config DisableIndexReference
' *******************************************************************************

Public Property Get Exists(Optional ByVal KeyOrOneBasedIndex As Variant, Optional ByRef ProperCaseItemKey As Variant) As Boolean
End Property

Public Property Get Count() As Long
End Property

Public Property Get Index(Optional ByVal ItemKey As Variant) As Long
End Property

Public Property Get Item(Optional ByVal KeyOrOneBasedIndex As Variant) As Variant
End Property

Public Property Get Child(Optional ByVal ReferencePath As Variant) As Variant
End Property

#If EnableToAny = 1 Then
Public Property Get ItemToAny(Optional ByVal KeyOrOneBasedIndex As Variant) As IValueToAny
End Property

Public Property Get ChildToAny(Optional ByVal ReferencePath As Variant) As IValueToAny
End Property
#End If

Public Property Get Key(ByVal OneBasedIndex As Long) As Variant
End Property

Public Property Get Keys(Optional ByVal LikeCriteria As Variant) As Variant
End Property

Public Property Get StringKeys(Optional ByVal LikeCriteria As Variant) As String()
End Property

Public Function GetData(ByRef ReturnData As Variant, Optional ByVal ReturnItems As Long = 0, Optional ByVal CreateNonMultiArray As Boolean = False, Optional ByRef RowCount As Long, Optional ByRef ErrorText As String) As Boolean
End Function
