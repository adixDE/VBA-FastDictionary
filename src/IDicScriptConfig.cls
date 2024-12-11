VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDicScriptConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Common.Dictionary")
'@IgnoreModule ProcedureNotUsed
'@NoIndent
'@Interface
Option Explicit
' *** Header ********************************************************************
' ObjectName: IDicScriptConfig
' Version:    1.1
' BuildDate:  11/01/2024
' --- Description ---------------------------------------------------------------
'  (Implements) > Interface for DicScript (to offer encapsulate config properties )
' *******************************************************************************

Public Property Get CaseInsensitiveKeys() As Boolean: End Property
Public Property Let CaseInsensitiveKeys(ByVal EnableCaseInsensitiveKeys As Boolean): End Property

Public Property Get DisableIndexReference() As Boolean: End Property
Public Property Let DisableIndexReference(ByVal NewValue As Boolean): End Property

Public Property Get ItemNullValue() As Variant: End Property
Public Property Set ItemNullValue(ByVal NullValue As Variant): End Property
Public Property Let ItemNullValue(ByVal NullValue As Variant): End Property

Public Property Get DefaultKey() As Variant: End Property
Public Property Set DefaultKey(ByVal NewValue As Variant): End Property
Public Property Let DefaultKey(ByVal NewValue As Variant): End Property

Public Property Get ChildPathSplitChar() As String: End Property
Public Property Let ChildPathSplitChar(ByVal SplitChar As String): End Property
