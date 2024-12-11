VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueToAny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Common.Dictionary")
'@IgnoreModule ProcedureNotUsed,ParameterCanBeByVal,ParameterNotUsed
'@NoIndent
'@Interface
Option Explicit
' *** Header ********************************************************************
' ObjectName: IVarValue
' Version:    1.0
' BuildDate:  11/16/2024
' --- Description ---------------------------------------------------------------
'  (Implements) > Interface for VarValue
'
'  VBA7 offers LongPtr and LongLong
'   - LongPtr is intended for pointers and handles.
'     LongPtr resolves to Long in 32-bit versions of Office
'     LongPtr resolves to LongLong in 64-bit versions of Office
'     However, its use in this context is unlikely > ignore via this interface
'   - LongLong is a signed 64-bit integer data type that is only available on 64-bit versions of Office.
'     Bounds: -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
'   - Decimal data type is only available as variant ( > ignore via this interface )
'     but allows to store 96-bit (12-byte) unsigned integers ( +/-79,228,162,514,264,337,593,543,950,335 )
'     the 8 byte (scaled integer) sister "Currency" should be big enough and
'     can handle -922,337,203,685,477.5808 to 922,337,203,685,477.5807
' *******************************************************************************

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
End Property

Public Property Get IsAllocated() As Boolean
End Property

Public Property Get VarType() As VBA.VbVarType
End Property

Public Property Get ToString(Optional ByVal DefaultValue As String = vbNullString) As String
End Property

Public Property Get ToLong(Optional ByVal DefaultValue As Long = 0) As Long
End Property

#If Win64 Then
Public Property Get ToLongLong(Optional ByVal DefaultValue As LongLong = 0) As LongLong
End Property
#End If

Public Property Get ToDouble(Optional ByVal DefaultValue As Double = 0) As Double
End Property

Public Property Get ToCurrency(Optional ByVal DefaultValue As Currency = 0) As Currency
End Property

Public Property Get ToBoolean(Optional ByVal DefaultValue As Boolean = False) As Boolean
End Property

Public Property Get ToDate(Optional ByVal DefaultValue As Date = 2958465) As Date 'MaxDate = DateSerial(9999,12,31) > 2958465
End Property

Public Property Get ToAny(Optional ByVal TargetType As VBA.VbVarType = vbVariant, Optional ByVal DefaultValue As Variant) As Variant
End Property

Public Function GetValueString(ByRef ReturnValue As String, Optional ByVal DefaultValue As String = vbNullString) As Boolean
End Function

Public Function GetValueLong(ByRef ReturnValue As Long, Optional ByVal DefaultValue As Long = 0) As Boolean
End Function

#If Win64 Then
Public Function GetValueLongLong(ByRef ReturnValue As LongLong, Optional ByVal DefaultValue As LongLong = 0) As Boolean
End Function
#End If

Public Function GetValueDouble(ByRef ReturnValue As Double, Optional ByVal DefaultValue As Double = 0) As Boolean
End Function

Public Function GetValueCurrency(ByRef ReturnValue As Currency, Optional ByVal DefaultValue As Currency = 0) As Boolean
End Function

Public Function GetValueDate(ByRef ReturnValue As Date, Optional ByVal DefaultValue As Date = 2958465) As Boolean
End Function

Public Function GetValueBoolean(ByRef ReturnValue As Boolean, Optional ByVal DefaultValue As Boolean = False) As Boolean
End Function
