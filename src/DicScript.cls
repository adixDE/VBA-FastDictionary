VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DicScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Common.Dictionary")
'@IgnoreModule FunctionReturnValueDiscarded,FunctionReturnValueAlwaysDiscarded,ProcedureNotUsed,ParameterCanBeByVal
'@NoIndent
Option Explicit
' *** Header ********************************************************************
' ObjectName: DicScript
' Version:    2.0.0
' BuildDate:  11/30/2024
' --- Description ---------------------------------------------------------------
'  Rebuild: based on the 'Dictionary' class from Cristian Buse which is
'  Mac OS compatible, quite fast and dont require a static reference to scripting.dictionary
'   Goal is ongoing to have a associative indexed array (1 Based)
'   I'm not a friend of 'For Each Enum' and NewEnum() As IUnknown
'   > therefore this approach offer access via KeyOrOneBasedIndex
'
'  Core logic:
'   * keys are optional (but sensible) - enclosed types are not supported: vbArray, vbError(10), vbDataObject(13) or NullValues
'     missing/blank keys will be set to a (unique) string key (always "Item" & Counter - not a full GUID for performance reasons and compatibility)
'   * .Item could be everything
'   * .Item (Let/Set) create a new item once the Key could not be resolved
'   * KeyOrOneBasedIndex approach (however Index is always 1based)
'   * .Child allows to access nested items via a path reference (Key1\ChildKey\.. )
'
'   This class handle selected limitations of 'Dictionary' like
'    * Change the compare mode once data is assigned
'    * access via Key or OneBasedIndex
'    * Sort (by Keys or Items) (fast natural sort without OS api requirements)
'    * offer item access via child enum
'
'  Properties (Config):         encapsulated (via IDicScriptConfig) below .Config
'   .CaseInsensitiveStringKeys  (Boolean) Toogle comparemode (Binary or Text)
'   .DisableIndexReference      (Boolean) disable numeric key reference once the passed value is not a key
'   .ItemNullValue              get/let/set a default value (default is vbNullString)
'   .DefaultKey                 optional default key (used once a blank key is passed)
'   .ChildPathSplitChar         Used by .Child to determine the childpath once string (default is \)
'
'  Properties (GetOnly):
'   .Count                      item count (long)
'   .Exists(KeyOrOneBasedIndex,ProperCaseKey)   Will return Boolean - but likewise the proper case/type key.
'   -- rather for compatibility ..
'   .Key(OneBasedIndex)     Get a variant Key
'   .Keys                   Get all Keys within a variant array (zero based) - optional Filtered (LikeCriteria with wildcard support)
'   .StringKeys             Get all Keys within a string array (zero based) -  optional Filtered (LikeCriteria with wildcard support)
'   .Index(ItemKey)         get the index of a key
'
'  Properties:
'   .Item                   will return/assign the .Item by IndexOrKey (Otherwise: ItemNullValue)
'                           Get will consider .Config.DefaultKey | Let/Set will create an item once IndexOrKey could not be resolved
'  Methods:
'   .Add                    allows to add new data
'   .Remove                 remove a item
'   .Clear                  reset the data
'   .Sort                   sort by Keys or Items (always: "natural" - sort)
'   -- (Once : #EnableToAny = 1)
'   .ItemToAny(Optional ByVal KeyOrOneBasedIndex As Variant) As IValueToAny
'   .ChildToAny(Optional ByVal ReferencePath As Variant) As IValueToAny
'
'  Public auxilary functions:
'   .GetData                Helper to get a variant array (Optional 2-dimensional array)
'   --                      ReturnItems  (Long: Default = 0)  1 will return Items, -1 will return both (Keys and Items) Else return Keys
'   .InitByData             dynamic way to init the dictionary (Keys or Items) by ParmArray Data
'   .InitByTextData         init the dictionary (Keys or Items) by text / seperator
'   .InitByDataPair         init Keys and Item data by a variant array (must be a 2D array with 2 columns)
'   .InitByDicScript        clone an dicscript into 'me'
'
'  Optional enable the Interface (IDicScript) (via EnableReadInterface = 1) to offer Read-Access to
'   .Exists(Optional ByVal KeyOrOneBasedIndex As Variant, Optional ByRef ProperCaseItemKey As Variant)
'   .Count
'   .Index(Optional ByVal ItemKey As Variant)
'   .Item(Optional ByVal KeyOrOneBasedIndex As Variant)
'   .Child(Optional ByVal ReferencePath As Variant)
'   .Key(ByVal OneBasedIndex As Long)
'   .Keys(Optional ByVal LikeCriteria As String = vbNullString)
'   .StringKeys(Optional ByVal LikeCriteria As String = vbNullString)
'   .GetData
'   -- (Once : #EnableToAny = 1)
'   .ItemToAny(Optional ByVal KeyOrOneBasedIndex As Variant) As IValueToAny
'   .ChildToAny(Optional ByVal ReferencePath As Variant) As IValueToAny
' *******************************************************************************

'Info: > The compiler variable 'EnableToAny' is below considered (See AppDeclare)
#Const EnableReadInterface = 1

Implements IDicScript   'manual disable (comment out) once EnableReadInterface <> 1 (to prevent compiler error)
Implements IDicScriptConfig

Private Type TModel
  Col                       As Dictionary
  '--- Config Properties
  CaseInsensitiveKeys       As Boolean
  DisableIndexReference     As Boolean
  DefaultKey                As Variant
  DefaultKeyIsAllocated     As Boolean
  ItemNullValue             As Variant
  ItemNullValueIsAllocated  As Boolean
  ChildPathSplitChar        As String
End Type

Private this As TModel

Private Sub Class_Initialize()
  this.CaseInsensitiveKeys = True
  this.ChildPathSplitChar = "\"
  Reset
End Sub

Public Property Get Config() As IDicScriptConfig
  Set Config = Me
End Property

Public Property Get Count() As Long: Count = this.Col.Count: End Property
Public Property Get Exists(Optional ByVal KeyOrOneBasedIndex As Variant, Optional ByRef ProperCaseItemKey As Variant) As Boolean: Exists = GetItemKey(KeyOrOneBasedIndex, ProperCaseItemKey, Not this.DisableIndexReference): End Property

''@Enumerator
'Public Function NewEnum() As IUnknown
'  #If Not Mac Then
'    Set NewEnum = CallByName(this.Col, "_NewEnum", VbMethod)
'  #End If
'End Function

'Item (GET/LET/SET): ----------------------------------------------------------------------
'@DefaultMember
Public Property Get Item(Optional ByRef KeyOrOneBasedIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0

  Dim ItemKey As Variant

  If IsMissing(KeyOrOneBasedIndex) Then
    If this.DefaultKeyIsAllocated Then
      If GetItemKey(this.DefaultKey, ItemKey, False) Then
        AssignValue Item, this.Col.Item(ItemKey)
        Exit Property
      End If
    End If
    If this.ItemNullValueIsAllocated Then
      AssignValue Item, this.ItemNullValue
    Else
      Item = vbNullString
    End If
  Else
    If GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference) Then
      AssignValue Item, this.Col.Item(ItemKey)
    Else
      If this.ItemNullValueIsAllocated Then
        AssignValue Item, this.ItemNullValue
      Else
        Item = vbNullString
      End If
    End If
  End If

End Property

Public Property Let Item(Optional ByRef KeyOrOneBasedIndex As Variant, ByVal ItemData As Variant)

  Dim ItemKey As Variant
  If IsMissing(KeyOrOneBasedIndex) Then
    Add , ItemData
  Else
    If GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference) Then
      If IsObject(ItemData) Then Set this.Col.Item(ItemKey) = ItemData Else this.Col.Item(ItemKey) = ItemData
    Else
      Add KeyOrOneBasedIndex, ItemData
    End If
  End If

End Property

Public Property Set Item(Optional ByRef KeyOrOneBasedIndex As Variant, ByVal ItemData As Variant)

  Dim ItemKey As Variant
  If IsMissing(KeyOrOneBasedIndex) Then
    Add , ItemData
  Else
    If GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference) Then
      If IsObject(ItemData) Then Set this.Col.Item(ItemKey) = ItemData Else this.Col.Item(ItemKey) = ItemData
    Else
      Add KeyOrOneBasedIndex, ItemData
    End If
  End If

End Property

'Child (GET/LET/SET): ---------------------------------------------------------------------
Public Property Get Child(Optional ByVal ReferencePath As Variant) As Variant
  
  Dim ItemKey As Variant
  Dim PathArray As Variant
  Dim EnumItem As DicScript
  Dim RowLBound As Long: Dim RowUbound As Long
  
  If IsMissing(ReferencePath) Then
    If this.DefaultKeyIsAllocated Then
      If GetItemKey(this.DefaultKey, ItemKey, False) Then
        If Not IsObject(this.Col.Item(ItemKey)) Then Child = this.Col.Item(ItemKey) Else Set Child = this.Col.Item(ItemKey)
        Exit Property
      End If
    End If
  Else
    If Not VarType(ReferencePath) = vbObject Then
      'A - Direct
      If GetItemKey(ReferencePath, ItemKey, Not this.DisableIndexReference) Then
        If IsObject(this.Col.Item(ItemKey)) Then Set Child = this.Col.Item(ItemKey) Else Child = this.Col.Item(ItemKey)
        Exit Property
      End If
    
      'B - PathArray resolvable?
      If ResolvePathArray(ReferencePath, PathArray) Then
        If GetArraySize(PathArray, , , RowLBound, RowUbound) Then
          Set EnumItem = Me
          If EnumGetItem(EnumItem, PathArray, RowLBound, RowUbound) Then
            ItemKey = PathArray(RowUbound)
            If IsObject(EnumItem.Item(ItemKey)) Then Set Child = EnumItem.Item(ItemKey) Else Child = EnumItem.Item(ItemKey)
            Exit Property
          End If
        End If
      End If
    
    End If
  End If
  
  'At this point i know that the item could not be resolved.
  If this.ItemNullValueIsAllocated Then
    If Not IsObject(this.ItemNullValue) Then Child = this.ItemNullValue Else Set Child = this.ItemNullValue
  Else
    Child = vbNullString
  End If
  
End Property

Public Property Let Child(ByVal ReferencePath As Variant, ByVal ItemData As Variant)
  EnumSetItem ReferencePath, ItemData
End Property

Public Property Set Child(ByVal ReferencePath As Variant, ByVal ItemData As Variant)
  EnumSetItem ReferencePath, ItemData
End Property

'Add: -------------------------------------------------------------------------------------
Public Sub Add(Optional ByRef KeyOrOneBasedIndex As Variant, Optional ByRef ItemData As Variant)

  Dim ItemKey As Variant
  Dim BlankKey As Boolean: Dim ItemExists As Boolean

  BlankKey = IsMissing(KeyOrOneBasedIndex)
  If Not BlankKey Then BlankKey = KeyIsBlank(KeyOrOneBasedIndex)
    
  If BlankKey Then
    If Not IsMissing(ItemData) Then
      ItemKey = GetUniqueKey
      this.Col.Add ItemKey, ItemData
    End If
  Else
    ItemExists = GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference)
    If Not ItemExists Then
      If IsMissing(ItemData) Then
        If this.ItemNullValueIsAllocated Then
          this.Col.Add KeyOrOneBasedIndex, this.ItemNullValue
        Else
          this.Col.Add KeyOrOneBasedIndex, vbNullString
        End If
      Else
        this.Col.Add KeyOrOneBasedIndex, ItemData
      End If
    Else
      If IsObject(ItemData) Then
        Set this.Col.Item(ItemKey) = ItemData
      Else
        this.Col.Item(ItemKey) = ItemData
      End If
    End If
  End If

End Sub

Public Property Get Index(Optional ByVal ItemKey As Variant) As Long
  Dim ItemProperKey As Variant
  If this.Col.Count > 0 Then
    If GetItemKey(ItemKey, ItemProperKey, False) Then
      Index = this.Col.Index(ItemProperKey)
    End If
  End If
End Property

Public Property Get Key(ByVal OneBasedIndex As Long) As Variant
  Dim ItemIndex As Long
  If this.Col.Count > 0 Then
    ItemIndex = OneBasedIndex
    If ItemIndex < 1 Or ItemIndex > this.Col.Count Then ItemIndex = 0
  End If
  If ItemIndex > 0 Then
    AssignValue Key, this.Col.KeyByIndex(ItemIndex)
  Else
    Key = vbNullString
  End If
End Property

Public Property Get Keys(Optional ByVal LikeCriteria As Variant) As Variant

  Dim vKeys As Variant: Dim ReturnValue As Variant
  Dim Counter As Long: Dim NextItem As Long
  Dim CompareTypeRating As Long: Dim ItemTypeRating As Long
  
  If this.Col.Count = 0 Then
    Keys = Array()
  Else
    If IsMissing(LikeCriteria) Then
      Keys = this.Col.Keys()
    Else
      vKeys = this.Col.Keys()
      CompareTypeRating = GetDataTypeRank(LikeCriteria)
      ReDim ReturnValue(LBound(vKeys) To UBound(vKeys)) As Variant
      NextItem = LBound(vKeys) - 1
      
      For Counter = LBound(vKeys) To UBound(vKeys)
        ItemTypeRating = GetDataTypeRank(vKeys(Counter))
        If ItemTypeRating = CompareTypeRating Then
          If ItemTypeRating > 7 Then
            If (vKeys(Counter) Like LikeCriteria) Then
              NextItem = NextItem + 1
              ReturnValue(NextItem) = vKeys(Counter)
            End If
          End If
        End If
      Next
      
      If NextItem = LBound(vKeys) - 1 Then
        Keys = Array()
      Else
        If NextItem < UBound(vKeys) Then
          ReDim Preserve ReturnValue(LBound(vKeys) To NextItem) As Variant
        End If
        Keys = ReturnValue
      End If
    End If
  End If

End Property

Public Property Get StringKeys(Optional ByVal LikeCriteria As Variant) As String()
  
  Dim vKeys As Variant
  Dim ReturnValue() As String
  Dim Counter As Long: Dim NextItem As Long
  Dim CompareTypeRating As Long: Dim ItemTypeRating As Long
  
  If this.Col.Count = 0 Then
    '@Ignore EmptyStringLiteral
    StringKeys = Split("", vbNullString)
  Else
    vKeys = this.Col.Keys()
    If IsMissing(LikeCriteria) Then
      ReDim ReturnValue(0 To UBound(vKeys)) As String
      For Counter = LBound(vKeys) To UBound(vKeys)
        ItemTypeRating = GetDataTypeRank(vKeys(Counter))
        If ItemTypeRating > 7 Then
          ReturnValue(Counter) = CStr(vKeys(Counter))
        Else
          ReturnValue(Counter) = CStr(Counter + 1)
        End If
      Next
      StringKeys = ReturnValue
    Else
      CompareTypeRating = GetDataTypeRank(LikeCriteria)
      If CompareTypeRating < 8 Then
        '@Ignore EmptyStringLiteral
        StringKeys = Split("", vbNullString)
      Else
        ReDim ReturnValue(0 To UBound(vKeys)) As String
        NextItem = LBound(vKeys) - 1
        
        For Counter = LBound(vKeys) To UBound(vKeys)
          ItemTypeRating = GetDataTypeRank(vKeys(Counter))
          If ItemTypeRating = CompareTypeRating Then
            If ItemTypeRating > 7 Then
              If (vKeys(Counter) Like LikeCriteria) Then
                NextItem = NextItem + 1
                ReturnValue(NextItem) = vKeys(Counter)
              End If
            End If
          End If
        Next
      
        If NextItem = LBound(vKeys) - 1 Then
          '@Ignore EmptyStringLiteral
          StringKeys = Split("", vbNullString)
        Else
          If NextItem < UBound(vKeys) Then
            ReDim Preserve ReturnValue(LBound(vKeys) To NextItem) As String
          End If
          StringKeys = ReturnValue
        End If
      End If
    End If
  End If
  
End Property

Public Sub Remove(ByVal KeyOrOneBasedIndex As Variant)
  Dim ItemKey As Variant
  If GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference) Then
    this.Col.Remove ItemKey
  End If
End Sub

Public Sub Clear(): Clear: End Sub

Public Function GetData(ByRef ReturnData As Variant, Optional ByVal ReturnItems As Long = 0, Optional ByVal CreateNonMultiArray As Boolean = False, Optional ByRef RowCount As Long, Optional ByRef ErrorText As String) As Boolean

  Dim Counter As Long: Dim ItemKey As Variant

  On Error GoTo Handle_Error
  ErrorText = vbNullString
  RowCount = this.Col.Count

  If RowCount = 0 Then ErrorText = "No data stored within this dictionary": GoTo Handle_Error
  If ReturnItems < -1 Or ReturnItems > 1 Then ErrorText = "Invalid Trigger ReturnItems (" & ReturnItems & ") - use ABS(Boolean) or -1 for both": GoTo Handle_Error

  If ReturnItems = -1 Then
    ReDim ReturnData(1 To RowCount, 1 To 2) As Variant
  Else
    If CreateNonMultiArray Then ReDim ReturnData(1 To RowCount) As Variant Else ReDim ReturnData(1 To RowCount, 1 To 1) As Variant
  End If
  
  For Counter = 1 To RowCount
    ItemKey = this.Col.KeyByIndex(Counter)
    If ReturnItems = 1 Then
      If CreateNonMultiArray Then
        AssignValue ReturnData(Counter, 1), this.Col.Item(ItemKey)
      Else
        AssignValue ReturnData(Counter), this.Col.Item(ItemKey)
      End If
    Else
      If ReturnItems = -1 Then
        AssignValue ReturnData(Counter, 1), ItemKey
        AssignValue ReturnData(Counter, 2), this.Col.Item(ItemKey)
      Else
        If CreateNonMultiArray Then
          AssignValue ReturnData(Counter, 1), ItemKey
        Else
          AssignValue ReturnData(Counter), ItemKey
        End If
      End If
    End If
  Next

  GetData = CBool(1)
  Exit Function

Handle_Error:
  If Err.Number Then Err.Clear
End Function

#If EnableToAny = 1 Then
  
Public Property Get ItemToAny(Optional ByVal KeyOrOneBasedIndex As Variant) As IValueToAny

  Dim ItemKey As Variant
  Dim ReturnValue As ValueToAny
  Set ReturnValue = New ValueToAny

  If IsMissing(KeyOrOneBasedIndex) Then
    If this.DefaultKeyIsAllocated Then
      If GetItemKey(this.DefaultKey, ItemKey, False) Then
        If IsObject(this.Col.Item(ItemKey)) Then Set ReturnValue.Item = this.Col.Item(ItemKey) Else ReturnValue.Item = this.Col.Item(ItemKey)
        Set ItemToAny = ReturnValue
        Exit Property
      End If
    End If
  Else
    If GetItemKey(KeyOrOneBasedIndex, ItemKey, Not this.DisableIndexReference) Then
      If IsObject(this.Col.Item(ItemKey)) Then Set ReturnValue.Item = this.Col.Item(ItemKey) Else ReturnValue.Item = this.Col.Item(ItemKey)
      Set ItemToAny = ReturnValue
      Exit Property
    End If
  End If

  If this.ItemNullValueIsAllocated Then
    If IsObject(this.ItemNullValue) Then Set ReturnValue.Item = this.ItemNullValue Else ReturnValue.Item = this.ItemNullValue
  End If
  Set ItemToAny = ReturnValue

End Property
  
Public Property Get ChildToAny(Optional ByVal ReferencePath As Variant) As IValueToAny

  Dim ReturnValue As ValueToAny
  Dim ItemKey As Variant
  Dim PathArray As Variant
  Dim EnumItem As DicScript
  Dim RowLBound As Long: Dim RowUbound As Long

  Set ReturnValue = New ValueToAny

  If IsMissing(ReferencePath) Then
    If this.DefaultKeyIsAllocated Then
      If GetItemKey(this.DefaultKey, ItemKey, False) Then
        If IsObject(this.Col.Item(ItemKey)) Then Set ReturnValue.Item = this.Col.Item(ItemKey) Else ReturnValue.Item = this.Col.Item(ItemKey)
        Set ChildToAny = ReturnValue
        Exit Property
      End If
    End If
  Else
    If Not VarType(ReferencePath) = vbObject Then
      'A - Direct
      If GetItemKey(ReferencePath, ItemKey, Not this.DisableIndexReference) Then
        If IsObject(this.Col.Item(ItemKey)) Then Set ReturnValue.Item = this.Col.Item(ItemKey) Else ReturnValue.Item = this.Col.Item(ItemKey)
        Set ChildToAny = ReturnValue
        Exit Property
      End If

      'B - PathArray resolvable?
      If ResolvePathArray(ReferencePath, PathArray) Then
        If GetArraySize(PathArray, , , RowLBound, RowUbound) Then
          Set EnumItem = Me
          If EnumGetItem(EnumItem, PathArray, RowLBound, RowUbound) Then
            If IsObject(EnumItem.Item(PathArray(RowUbound))) Then Set ReturnValue.Item = EnumItem.Item(PathArray(RowUbound)) Else ReturnValue.Item = EnumItem.Item(PathArray(RowUbound))
            Set ChildToAny = ReturnValue
            Exit Property
          End If
        End If
      End If
    End If
  End If

  'At this point i know that the item could not be resolved.
  If this.ItemNullValueIsAllocated Then
    If IsObject(this.ItemNullValue) Then Set ReturnValue.Item = this.ItemNullValue Else ReturnValue.Item = this.ItemNullValue
  End If
  Set ChildToAny = ReturnValue

End Property

#End If

Public Function Sort(Optional ByVal SortByKeys As Boolean = True, Optional ByVal SortDescending As Boolean = False) As Boolean
  
  Dim SortData() As Variant: Dim IndexData() As Long: Dim SortOrder() As Long
  Dim ItemIndex As Long: Dim ItemCount As Long: Dim ItemDataRank As Long
  Dim SourceIndex As Long: Dim TargetIndex As Long
  Dim vKeys As Variant
  Dim NoChange As Boolean
  Dim Clone As Dictionary
  
  On Error GoTo Handle_Error

  ItemCount = this.Col.Count
  If ItemCount < 2 Then
    Sort = CBool(1)
    Exit Function
  End If

  ReDim SortData(1 To ItemCount, 1 To 2) As Variant
  ReDim IndexData(1 To ItemCount, 1 To 2) As Long
  ReDim SortOrder(1 To ItemCount) As Long
  
  vKeys = this.Col.Keys
  For ItemIndex = 1 To ItemCount
    If SortByKeys Then
      SortData(ItemIndex, 1) = GetSortValue(vKeys(ItemIndex - 1), ItemIndex, ItemDataRank)
    Else
      SortData(ItemIndex, 1) = GetSortValue(this.Col.Item(vKeys(ItemIndex - 1)), ItemIndex, ItemDataRank)
      If ItemDataRank > 6 Then
        SortData(ItemIndex, 1) = CStr(SortData(ItemIndex, 1)) & "-" & ItemIndex
        ItemDataRank = 8
      End If
    End If
    SortData(ItemIndex, 2) = ItemDataRank
    IndexData(ItemIndex, 1) = ItemIndex: IndexData(ItemIndex, 2) = ItemIndex
  Next
  
  LocalSort SortData, IndexData, 1, ItemCount
  
  NoChange = True
  For ItemIndex = 1 To ItemCount
    If SortDescending Then
      TargetIndex = ItemCount - IndexData(ItemIndex, 1) + 1
    Else
      TargetIndex = IndexData(ItemIndex, 1)
    End If
    If Not TargetIndex = IndexData(ItemIndex, 2) Then
      If NoChange Then NoChange = False
    End If
    SortOrder(TargetIndex) = IndexData(ItemIndex, 2)
  Next
  
  If NoChange Then
    Erase SortData: Erase IndexData: Erase SortOrder
    Sort = CBool(1)
    Exit Function
  End If
  
  Set Clone = New Dictionary
  
  Clone.CompareMode = this.Col.CompareMode
  Clone.PredictCount ItemCount
  
  For ItemIndex = 1 To ItemCount
    SourceIndex = SortOrder(ItemIndex)
    Clone.Add vKeys(SourceIndex - 1), this.Col.Item(vKeys(SourceIndex - 1))
  Next
  
  Set this.Col = Clone
  
  Erase SortData: Erase IndexData: Erase SortOrder: Set Clone = Nothing
  Sort = CBool(1)
  Exit Function

Handle_Error:
  Err.Raise 5, , Err.Description
  On Error GoTo 0
End Function

Public Function InitByTextData(ByVal TextData As String, Optional ByVal ItemDivChar As String = ";", Optional ByVal InitAsKeys As Boolean = True, Optional ByVal NullKey As String = "(NULL)", Optional ByRef ErrorText As String, Optional ByVal ResetMe As Boolean = True) As Boolean

  Dim StringValue As String
  Dim DataArray() As String
  Dim Counter As Long

  On Error GoTo Handle_Error
  ErrorText = vbNullString

  If TextData = vbNullString Then InitByTextData = CBool(1): Exit Function
  If ItemDivChar = vbNullString Then ErrorText = "Missing delimeter (ItemDivChar)": GoTo Handle_Error

  If ResetMe Then Clear

  Counter = InStr(1, TextData, ItemDivChar)
  If Counter > 0 Then
    DataArray = Split(TextData, ItemDivChar)

    For Counter = LBound(DataArray) To UBound(DataArray)
      If InitAsKeys Then
        StringValue = CStr(DataArray(Counter))
        If Trim$(StringValue) = vbNullString Then StringValue = NullKey

        If Not StringValue = vbNullString Then
          If Not Exists(StringValue) Then
            Add StringValue, vbNullString
          End If
        End If
      Else
        StringValue = CStr(DataArray(Counter))
        If Not Trim$(StringValue) = vbNullString Then
          Add vbNullString, StringValue
        End If
      End If
    Next
  Else
    If InitAsKeys Then
      Add TextData, vbNullString
    Else
      Add vbNullString, TextData
    End If
  End If

  InitByTextData = CBool(1)
  Exit Function

Handle_Error:
  If ErrorText = vbNullString Then ErrorText = GetError
End Function

Public Function InitByData(Optional ByVal InitAsKeys As Boolean = True, Optional ByVal ResetMe As Boolean = False, Optional ByRef VarData As Variant) As Boolean
  If ResetMe Then Me.Clear
  InitByData = InitByParamArray(InitAsKeys, VarData)
End Function

Public Function InitByDataPair(Optional ByRef VarData As Variant, Optional ByRef ErrorText As String) As Boolean

  Dim RowItem As Long: Dim ColCount As Long: Dim RowLBound As Long: Dim RowUbound As Long: Dim ColLBound As Long: Dim ColUBound As Long
  Dim IsMulti As Boolean

  On Error GoTo Handle_Error

  If Not GetArraySize(VarData, IsMulti, , RowLBound, RowUbound, ColCount, ColLBound, ColUBound, ErrorText) Then GoTo Handle_Error

  If Not ColCount = 2 Then
    ErrorText = "Invalid SourceData - "
    If Not IsMulti Then ErrorText = ErrorText & "the passed vector is a not two dimensional" Else ErrorText = ErrorText & "the passed vector contains not sufficient columns (" & ColCount & " instead of 2)"
    GoTo Handle_Error
  End If

  For RowItem = RowLBound To RowUbound
    Add VarData(RowItem, ColLBound), VarData(RowItem, ColLBound + 1)
  Next

  InitByDataPair = True
  Exit Function

Handle_Error:
  If ErrorText = vbNullString Then ErrorText = GetError
End Function

Public Function InitByDicScript(ByRef CloneDic As DicScript, Optional ByVal ResetMe As Boolean = True) As Boolean
  Dim Counter As Long: Dim ResetIndex As Boolean
  If ResetMe Then Me.Clear
  If Not CloneDic Is Nothing Then
    If CloneDic.Config.DisableIndexReference Then
      CloneDic.Config.DisableIndexReference = False: ResetIndex = True
    End If
    For Counter = 1 To CloneDic.Count
      Me.Add CloneDic.Key(Counter), CloneDic.Item(Counter)
    Next
    If ResetIndex Then CloneDic.Config.DisableIndexReference = True
    InitByDicScript = True
  End If
End Function

'Private (config): --- (start) ------------------------------------------------------------

'CaseInsensitiveKeys
Private Property Get IDicScriptConfig_CaseInsensitiveKeys() As Boolean: IDicScriptConfig_CaseInsensitiveKeys = this.CaseInsensitiveKeys: End Property
Private Property Let IDicScriptConfig_CaseInsensitiveKeys(ByVal NewValue As Boolean)
  
  'Dictionary does not allow to change the CompareMode once data is allocated
  If this.Col.CompareMode = Abs(NewValue) Then Exit Property    'vbBinaryCompare = 0 / vbTextCompare = 1
  If this.Col.Count = 0 Then
    this.Col.CompareMode = Abs(NewValue)
    this.CaseInsensitiveKeys = NewValue
    Exit Property
  End If
  
  'rebuild / clone
  Dim Clone As Dictionary
  Dim vKeys As Variant
  Dim Counter As Long
  
  Set Clone = New Dictionary
  Clone.CompareMode = Abs(NewValue)
  vKeys = this.Col.Keys
  For Counter = LBound(vKeys) To UBound(vKeys)
    Clone.Add vKeys(Counter), this.Col.Item(vKeys(Counter))
  Next
  Set this.Col = Clone
  this.CaseInsensitiveKeys = NewValue

End Property

'DisableIndexReference
Private Property Get IDicScriptConfig_DisableIndexReference() As Boolean: IDicScriptConfig_DisableIndexReference = this.DisableIndexReference: End Property
Private Property Let IDicScriptConfig_DisableIndexReference(ByVal NewValue As Boolean): this.DisableIndexReference = NewValue: End Property

'DefaultKey
Private Property Get IDicScriptConfig_DefaultKey() As Variant
  If this.DefaultKeyIsAllocated Then
    If IsObject(this.DefaultKey) Then Set IDicScriptConfig_DefaultKey = this.DefaultKey Else IDicScriptConfig_DefaultKey = this.DefaultKey
  Else
    IDicScriptConfig_DefaultKey = vbNullString
  End If
End Property
Private Property Let IDicScriptConfig_DefaultKey(ByVal NewValue As Variant): AssignDefaultKey NewValue: End Property
Private Property Set IDicScriptConfig_DefaultKey(ByVal NewValue As Variant): AssignDefaultKey NewValue: End Property

'ItemNullValue
Private Property Get IDicScriptConfig_ItemNullValue() As Variant
  If this.ItemNullValueIsAllocated Then
    If IsObject(this.ItemNullValue) Then Set IDicScriptConfig_ItemNullValue = this.ItemNullValue Else IDicScriptConfig_ItemNullValue = this.ItemNullValue
  Else
    IDicScriptConfig_ItemNullValue = vbNullString
  End If
End Property
Private Property Set IDicScriptConfig_ItemNullValue(ByVal NullValue As Variant): AssignItemNullValue NullValue: End Property
Private Property Let IDicScriptConfig_ItemNullValue(ByVal NullValue As Variant): AssignItemNullValue NullValue: End Property

'ChildPathSplitChar
Private Property Get IDicScriptConfig_ChildPathSplitChar() As String: IDicScriptConfig_ChildPathSplitChar = this.ChildPathSplitChar: End Property
Private Property Let IDicScriptConfig_ChildPathSplitChar(ByVal SplitChar As String)
  If SplitChar = vbNullString Then this.ChildPathSplitChar = "\" Else this.ChildPathSplitChar = SplitChar
End Property

'Private (config): --- (end) --------------------------------------------------------------


#If EnableReadInterface = 1 Then
'Private (Read Interface via IDicScript): --- (start) ---------------------------------
  Private Property Get IDicScript_Exists(Optional ByVal KeyOrOneBasedIndex As Variant, Optional ProperCaseItemKey As Variant) As Boolean
    IDicScript_Exists = Exists(KeyOrOneBasedIndex, ProperCaseItemKey)
  End Property
  
  Private Property Get IDicScript_Count() As Long
    IDicScript_Count = Count()
  End Property
  
  Private Property Get IDicScript_Index(Optional ByVal ItemKey As Variant) As Long
    IDicScript_Index = Index(ItemKey)
  End Property

  Private Property Get IDicScript_Item(Optional ByVal KeyOrOneBasedIndex As Variant) As Variant
    AssignValue IDicScript_Item, Item(KeyOrOneBasedIndex)
  End Property

  Private Property Get IDicScript_Child(Optional ByVal ReferencePath As Variant) As Variant
    AssignValue IDicScript_Child, Child(ReferencePath)
  End Property

  Private Property Get IDicScript_Key(ByVal OneBasedIndex As Long) As Variant
    AssignValue IDicScript_Key, Key(OneBasedIndex)
  End Property

  Private Property Get IDicScript_Keys(Optional ByVal LikeCriteria As Variant) As Variant
    IDicScript_Keys = Keys(LikeCriteria)
  End Property

  Private Property Get IDicScript_StringKeys(Optional ByVal LikeCriteria As Variant) As String()
    IDicScript_StringKeys = StringKeys(LikeCriteria)
  End Property

  Private Function IDicScript_GetData(ByRef ReturnData As Variant, Optional ByVal ReturnItems As Long = 0, Optional ByVal CreateNonMultiArray As Boolean = False, Optional ByRef RowCount As Long, Optional ByRef ErrorText As String) As Boolean
    IDicScript_GetData = GetData(ReturnData, ReturnItems, CreateNonMultiArray, RowCount, ErrorText)
  End Function

  #If EnableToAny = 1 Then
    Private Property Get IDicScript_ItemToAny(Optional ByVal KeyOrOneBasedIndex As Variant) As IValueToAny
      Set IDicScript_ItemToAny = ItemToAny(KeyOrOneBasedIndex)
    End Property

    Private Property Get IDicScript_ChildToAny(Optional ByVal ReferencePath As Variant) As IValueToAny
      Set IDicScript_ChildToAny = ChildToAny(ReferencePath)
    End Property
  #End If

'Private (Read Interface via IDicScript): --- (end) -----------------------------------
#End If


'Private Helpers --------------------------------------------------------------------------

Private Sub Reset()
  Set this.Col = New Dictionary
  If this.CaseInsensitiveKeys Then this.Col.CompareMode = 1
End Sub

Private Sub AssignDefaultKey(ByRef NewValue As Variant)
  If VariantIsValid(NewValue, True) Then
    If IsObject(NewValue) Then Set this.DefaultKey = NewValue Else this.DefaultKey = NewValue
    this.DefaultKeyIsAllocated = True
  Else
    this.DefaultKey = vbNullString
    this.DefaultKeyIsAllocated = False
  End If
End Sub

Private Sub AssignItemNullValue(ByVal NewValue As Variant)
  If VariantIsValid(NewValue, True) Then
    If IsObject(NewValue) Then Set this.ItemNullValue = NewValue Else this.ItemNullValue = NewValue
    this.DefaultKeyIsAllocated = True
  Else
    this.DefaultKey = vbNullString
    this.DefaultKeyIsAllocated = False
  End If
End Sub

Private Function GetUniqueKey() As String
  Dim ReturnValue As String
  Dim Counter As Long
  For Counter = this.Col.Count + 1 To (this.Col.Count * 2) + 1
    ReturnValue = "item" & Counter
    If Not this.Col.Exists(ReturnValue) Then
      GetUniqueKey = ReturnValue
      Exit Function
    End If
  Next
End Function

Private Sub AssignValue(ByRef TargetValue As Variant, ByRef SourceValue As Variant)
  If IsObject(SourceValue) Then Set TargetValue = SourceValue Else TargetValue = SourceValue
End Sub

Private Function GetItemKey(ByVal KeyOrOneBasedIndex As Variant, Optional ByRef ProperCaseItemKey As Variant, Optional ByVal EnableIndexMatch As Boolean = True) As Boolean

  Dim ItemIndex As Long: Dim ValueType As Long
  Dim ItemKey As String
  
  ItemKey = vbNullString

  If this.Col.Count = 0 Then GoTo Handle_Exit
  If IsArray(KeyOrOneBasedIndex) Then GoTo Handle_Exit    'Arrays are not supported as Key
  If IsNull(KeyOrOneBasedIndex) Then GoTo Handle_Exit
  
  On Error Resume Next
  If this.Col.Exists(KeyOrOneBasedIndex) Then
    If Not IsMissing(ProperCaseItemKey) Then
      ProperCaseItemKey = this.Col.KeyByIndex(this.Col.Index(KeyOrOneBasedIndex))
    End If
    GetItemKey = CBool(1)
    Exit Function
  End If
     
  If EnableIndexMatch Then
    ValueType = VarType(KeyOrOneBasedIndex)
    ItemIndex = 0
    
    'Ignore vbObject(9) from scratch
    'Consider directly: vbInteger(2),vbLong(3),vbSingle(4),vbDouble(5),vbCurrency(6),vbDecimal(14),vbLongLong(20) (once they fit into a long and are within bounds)
    'Check vbString(8) and vbVariant(12)
    'But ignore the rest [ including vbBoolean(11),vbByte(17) and vbDate(7) ]
    
    If Not ValueType = 9 Then
      Select Case ValueType
        Case 2 To 6, 14, 20
          If KeyOrOneBasedIndex > -2147483648# Then
            If KeyOrOneBasedIndex < 2147483647# Then
              If ValueType > 3 Then
                If Len(CStr(KeyOrOneBasedIndex)) = Len(CStr(CLng(KeyOrOneBasedIndex))) Then
                  ItemIndex = CLng(KeyOrOneBasedIndex)
                End If
              Else
                ItemIndex = CLng(KeyOrOneBasedIndex)
              End If
              If ItemIndex < 1 Or ItemIndex > this.Col.Count Then ItemIndex = 0
            End If
          End If
        Case 8          'vbString
          If Not Trim$(KeyOrOneBasedIndex) = vbNullString Then
            If IsNumeric(KeyOrOneBasedIndex) Then
              If ItemKey > -2147483648# Then
                If ItemKey < 2147483647# Then
                  If Len(Trim$(KeyOrOneBasedIndex)) = Len(CStr(CLng(KeyOrOneBasedIndex))) Then
                    ItemIndex = CLng(KeyOrOneBasedIndex)
                    If ItemIndex < 1 Or ItemIndex > this.Col.Count Then ItemIndex = 0
                  End If
                End If
              End If
            End If
          End If
        Case 12         'vbVariant(12)
          If Not IsObject(KeyOrOneBasedIndex) Then
            If IsNumeric(KeyOrOneBasedIndex) Then
              If ItemKey > -2147483648# Then
                If ItemKey < 2147483647# Then
                  If Len(Trim$(KeyOrOneBasedIndex)) = Len(CStr(CLng(KeyOrOneBasedIndex))) Then
                    ItemIndex = CLng(KeyOrOneBasedIndex)
                    If ItemIndex < 1 Or ItemIndex > this.Col.Count Then ItemIndex = 0
                  End If
                End If
              End If
            End If
          End If
      End Select
      
      If ItemIndex > 0 Then
        ProperCaseItemKey = this.Col.KeyByIndex(ItemIndex)
        GetItemKey = CBool(1)
      End If
    End If
  End If
  
Handle_Exit:
  If Err.Number Then Err.Clear
  On Error GoTo 0
End Function

Private Function VariantIsValid(ByRef DataValue As Variant, Optional ByVal SkipVBNullstring As Boolean) As Boolean

  Dim VType As Long

  'catch invalid types like: vbObject (vbDataObject, vbUserDefinedType), vbEmpty or vbNull, vbError and each type of vbArray
  If IsArray(DataValue) Then Exit Function
  If IsObject(DataValue) Then
    If DataValue Is Nothing Then
      Exit Function
    Else
      VariantIsValid = True: Exit Function
    End If
  Else
    If IsNull(DataValue) Then Exit Function        'catch vbEmpty or vbNull
  End If

  VType = VarType(DataValue)
  Select Case VType
    Case 10, 13 'vbError (10), vbDataObject (13)
      Exit Function
    Case 12     'vbVariant
      On Error Resume Next
        VariantIsValid = Not CBool(CStr(DataValue) = vbNullString)
        If Err.Number Then Err.Clear
      On Error GoTo 0
      Exit Function
    Case Else
      If VType > 1 Then
        If VType < 21 Then
          'could be valid (type shold be:  vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle or vbString)
          VariantIsValid = CBool(1)
          If SkipVBNullstring Then
            If VType = 8 Then
              VariantIsValid = Not CBool(DataValue = vbNullString)
            End If
          End If
        End If
      End If
  End Select

End Function

Private Function KeyIsBlank(ByVal KeyOrOneBasedIndex As Variant) As Boolean

  Dim ValueType As Long
  Dim ReturnValue As Boolean

  On Error GoTo Handle_Exit
  
  ValueType = VarType(KeyOrOneBasedIndex): ReturnValue = False
  If ValueType And vbArray Then GoTo Handle_Exit  'Arrays are not supported as "Key" within a Scripting.Dictionary > Return False
  
  If ValueType = vbString Then    'vbString(8)
    If Trim$(KeyOrOneBasedIndex) = vbNullString Then ReturnValue = True: GoTo Handle_Exit
  Else
    If IsNull(KeyOrOneBasedIndex) Then ReturnValue = True: GoTo Handle_Exit
    If IsObject(KeyOrOneBasedIndex) Then
      'could impact 'CaseInsensitiveKeys' match .. but allow it (once not Nothing)
      ReturnValue = KeyOrOneBasedIndex Is Nothing
      GoTo Handle_Exit
    End If
    
    Select Case ValueType
      'allow all number types
      Case 2 To 7, 11, 14, 17, 20  'vbInteger(2),vbLong(3),vbSingle(4),vbDouble(5),vbCurrency(6),vbDate(7),vbBoolean(11),vbDecimal(14),vbByte(17),vbLongLong(20)
        ReturnValue = False: GoTo Handle_Exit
      Case 12                      'vbVariant(12)
        If Trim$(CStr(KeyOrOneBasedIndex)) = vbNullString Then ReturnValue = True: GoTo Handle_Exit
      Case Else   'includes: vbEmpty(0),vbNull(1),vbError(10),vbDataObject(13),vbUserDefinedType(36)
        ReturnValue = True: GoTo Handle_Exit
    End Select
  End If

Handle_Exit:
  KeyIsBlank = ReturnValue
  If Err.Number Then Err.Clear
  On Error GoTo 0
End Function

Private Function GetDataTypeRank(ByRef VariantValue As Variant) As Long
  'Empty = 1  | UDT = 2 | Object = 3 | Array = 4 | Null = 5 | Error = 6 | Boolean = 7 | Text = 8 | Number = 9
  If IsObject(VariantValue) Then
    GetDataTypeRank = 3
    Exit Function
  End If
  Select Case VarType(VariantValue)
    Case vbNull
      GetDataTypeRank = 5
    Case vbEmpty
      GetDataTypeRank = 1
    Case vbError
      GetDataTypeRank = 6
    Case vbBoolean
      GetDataTypeRank = 7
    Case vbString
      GetDataTypeRank = 8   'Optional adjust Rank for numbers stored as text .. but handle this during a compare
    Case vbByte, vbInteger, vbLong, vbLongLong 'Integers
      GetDataTypeRank = 9
    Case vbCurrency, vbDecimal, vbDouble, vbSingle, vbDate 'Decimal-point
      GetDataTypeRank = 9
    Case vbArray To vbArray + vbUserDefinedType
      GetDataTypeRank = 4
    Case vbUserDefinedType
      GetDataTypeRank = 2
    Case vbDataObject
      GetDataTypeRank = 3
  End Select
End Function

Private Function ResolvePathArray(ByRef SourceValue As Variant, ByRef ReturnPathArray As Variant) As Boolean
  
  'Even you can simplify the collecting via DataArrayCol .. avoid this dependency here
  Dim vData As Variant
  Dim TextPath As String
  Dim RowLBound As Long: Dim RowUbound As Long: Dim RowCount As Long: Dim RowIndex As Long
  Dim ColLBound As Long: Dim ColUBound As Long: Dim ColCount As Long: Dim ColIndex As Long: Dim DataIndex As Long
  Dim IsMulti As Boolean

  If IsArray(SourceValue) Then
    If GetArraySize(SourceValue, IsMulti, RowCount, RowLBound, RowUbound, ColCount, ColLBound, ColUBound) Then
      ReDim vData(RowLBound To RowUbound) As Variant
      DataIndex = RowLBound - 1
      If IsMulti Then
        For RowIndex = RowLBound To RowUbound
          For ColIndex = ColLBound To ColUBound
            If VariantIsValid(SourceValue(RowIndex, ColIndex), False) Then
              DataIndex = DataIndex + 1
              vData(DataIndex) = SourceValue(RowIndex, ColIndex)
            End If
          Next
        Next
      Else
        For RowIndex = RowLBound To RowUbound
          If VariantIsValid(SourceValue(RowIndex), False) Then
            DataIndex = DataIndex + 1
            vData(DataIndex) = SourceValue(RowIndex)
          End If
        Next
      End If
      If DataIndex > RowLBound - 1 Then
        ReDim ReturnPathArray(RowLBound To DataIndex) As Variant
        For RowIndex = RowLBound To DataIndex
          ReturnPathArray(RowIndex) = vData(RowIndex)
        Next
        ResolvePathArray = True
      End If
    End If
  Else
    If VariantIsValid(SourceValue, False) Then
      TextPath = CStr(SourceValue)
      If InStr(1, TextPath, this.ChildPathSplitChar, vbTextCompare) > 0 Then
        ReturnPathArray = Split(TextPath, this.ChildPathSplitChar, , vbBinaryCompare)
        ResolvePathArray = True
      End If
    End If
  End If

End Function

Private Function EnumGetItem(ByRef Container As DicScript, ByRef PathArray As Variant, ByVal ItemLevel As Long, ByVal LevelUBound As Long, Optional ByRef ErrorText As String) As Boolean

  Dim ItemSourceKey As Variant: ItemSourceKey = vbNullString
  Dim ItemKey As Variant
  
  ItemKey = PathArray(ItemLevel)
  If ItemLevel < LevelUBound Then
    If Container.Exists(ItemKey, ItemSourceKey) Then
      If TypeName(Container.Item(ItemSourceKey)) = "DicScript" Then
        Set Container = Container.Item(ItemSourceKey)
        EnumGetItem = EnumGetItem(Container, PathArray, ItemLevel + 1, LevelUBound, ErrorText)
      Else
        ErrorText = "Item on Level(" & ItemLevel & ") is not a DicScript"
      End If
    Else
      ErrorText = "Item on Level(" & ItemLevel & ") does not exist"
    End If
  Else
    If Not Container.Exists(ItemKey, ItemSourceKey) Then
      ErrorText = "Item on Level(" & ItemLevel & ") does not exist"
    Else
      EnumGetItem = True
    End If
  End If
  
End Function

Private Sub EnumSetItem(Optional ByVal ReferencePath As Variant, Optional ByVal ItemData As Variant)
  
  Dim ItemKey As Variant
  Dim vPath As Variant
  Dim RowLBound As Long: Dim RowUbound As Long: Dim RowIndex As Long
  Dim EnumItem As DicScript
  
  If IsMissing(ReferencePath) Then
    Add , ItemData
  Else
    If Not VarType(ReferencePath) = vbObject Then
      'A - Direct
      If GetItemKey(ReferencePath, ItemKey, Not this.DisableIndexReference) Then
        If IsObject(ItemData) Then Set this.Col.Item(ItemKey) = ItemData Else this.Col.Item(ItemKey) = ItemData
        Exit Sub
      End If
      
      'Resolve and create
      If ResolvePathArray(ReferencePath, vPath) Then
        If GetArraySize(vPath, , , RowLBound, RowUbound) Then
          Set EnumItem = Me
          For RowIndex = RowLBound To RowUbound
            
            ItemKey = vPath(RowIndex)
            If Not EnumItem.Exists(ItemKey) Then
              If RowIndex < RowUbound Then
                EnumItem.Add ItemKey, New DicScript
                Set EnumItem = EnumItem.Item(ItemKey)
              Else
                EnumItem.Add ItemKey, ItemData
              End If
            Else
              If RowIndex < RowUbound Then
                If TypeName(EnumItem.Item(ItemKey)) = "DicScript" Then
                  Set EnumItem = EnumItem.Item(ItemKey)
                Else
                  Set EnumItem.Item(ItemKey) = New DicScript
                  Set EnumItem = EnumItem.Item(ItemKey)
                End If
              Else
                If IsObject(ItemData) Then Set EnumItem.Item(ItemKey) = ItemData Else EnumItem.Item(ItemKey) = ItemData
              End If
            End If
          Next
        End If
      End If
      
    End If
  End If
  
End Sub

Private Function GetArraySize(ByRef DataArray As Variant, Optional ByRef IsTwoDimensional As Boolean, _
                              Optional ByRef RowCount As Long, Optional ByRef RowLBound As Long, Optional ByRef RowUbound As Long, _
                              Optional ByRef ColCount As Long, Optional ByRef ColLBound As Long, Optional ByRef ColUBound As Long, _
                              Optional ByRef ErrorText As String) As Boolean

  Dim ReturnValue As Boolean: Dim ArrayIsAllocated As Boolean
  
  On Error Resume Next
  IsTwoDimensional = CBool(0)
  RowCount = -1: RowLBound = -1: RowUbound = -1
  ColCount = -1: ColLBound = -1: ColUBound = -1
  ErrorText = vbNullString
  
  ArrayIsAllocated = Not (IsError(LBound(DataArray))) And IsArray(DataArray) And (LBound(DataArray) <= UBound(DataArray))

  If ArrayIsAllocated Then
    RowUbound = UBound(DataArray)
    RowUbound = UBound(DataArray, 1)
    If RowUbound > -1 Then
      RowLBound = LBound(DataArray)
      RowLBound = LBound(DataArray, 1)
      RowCount = RowUbound - RowLBound + 1
      ColUBound = UBound(DataArray, 2)
      If ColUBound > -1 Then
        IsTwoDimensional = CBool(1)
        ColLBound = LBound(DataArray, 2)
        ColCount = ColUBound - ColLBound + 1
      End If
      ReturnValue = CBool(1)
    End If
  End If
  If Not ReturnValue Then ErrorText = "The passed vector is not an allocated data array!"

  If Err.Number Then Err.Clear
  On Error GoTo 0
  GetArraySize = ReturnValue

End Function

Private Function GetSortValue(ByRef SourceValue As Variant, ByVal ItemIndex As Long, Optional ByRef DataTypeRank As Long) As Variant
  
  'Empty = 1  | UDT = 2 | Object = 3 | Array = 4 | Null = 5 | Error = 6 | Boolean = 7 | Text = 8 | Number = 9
  'Some Types could not be sorted - but handle them as string in order to be sorted by Group at the bottom.
  'For for non-sortable items (<7): apply a prefix ( 10xZ ), mid section is CHR$(64+{SortRank} and the suffix is the ItemIndex
  
  Dim NumberValue As Double
  
  If IsObject(SourceValue) Then
    GetSortValue = CStr("ZZZZZZZZZZ_C_" & ItemIndex)
    Exit Function
  End If
  
  Select Case VarType(SourceValue)
    Case vbNull
      GetSortValue = CStr("ZZZZZZZZZZ_E_" & ItemIndex)
      DataTypeRank = 8
      Exit Function
    Case vbEmpty
      GetSortValue = CStr("ZZZZZZZZZZ_A_" & ItemIndex)
      DataTypeRank = 8
    Case vbError
      GetSortValue = CStr("ZZZZZZZZZZ_F_" & ItemIndex)
      DataTypeRank = 8
    Case vbBoolean
      GetSortValue = Abs(SourceValue)
      DataTypeRank = 7
    Case vbString
      If IsNumeric(SourceValue) Then
        If CastToDouble(SourceValue, NumberValue) Then
          DataTypeRank = 9
          GetSortValue = NumberValue
          Exit Function
        End If
      End If
      DataTypeRank = 8
      GetSortValue = CStr(SourceValue)
    Case vbByte, vbInteger, vbLong, 20    'Integers (vbLongLong = 20)
      DataTypeRank = 9
      GetSortValue = SourceValue
    Case vbCurrency, vbDecimal, vbDouble, vbSingle  'Decimal-point
      DataTypeRank = 9
      GetSortValue = SourceValue
    Case vbDate
      If CastToDouble(SourceValue, NumberValue) Then
        DataTypeRank = 9
        GetSortValue = NumberValue
      End If
    Case vbArray To vbArray + vbUserDefinedType
      GetSortValue = CStr("ZZZZZZZZZZ_D_" & ItemIndex)
      DataTypeRank = 4
    Case vbUserDefinedType
      GetSortValue = CStr("ZZZZZZZZZZ_B_" & ItemIndex)
      DataTypeRank = 2
    Case vbDataObject
      GetSortValue = CStr("ZZZZZZZZZZ_C_" & ItemIndex)
      DataTypeRank = 3
  End Select
  
End Function

Private Function CastToDouble(ByVal SourceValue As Variant, ByRef ReturnDouble As Double) As Boolean
  On Error GoTo Handle_Error
  ReturnDouble = CDbl(SourceValue)
  CastToDouble = True
  Exit Function
  
Handle_Error:
  If Err.Number Then Err.Clear
  On Error GoTo 0
End Function

Private Sub LocalSort(ByRef SourceData As Variant, ByRef IndexData() As Long, Optional ByVal StartIndex As Long, Optional ByVal StopIndex As Long)
  
  Dim IndexLower As Long: Dim IndexUpper As Long: Dim TmpIndex As Long: Dim IndexLBound As Long
  Dim MatchDataIndex As Long
  
  IndexLower = StartIndex: IndexUpper = StopIndex
  
  On Error Resume Next
    '@Ignore AssignmentNotUsed
    IndexLBound = -1
    IndexLBound = UBound(IndexData, 2)
    If Err.Number Then Err.Clear
  On Error GoTo 0

  If IndexLBound = -1 Then
    MatchDataIndex = IndexData((StartIndex + StopIndex) \ 2)
  Else
    MatchDataIndex = IndexData((StartIndex + StopIndex) \ 2, IndexLBound)
  End If
    
  Do While IndexLower < IndexUpper
    If IndexLBound = -1 Then
      'API way (which "could" be way faster - but is slowed down by multiple factors)
      'Do While StrCmpLogicalW(StrPtr(SourceData(IndexData(IndexLower), 1)), StrPtr(SourceData(MatchDataIndex, 1))) = -1 And IndexLower < StopIndex
      '  IndexLower = IndexLower + 1
      'Loop
      'Do While StrCmpLogicalW(StrPtr(SourceData(IndexData(IndexUpper), 1)), StrPtr(SourceData(MatchDataIndex, 1))) = 1 And IndexUpper > StartIndex
      '  IndexUpper = IndexUpper - 1
      'Loop
      Do While CompareValue(SourceData(IndexData(IndexLower), 1), SourceData(IndexData(IndexLower), 2), SourceData(MatchDataIndex, 1), SourceData(MatchDataIndex, 2)) = -1 And IndexLower < StopIndex
        IndexLower = IndexLower + 1
      Loop
      Do While CompareValue(SourceData(IndexData(IndexUpper), 1), SourceData(IndexData(IndexUpper), 2), SourceData(MatchDataIndex, 1), SourceData(MatchDataIndex, 2)) = 1 And IndexUpper > StartIndex
        IndexUpper = IndexUpper - 1
      Loop
      If IndexLower <= IndexUpper Then
        TmpIndex = IndexData(IndexUpper)
        IndexData(IndexUpper) = IndexData(IndexLower)
        IndexData(IndexLower) = TmpIndex
        IndexLower = IndexLower + 1
        IndexUpper = IndexUpper - 1
      End If
    Else
      'API way (which "could" be way faster - but is slowed down by multiple factors)
      'Do While StrCmpLogicalW(StrPtr(SourceData(IndexData(IndexLower, IndexLBound), 1)), StrPtr(SourceData(MatchDataIndex, 1))) = -1 And IndexLower < StopIndex
      '  IndexLower = IndexLower + 1
      'Loop
      'Do While StrCmpLogicalW(StrPtr(SourceData(IndexData(IndexUpper, IndexLBound), 1)), StrPtr(SourceData(MatchDataIndex, 1))) = 1 And IndexUpper > StartIndex
      '  IndexUpper = IndexUpper - 1
      'Loop
      Do While CompareValue(SourceData(IndexData(IndexLower, IndexLBound), 1), SourceData(IndexData(IndexLower, IndexLBound), 2), SourceData(MatchDataIndex, 1), SourceData(MatchDataIndex, 2)) = -1 And IndexLower < StopIndex
        IndexLower = IndexLower + 1
      Loop
      Do While CompareValue(SourceData(IndexData(IndexUpper, IndexLBound), 1), SourceData(IndexData(IndexUpper, IndexLBound), 2), SourceData(MatchDataIndex, 1), SourceData(MatchDataIndex, 2)) = 1 And IndexUpper > StartIndex
        IndexUpper = IndexUpper - 1
      Loop
      If IndexLower <= IndexUpper Then
        TmpIndex = IndexData(IndexUpper, IndexLBound)
        IndexData(IndexUpper, IndexLBound) = IndexData(IndexLower, IndexLBound)
        IndexData(IndexLower, IndexLBound) = TmpIndex
        IndexLower = IndexLower + 1
        IndexUpper = IndexUpper - 1
      End If
    End If
  Loop

  If IndexUpper > StartIndex Then LocalSort SourceData, IndexData, StartIndex, IndexUpper
  If IndexLower < StopIndex Then LocalSort SourceData, IndexData, IndexLower, StopIndex

End Sub

Private Function CompareValue(ByVal Value1 As Variant, ByVal Value1Rank As Variant, ByVal Value2 As Variant, ByVal Value2Rank As Variant) As Long
  
  If Value1Rank = 7 And Value2Rank = 7 Then
    If Value1 = Value2 Then
      CompareValue = 0
    Else
      'Apply: False < True
      If (True Xor Value2) Then CompareValue = 1 Else CompareValue = -1
    End If
    Exit Function
  End If
    
  If Value1Rank = 7 Or Value2Rank = 7 Then
    If Value1Rank = 7 Then CompareValue = 1 Else CompareValue = -1
    Exit Function
  End If
    
  If Value1Rank = 9 And Value2Rank = 9 Then
    If Value1 = Value2 Then
      CompareValue = 0
    Else
      If Value1 < Value2 Then CompareValue = 1 Else CompareValue = -1
    End If
    Exit Function
  End If
  
  'Else - apply (custom) manual/natural compare which will work likewise on a MAC (without the API StrCmpLogicalW)
  CompareValue = NaturalStringCompare(Value1, Value2, True)

End Function

Private Function NaturalStringCompare(ByVal TextValueA As String, ByVal TextValueB As String, Optional ByVal CaseSensitive As Boolean) As Long
  
  'Return: 1 = x > y, -1 = y > x, 0 = x == y
  'Rules: Numbers < Letters. Space < everything
  
  Dim CharPos As Long: Dim TextLenA As Long: Dim TextLenB As Long: Dim NumRet As Long: Dim NumOffset As Long
  Dim CharBlockA As String: Dim CharBlockB As String
  Dim WhiteSpace As String: WhiteSpace = " "
  
  If TextValueA = TextValueB Then NaturalStringCompare = 0: Exit Function
  If TextValueA = vbNullString And Not TextValueB = vbNullString Then NaturalStringCompare = -1: Exit Function
  If Not TextValueA = vbNullString And TextValueB = vbNullString Then NaturalStringCompare = 1: Exit Function
  
  TextLenA = Len(TextValueA)
  TextLenB = Len(TextValueB)
  
  If TextLenB > TextLenA Then
    If StrComp(Left$(TextValueB, TextLenA), TextValueA, Abs(CaseSensitive)) = 0 Then NaturalStringCompare = -1: Exit Function
  End If
  
  For CharPos = 1 To TextLenA
    
    If TextLenB < CharPos Then NaturalStringCompare = 1: Exit Function
    
    CharBlockA = Mid$(TextValueA, CharPos, 1)
    CharBlockB = Mid$(TextValueB, CharPos, 1)
    
    If CharBlockA = WhiteSpace And Not CharBlockB = WhiteSpace Then NaturalStringCompare = -1: Exit Function
    If Not CharBlockA = WhiteSpace And CharBlockB = WhiteSpace Then NaturalStringCompare = 1: Exit Function
    
    If IsNumeric(CharBlockA) Then
      If Not IsNumeric(CharBlockB) Then NaturalStringCompare = -1: Exit Function
      NumRet = CompareNumbers(TextValueA, TextValueB, CharPos, NumOffset)
      If Not NumRet = 0 Then NaturalStringCompare = NumRet: Exit Function
      CharPos = CharPos + NumOffset
    ElseIf IsNumeric(CharBlockB) Then
      NaturalStringCompare = 1: Exit Function
    Else
      NumRet = StrComp(CharBlockA, CharBlockB, Abs(CaseSensitive))
      If Not NumRet = 0 Then NaturalStringCompare = NumRet: Exit Function
    End If
  Next
  
  NaturalStringCompare = 0
  
End Function

Private Function CompareNumbers(ByVal TextValueA As String, ByVal TextValueB As String, ByVal SourceOffset As Long, ByRef ReturnOffset As Long) As Long
  
  Dim NumBlockA As String: Dim NumBlockB As String
  Dim NumValueA As Variant: Dim NumValueB As Variant: Dim LeadingA As Long: Dim LeadingB As Long
  
  'The initial approach (working with long's) could result in a buffer overrun once the textblock exceed 2147483647
  'Therefore: Changed to variant and 'ExtractNumber' cast into the required type (min: long - else Currency)
  
  NumValueA = ExtractNumber(TextValueA, SourceOffset, NumBlockA, LeadingA)
  NumValueB = ExtractNumber(TextValueB, SourceOffset, NumBlockB, LeadingB)
  
  ReturnOffset = Len(NumBlockA)
  
  If NumBlockA = NumBlockB Then CompareNumbers = 0: Exit Function
  
  If NumValueA > NumValueB Then CompareNumbers = 1: Exit Function
  If NumValueA < NumValueB Then CompareNumbers = -1: Exit Function
  
  If NumValueA = NumValueB Then
    ' 001 > 01 > 1
    If LeadingA > LeadingB Then CompareNumbers = -1: Exit Function
    If LeadingA < LeadingB Then CompareNumbers = 1: Exit Function
    CompareNumbers = 0
  End If
  
End Function

Private Function ExtractNumber(ByVal SourceString As String, ByVal Offset As Long, ByRef ReturnBlock As String, ByRef LeadingZeroes As Long) As Variant
    
  Dim CharPos As Long
  Dim ItemChar As String
  Dim NotLeading As Boolean
  
  ReturnBlock = vbNullString: LeadingZeroes = 0
  
  For CharPos = Offset To Len(SourceString)
    ItemChar = Mid$(SourceString, CharPos, 1)
    If IsNumeric(ItemChar) Then
      If ItemChar = "0" Then
        If Not NotLeading Then LeadingZeroes = LeadingZeroes + 1
      Else
        NotLeading = True
      End If
      ReturnBlock = ReturnBlock & ItemChar
    Else
      Exit For
    End If
  Next
  
  If ReturnBlock < 2147483648# Then
    ExtractNumber = CLng(ReturnBlock)
  Else
    'Check for double (8 byte)
    If ReturnBlock < 4.94065645841247E-324 Then
      ExtractNumber = CDbl(ReturnBlock)
    Else
      'check for Currency (8 byte)
      If ReturnBlock < 922337203685478# Then
        ExtractNumber = CCur(ReturnBlock)
      Else
        'check for Decimal (14 byte)
        If ReturnBlock < 7.92281625142643E+28 Then  '79228162514264337593543950335
          ExtractNumber = CDec(ReturnBlock)
        Else
          'crop the zero's and return a string
          If LeadingZeroes > 0 Then
            ExtractNumber = Right$(ReturnBlock, Len(ReturnBlock) - LeadingZeroes)
          Else
            ExtractNumber = ReturnBlock
          End If
        End If
      End If
    End If
  End If
  
End Function

Private Function InitByParamArray(ByVal InitAsKey As Boolean, ParamArray ParmData() As Variant) As Boolean

  Dim Counter1 As Long: Dim Counter2 As Long
  Dim RowItem As Long: Dim ColItem As Long: Dim RowLBound As Long: Dim RowUbound As Long: Dim ColLBound As Long: Dim ColUBound As Long
  Dim SkipEntry As Boolean: Dim IsMulti As Boolean
  Dim VarData As Variant

  For Counter1 = LBound(ParmData) To UBound(ParmData)
    'A - Object?
    If IsObject(ParmData(Counter1)) Then
      If CBool(TypeName(ParmData(Counter1)) = "DicScript") Then
        'consider the lower level only - in order to collect recursive (clone) .. use InitByDicScript
        For Counter2 = 1 To ParmData(Counter1).Count
          If VariantIsValid(ParmData(Counter1).Item(Counter2)) Then
            If InitAsKey Then
              If Not Exists(ParmData(Counter1).Item(Counter2)) Then
                Add ParmData(Counter1).Item(Counter2), vbNullString
              End If
            Else
              Add vbNullString, ParmData(Counter1).Item(Counter2)
            End If
          End If
        Next
      Else
        If Not InitAsKey Then
          Add vbNullString, ParmData(Counter1)
        End If
      End If
    Else
      SkipEntry = CBool(0)
      SkipEntry = IsEmpty(ParmData(Counter1))
      If Not SkipEntry Then SkipEntry = IsNull(ParmData(Counter1))

      If Not SkipEntry Then
        'B - Array?
        If VarType(ParmData(Counter1)) And vbArray Then
          If GetArraySize(ParmData(Counter1), IsMulti, , RowLBound, RowUbound, , ColLBound, ColUBound) Then
            If Not IsMulti Then
              For RowItem = RowLBound To RowUbound
                If VarType(ParmData(Counter1)(RowItem)) And vbArray Then
                  VarData = ParmData(Counter1)(RowItem)
                  InitByParamArray InitAsKey, VarData
                Else
                  If VariantIsValid(ParmData(Counter1)(RowItem)) Then
                    If InitAsKey Then
                      If Not Exists(ParmData(Counter1)(RowItem)) Then
                        Add ParmData(Counter1)(RowItem), vbNullString
                      End If
                    Else
                      Add vbNullString, ParmData(Counter1)(RowItem)
                    End If
                  End If
                End If
              Next
            Else
              For RowItem = RowLBound To RowUbound
                For ColItem = ColLBound To ColUBound
                  If VarType(ParmData(Counter1)(RowItem, ColItem)) And vbArray Then
                    VarData = ParmData(Counter1)(RowItem, ColItem)
                    InitByParamArray InitAsKey, VarData
                  Else
                    If VariantIsValid(ParmData(Counter1)(RowItem, ColItem)) Then
                      If InitAsKey Then
                        If Not Exists(ParmData(Counter1)(RowItem, ColItem)) Then
                          Add ParmData(Counter1)(RowItem, ColItem), vbNullString
                        End If
                      Else
                        Add vbNullString, ParmData(Counter1)(RowItem, ColItem)
                      End If
                    End If
                  End If
                Next
              Next
            End If
          End If
        Else
          If VariantIsValid(ParmData(Counter1)) Then
            If InitAsKey Then
              If Not Exists(ParmData(Counter1)) Then
                Add ParmData(Counter1), vbNullString
              End If
            Else
              Add vbNullString, ParmData(Counter1)
            End If
          End If
        End If
      End If
    End If
  Next

  InitByParamArray = CBool(Count > 0)

End Function
