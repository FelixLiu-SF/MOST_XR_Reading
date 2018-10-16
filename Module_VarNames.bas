Attribute VB_Name = "Module_VarNames"
Option Compare Database
Option Explicit

'---CONCAT_VISITVARSIDE---'
Public Function Concat_VisitVarSide(VisitArray() As String, SideArray() As String, VarArray() As String) As String()
'Concatenate variable roots with visit prefix-1 and side prefix-2

    Dim ControlArray() As String
    Dim VisitUB As Integer
    Dim VarUB As Integer
    Dim SideUB As Integer
    Dim ArrayUB As Integer
    Dim VisitIX As Integer
    Dim VarIX As Integer
    Dim SideIX As Integer
    Dim ArrayIX As Integer
    Dim Index As Integer

    'Get upper bound for each string array
    VisitUB = UBound(VisitArray,1)
    VarUB = UBound(VarArray,1)
    SideUB = UBound(SideArray,1)

    'Calculate size of output string array and redim
    ArrayUB = VisitUB*VarUB*SideUB
    ReDim ControlArray(ArrayUB) As String

    'Loop through all combinations and concatenate
    VisitIX = 0
    VarIX = 0
    SideIX = 0
    Index = 0

    Index = 0
    For VisitIX = 0 To (VisitUB - 1)
      For SideIX = 0 to (SideUB - 1)
        For VarIX = 0 To (VarUB - 1)
            ControlArray(Index) = CStr(VisitArray(VisitIX)) & CStr(SideArray(SideIX)) & CStr(VarArray(VarIX))
            Index = Index + 1
        Next
      Next
    Next

    Concat_VisitVarSide = ControlArray

End Function

'---CONCAT_VISITVAR---'
Public Function Concat_VisitVar(VisitArray() As String, VarArray() As String) As String()
'Concatenate variable roots with visit prefix

    Dim ControlArray() As String
    Dim VisitUB As Integer
    Dim VarUB As Integer

    Dim ArrayUB As Integer
    Dim VisitIX As Integer
    Dim VarIX As Integer

    Dim ArrayIX As Integer
    Dim Index As Integer

    'Get upper bound for each string array
    VisitUB = UBound(VisitArray,1)
    VarUB = UBound(VarArray,1)

    'Calculate size of output string array and redim
    ArrayUB = VisitUB*VarUB
    ReDim ControlArray(ArrayUB) As String

    'Loop through all combinations and concatenate
    VisitIX = 0
    VarIX = 0
    Index = 0

    For VisitIX = 0 To (VisitUB - 1)
      For VarIX = 0 To (VarUB - 1)
        ControlArray(Index) = CStr(VisitArray(VisitIX)) & CStr(VarArray(VarIX))
        Index = Index + 1
      Next
    Next

    Concat_VisitVar = ControlArray

End Function

'---CONCAT_PREFIX---'
Public Function Concat_Prefix(PrefixStr As String, ControlArray() As String) As String()
'Concatenate control array with static prefix

    Dim ArrayUB As Integer
    Dim Index As Integer
    Dim OutputArray() As String

    ArrayUB = UBound(ControlArray,1)
    ReDim OutputArray(ArrayUB) As String

    Index = 0
    For Index = 0 To (ArrayUB - 1)
        OutputArray(Index) = PrefixStr & CStr(ControlArray(Index))
    Next

    Concat_Prefix = OutputArray

End Function

'---CONCAT_SUFFIX---'
Public Function Concat_Suffix(ControlArray() As String, SuffixStr As String) As String()
'Concatenate control array with static suffix

    Dim ArrayUB As Integer
    Dim Index As Integer
    Dim OutputArray() As String

    ArrayUB = UBound(ControlArray,1)
    ReDim OutputArray(ArrayUB) As String

    Index = 0
    For Index = 0 To (ArrayUB - 1)
        OutputArray(Index) = CStr(ControlArray(Index)) & SuffixStr
    Next

    Concat_Suffix = OutputArray

End Function
