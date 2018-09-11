Attribute VB_Name = "Module_VarNames"
Option Compare Database
Option Explicit

'---CONCAT_VISITVARSIDE---'
Public Function Concat_VisitVarSide(VisitArray() As String, VarArray() As String, SideArray() As String) As String()

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

    VisitUB = UBound(VisitArray,1)
    VarUB = UBound(VarArray,1)
    SideUB = UBound(SideArray,1)

    ArrayUB = VisitUB*VarUB*SideUB

    ReDim ControlArray(ArrayUB) As String

    VisitIX = 0
    VarIX = 0
    SideIX = 0
    Index = 0

    Index = 0
    For VisitIX = 0 To VisitUB
      For VarIX = 0 To VarUB
        For SideIX = 0 to SideUB
            ControlArray(Index) = CStr(VisitArray(VisitIX)) & CStr(VarArray(VarIX)) & CStr(SideArray(SideIX))
            Index = Index + 1
        Next
      Next
    Next

    Concat_VisitVarSide = ControlArray

End Function

'---CONCAT_PREFIX---'
Public Function Concat_Prefix(PrefixStr As String, ControlArray() As String) As String()

    Dim ArrayUB As Integer
    Dim Index As Integer

    ArrayUB = UBound(ControlArray,1) - 1

    Index = 0
    For Index = 0 To ArrayUB
        ControlArray(Index) = PrefixStr & CStr(ControlArray(Index))
    Next

    Concat_Prefix = ControlArray

End Function

'---CONCAT_SUFFIX---'
Public Function Concat_Suffix(ControlArray() As String, SuffixStr As String) As String()

    Dim ArrayUB As Integer
    Dim Index As Integer

    ArrayUB = UBound(ControlArray,1) - 1

    Index = 0
    For Index = 0 To ArrayUB
        ControlArray(Index) = CStr(ControlArray(Index)) & SuffixStr
    Next

    Concat_Suffix = ControlArray

End Function
