Option Explicit

Public Function ProductionValueEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) Or IsEmpty(b) Then
        ProductionValueEqual = IsEmpty(a) And IsEmpty(b)
    ElseIf VarType(a) = vbString Or VarType(b) = vbString Then
        ProductionValueEqual = (VarType(a) = VarType(b))
        If ProductionValueEqual Then ProductionValueEqual = StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0
    Else
        ProductionValueEqual = (a = b)
    End If
End Function

Public Function ProductionNormField(ByVal field As String) As Boolean
    Select Case field
        Case "Profile", "Quantity", "UnitWeight", "TotalWeight", "Length", "Width", "Height", "UnitSurface", "TotalSurface", "Drawing", "Revision", "Coating": ProductionNormField = True
    End Select
End Function

Private Function SourceChange(ByVal key As String, ByVal status As String, ByVal field As String, ByVal before As Variant, ByVal after As Variant, ByVal norm As Boolean) As Object
    Dim item As Object
    Set item = ProductionDictionary()
    item.Add "Key", key: item.Add "Status", status: item.Add "Field", field
    item.Add "Before", before: item.Add "After", after: item.Add "NormReview", norm
    Set SourceChange = item
End Function

Public Function CompareProductionSources(ByVal previous As Object, ByVal incoming As Object, ByVal completeOrder As String, ByVal completeScope As String, Optional ByVal previousMissing As Object = Nothing) As Collection
    Dim result As New Collection, key As Variant, field As Variant, old As Object, current As Object, missing As Boolean
    If (Len(completeOrder) = 0) <> (Len(completeScope) = 0) Then Fail "Porównanie kompletnego zakresu wymaga zamówienia i zakresu."
    For Each key In incoming.Keys
        Set current = incoming(key)
        If Len(completeScope) > 0 Then
            If current("Order") <> completeOrder Or current("Scope") <> completeScope Then Fail "Porównanie pełnego zakresu obejmuje inne zamówienie lub zakres."
        End If
        If Not previous.Exists(key) Then
            result.Add SourceChange(CStr(key), "NOWA", "Rekord", Empty, ProductionSymbol(current("Position"), current("Lfd")), True)
        Else
            Set old = previous(key): missing = False
            If Not previousMissing Is Nothing Then missing = previousMissing.Exists(key)
            If missing Then result.Add SourceChange(CStr(key), "ZMIENIONA", "Obecność w źródle", "BRAK W ŹRÓDLE", "OBECNA", False)
            For Each field In ProductionFields()
                If Not old.Exists(field) Or Not current.Exists(field) Then Fail "Niekompletny stan importu: " & CStr(field)
                If Not ProductionValueEqual(old(field), current(field)) Then result.Add SourceChange(CStr(key), "ZMIENIONA", CStr(field), old(field), current(field), ProductionNormField(CStr(field)))
            Next field
        End If
    Next key
    If Len(completeScope) > 0 Then
        For Each key In previous.Keys
            Set old = previous(key)
            If old("Order") = completeOrder And old("Scope") = completeScope Then
                If Not incoming.Exists(key) Then
                    missing = False
                    If Not previousMissing Is Nothing Then missing = previousMissing.Exists(key)
                    If Not missing Then result.Add SourceChange(CStr(key), "BRAK W ŹRÓDLE", "Obecność w źródle", "OBECNA", "BRAK W ŹRÓDLE", False)
                End If
            End If
        Next key
    End If
    Set CompareProductionSources = result
End Function

Public Function ProductionSnapshot(ByVal source As Object) As Object
    Dim result As Object, key As Variant, field As Variant, item As Object
    Set result = ProductionDictionary()
    For Each key In source.Keys
        Set item = ProductionDictionary()
        For Each field In ProductionFields(): item.Add CStr(field), source(key)(field): Next field
        result.Add CStr(key), item
    Next key
    Set ProductionSnapshot = result
End Function
