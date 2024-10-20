 Public Function IsCVErr(ByVal obj As Object) As Boolean
        If IsNumeric(obj) Then
            Select Case CType(obj, Integer)
                Case CVErrEnum.ErrDiv0, CVErrEnum.ErrNA, CVErrEnum.ErrName, CVErrEnum.ErrNull, CVErrEnum.ErrNum, CVErrEnum.ErrRef, CVErrEnum.ErrValue
                    Return True
                Case Else
                    Return False
            End Select
        End If
        Return False
    End Function
