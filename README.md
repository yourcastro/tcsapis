Public Function Execute(ByRef objService As BaseServiceObject, ByVal FunctionName As String, ByVal Arguments As RequestArguments) As ResultBody
    Dim mResultBody As ResultBody = objService.GetResultBody
 
    Try
        InitializeService(mResultBody)
        CallByName(objService, FunctionName, CallType.Method, New Object(1) {Arguments, mResultBody})
        CompleteService(mResultBody)
    Catch ex As Exception
        HandleServiceExceptions(mResultBody, ex)
    End Try
 
    Return mResultBody
End Function
