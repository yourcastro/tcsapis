Private Function CheckForExistingExcellProcesses()
        Dim AllProcesses() As Process = Process.GetProcessesByName("EXCEL")
        For Each excelProcess As Process In AllProcesses
            listProcess.Add(excelProcess.Id)
        Next
    End Function

    Private Function GetExcelProcessID()
        Dim AllProcesses() As Process = Process.GetProcessesByName("EXCEL")
        For Each excelProcess As Process In AllProcesses
            If listProcess.Contains(excelProcess.Id) = False Then
                Return excelProcess.Id
            End If
        Next
        AllProcesses = Nothing
    End Function
