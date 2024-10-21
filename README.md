Try
                Dim aProcess As Process = Nothing
                strProcessName = Process.GetProcessById(myExcelProcessId).ProcessName
                aProcess = Process.GetProcessById(myExcelProcessId)
                If aProcess IsNot Nothing And UCase(strProcessName) = "EXCEL" Then
                    aProcess.Kill()
                End If
            Catch ex As Exception
            End Try
