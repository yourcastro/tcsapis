using System;
using System.Diagnostics;

try
{
    Process aProcess = null;
    string strProcessName = Process.GetProcessById(myExcelProcessId).ProcessName;
    aProcess = Process.GetProcessById(myExcelProcessId);
    
    if (aProcess != null && strProcessName.ToUpper() == "EXCEL")
    {
        aProcess.Kill();
    }
}
catch (Exception ex)
{
    // Handle the exception if needed
}
