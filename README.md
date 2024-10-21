private void CheckForExistingExcelProcesses()
{
    Process[] allProcesses = Process.GetProcessesByName("EXCEL");
    foreach (Process excelProcess in allProcesses)
    {
        listProcess.Add(excelProcess.Id);
    }
}

private int? GetExcelProcessID()
{
    Process[] allProcesses = Process.GetProcessesByName("EXCEL");
    foreach (Process excelProcess in allProcesses)
    {
        if (!listProcess.Contains(excelProcess.Id))
        {
            return excelProcess.Id;
        }
    }
    return null;
}
