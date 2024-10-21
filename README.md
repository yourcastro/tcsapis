private bool RangeNameExists(Excel.Workbook activeWorkbook, string nname)
{
    bool rangeNameExists = false;
    foreach (Excel.Name n in activeWorkbook.Names)
    {
        if (n.Name == nname)
        {
            rangeNameExists = true;
            break; // Exit the function
        }
    }
    return rangeNameExists;
}
