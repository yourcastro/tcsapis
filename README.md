using Microsoft.EntityFrameworkCore;
using System.Data.Common;

// Assuming you have a DbContext instance (e.g., _context)
var sql = "EXEC YourStoredProcedureName @param1, @param2"; // Adjust for your stored procedure and parameters

using (var command = _context.Database.GetDbConnection().CreateCommand())
{
    command.CommandText = sql;
    command.CommandType = System.Data.CommandType.Text;

    // If you need to pass parameters to the stored procedure
    var param1 = command.CreateParameter();
    param1.ParameterName = "@param1";
    param1.Value = param1Value; // Set your parameter value
    command.Parameters.Add(param1);

    var param2 = command.CreateParameter();
    param2.ParameterName = "@param2";
    param2.Value = param2Value; // Set your parameter value
    command.Parameters.Add(param2);

    // Open connection if it’s not open
    if (command.Connection.State != System.Data.ConnectionState.Open)
    {
        await command.Connection.OpenAsync();
    }

    using (var reader = await command.ExecuteReaderAsync())
    {
        // Loop through the data rows
        while (await reader.ReadAsync())
        {
            // Access the columns by index or name
            var column1 = reader.GetString(0); // Assuming the first column is a string
            var column2 = reader.GetInt32(1);  // Assuming the second column is an int

            Console.WriteLine($"Column1: {column1}, Column2: {column2}");
        }
    }
}
