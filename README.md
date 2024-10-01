using (var context = new YourDbContext())
{
    var outParameter = new SqlParameter
    {
        ParameterName = "@OutParam",
        SqlDbType = SqlDbType.Int, // or the appropriate type
        Direction = ParameterDirection.Output
    };

    var result = context.Database.ExecuteSqlRaw("EXEC YourStoredProcedure @OutParam OUTPUT", outParameter);

    // Retrieve the value of the OUT parameter
    int outValue = (int)outParameter.Value;

    Console.WriteLine($"The OUT parameter value is: {outValue}");
}
