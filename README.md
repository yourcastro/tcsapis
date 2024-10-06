string sql = "SELECT * FROM Products WHERE 1=1"; // Using 1=1 to simplify appending conditions
List<object> parameters = new List<object>();

if (!string.IsNullOrEmpty(requestBody.Name))
{
    sql += " AND Name = @name";
    parameters.Add(new SqlParameter("@name", requestBody.Name));
}

if (requestBody.MinPrice.HasValue)
{
    sql += " AND Price >= @minPrice";
    parameters.Add(new SqlParameter("@minPrice", requestBody.MinPrice.Value));
}

if (requestBody.MaxPrice.HasValue)
{
    sql += " AND Price <= @maxPrice";
    parameters.Add(new SqlParameter("@maxPrice", requestBody.MaxPrice.Value));
}
