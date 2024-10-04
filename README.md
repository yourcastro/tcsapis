// Define the output parameter
var outParameter = new SqlParameter 
{ 
    ParameterName = "@totalRows", 
    SqlDbType = SqlDbType.Int, 
    Direction = ParameterDirection.Output 
};

// Build the SQL query string
string strquery = "exec [VRRPartyEntityDatastore].[dbo].[inv_party_entity_search_sp] " +
    parameters.@pageIndex + "," + 
    parameters.@rowsPerPage + ",'" + 
    parameters.@SortKey + "'," + 
    parameters.@GetTotal + ",'" + 
    parameters.@EntityID + "','" + 
    parameters.@LegalNm + "','" + 
    parameters.@DomCountryCode + "','" + 
    parameters.@UltCountryCode + "'," + 
    parameters.@RegulatoryClass + ",'" + 
    parameters.@RegulatorySubClass + "','" + 
    parameters.@Bloomberg + "','" + 
    parameters.@GICS + "','" + 
    parameters.@Lehman + "','" + 
    parameters.@MSCI + "','" + 
    parameters.@Status + "', @totalRows OUTPUT";

// Execute the SQL query with the output parameter
context.Database.ExecuteSqlRaw(strquery, outParameter);

// Retrieve the output value
int totalRows = (int)outParameter.Value;




I have an appointment on Oct 17th so I will be taking a leave.  I can be contacted through phone or TCS Teams.
