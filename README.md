   var outParameter = new SqlParameter
   {
       ParameterName = "@totalRows",
       SqlDbType = SqlDbType.Int, // or the appropriate type
       Direction = ParameterDirection.Output
   };

   //var result = context.Database.ExecuteSqlRaw("EXEC YourStoredProcedure @OutParam OUTPUT", outParameter);

   //string strquery = "DECLARE @totalRows INT ";

  string strquery =  "exec [VRRPartyEntityDatastore].[dbo].[inv_party_entity_search_sp] " + parameters.@pageIndex + "," + parameters.@rowsPerPage + ",'" + parameters.@SortKey + "'," + parameters.@GetTotal + ",'" + parameters.@EntityID + "','" +
       parameters.@LegalNm + "','" + parameters.@DomCountryCode + "','" + parameters.@UltCountryCode + "'," + parameters.@RegulatoryClass + ",'" + parameters.@RegulatorySubClass + "'," +
       "'" + parameters.@Bloomberg + "','" + parameters.@GICS + "','" + parameters.@Lehman + "','" + parameters.@MSCI + "','" + parameters.@Status + "',"+
outParameter.Value+"OUTPUT";
