 public Task<int> UpdatePDScorecardSPDetails(CRInvPartyEntityScorecardFactors crInvPartyEntityScorecardFactors)
 {
     var result = 0;

     var parameters = new
     {
         @party_entity_scorecard_doc_id = crInvPartyEntityScorecardFactors.party_entity_scorecard_doc_id,
         @party_entity_scorecard_factors_id = crInvPartyEntityScorecardFactors.party_entity_scorecard_factors_id
     };

     string sqlquery = "UPDATE [VRRPartyEntityDatastore].[dbo].[inv_party_entity_scorecard_factors_t] SET [party_entity_scorecard_doc_id] ={0} WHERE [party_entity_scorecard_factors_id]={1} ";

     result = _appDbContext.Database.ExecuteSqlRawAsync(sqlquery, parameters.party_entity_scorecard_doc_id, parameters.party_entity_scorecard_factors_id);

     return 0;
 }
