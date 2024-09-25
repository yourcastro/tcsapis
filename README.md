public class DataAccessLayer
{

    private readonly AppDbContext _appDbContext;
    public DataAccessLayer(AppDbContext appDbContext)
    {
        _appDbContext = appDbContext;
    }

    public int UpdatePDScorecardSPDetails(CRInvPartyEntityScorecardFactors crInvPartyEntityScorecardFactors)
    {
        var result = 0;
        var parameters = new
        {
            party_entity_scorecard_doc_id = crInvPartyEntityScorecardFactors.party_entity_scorecard_doc_id,
            party_entity_scorecard_factors_id = crInvPartyEntityScorecardFactors.party_entity_scorecard_factors_id
        };
        string sqlquery = "UPDATE [VRRPartyEntityDatastore].[dbo].[inv_party_entity_scorecard_factors_t] SET [party_entity_scorecard_doc_id] = @party_entity_scorecard_doc_id WHERE [party_entity_scorecard_factors_id]=@party_entity_scorecard_factors_id ";
        result = _appDbContext.Database.ExecuteSqlRaw(sqlquery, parameters);
        return 0;
    }
}
