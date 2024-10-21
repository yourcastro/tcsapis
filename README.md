public string GetScoredCardForCommerialBank(string filename, ref int ID, ref string name)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForCommerialBank(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredcardForGenericLargeCorporate(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredcardForGenericLargeCorporate(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name, GICS_Sector, GICS_Industry_Group, Lehmans_Sector
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForInsuranceCIQ(string filename, ref int ID, ref string name)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForInsuranceCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForInsuranceNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForInsuranceNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name, GICS_Sector, GICS_Industry_Group
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredcardForLeaseFinance(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForLeaseFinance(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name, GICS_Sector, GICS_Industry_Group
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredcardForMortgageBonds(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForMortgageBonds(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name, GICS_Sector, GICS_Industry_Group
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForNonBankFinancialCIQ(string filename, ref int ID, ref string name)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForNonBankFinancialNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForNonBankFinancialNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ID, name, GICS_Sector, GICS_Industry_Group
        );
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

// ... and the rest of the functions follow a similar pattern
