public string GetScoredCardForOilGasExplorationProductionCIQ(string filename, ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasExplorationProductionNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasExplorationProductionNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref GICS_Sector, ref GICS_Industry_Group, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasMidstreamCIQ(string filename, ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasMidstreamNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasMidstreamNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref GICS_Sector, ref GICS_Industry_Group, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasOilFieldServicesCIQ(string filename, ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasOilFieldServicesNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasOilFieldServicesNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref GICS_Sector, ref GICS_Industry_Group, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasRefiningMarketingCIQ(string filename, ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public string GetScoredCardForOilGasRefiningMarketingNonCIQ(string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        return ClsPDScoreCardFunctions.GetScoredCardForOilGasRefiningMarketingNonCIQ(
            AppSettings.Get(CONFIG_APPNAME),
            AppSettings.Get(CONFIG_DB_DATASTORE_CONNECTION),
            AppSettings.Get(CONFIG_DB_USER),
            AppSettings.Get(CONFIG_DB_PWD),
            AppSettings.Get(CONFIG_DB_OPERATIONAL_CONNECTION),
            CreateSessionId(),
            filename, ref ID, ref name, ref GICS_Sector, ref GICS_Industry_Group, ref Lehmans_Sector);
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}
