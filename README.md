public static bool CheckAccess(string userID, string filename, string appName, string strConnDB, 
                               string user, string password, string strConnAuditDB, 
                               string sessionKey, ref string errorMsg)
{
    errorMsg = string.Empty;
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            return true;
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName);
        string currentUser;
        string currentStatus;

        objArguments.TableNames = new string[1];
        objArguments.FilterConditions = new string[1];
        objArguments.dsData = ds;
        objArguments.ConnectionDatabase = strConnDB;
        objArguments.ConnectionUser = user;
        objArguments.ConnectionPwd = password;
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_and_factor_v";
        objArguments.FilterConditions[0] = $"party_entity_scorecard_file_nm = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];

        if (tb.Rows.Count == 0)
        {
            errorMsg = $"File {filename} has no record saved in the database.";
            return false;
        }
        else
        {
            DataRow row = tb.Rows[0];
            currentUser = row["last_update_process_id"].ToString();
            currentStatus = row["scorecard_status_cd"].ToString();

            if (currentUser.ToUpper() == userID.ToUpper() && currentStatus.ToUpper() == "O")
            {
                return true;
            }
        }
        return false;
    }
    catch (Exception ex)
    {
        errorMsg = ex.Message;
        return false;
    }
}

public static string GetScoredCardForCommerialBank(string appName, string strConnDB, 
                                                   string user, string password, 
                                                   string strConnAuditDB, string sessionKey, 
                                                   string filename, ref int ID, ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName);

        objArguments.TableNames = new string[1];
        objArguments.FilterConditions = new string[1];
        objArguments.dsData = ds;
        objArguments.ConnectionDatabase = strConnDB;
        objArguments.ConnectionUser = user;
        objArguments.ConnectionPwd = password;
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_commerial_bank";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];

        if (tb.Rows.Count == 0)
        {
            return $"Can not find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = row["Entity_Name"].ToString();
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public static string GetScoredCardForPublicFinancNonUS(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, 
    string filename, ref int ID, ref string name, ref string GICS_Sector, ref string GICS_IG)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_IG = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_public_finance_nonus";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            GICS_Sector = Convert.ToString(row["GICS_Sector"]);
            GICS_IG = Convert.ToString(row["GICS_Industry_Group"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public static string GetScoredCardForGenericCorporateCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, 
    string filename, ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_generic_corp_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForGenericCorporateNonCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, 
    string filename, ref int ID, ref string name, ref string GICS_Sector, 
    ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_generic_corp_non_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
            GICS_Sector = Convert.ToString(row["GICS_Sector"]);
            GICS_Industry_Group = Convert.ToString(row["GICS_Industry_Group"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForUtilities(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, 
    string filename, ref int ID, ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_Utilities";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasExplorationProductionCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, string filename, 
    ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasExplorationProductionCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasMidstreamCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, string filename, 
    ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasMidstreamCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasOilFieldServicesCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, string filename, 
    ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasOilFieldServicesCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasRefiningMarketingCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, string filename, 
    ref int ID, ref string name, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasRefiningMarketingCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasExplorationProductionNonCIQ(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, string sessionKey, string filename, 
    ref int ID, ref string name, ref string GICS_Sector, ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1],
            AuditConnectionDatabases = new string[1],
            AuditConnectionUsers = new string[1],
            AuditConnectionPwds = new string[1]
        };

        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasExplorationProductionNonCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);

        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return $"Cannot find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
            GICS_Sector = Convert.ToString(row["GICS_Sector"]);
            GICS_Industry_Group = Convert.ToString(row["GICS_Industry_Group"]);
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasMidstreamNonCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group, 
    ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasMidstreamNonCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            Lehmans_Sector = (string)row["Lehmans_Sector"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasOilFieldServicesNonCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group, 
    ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasOilFieldServicesNonCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            Lehmans_Sector = (string)row["Lehmans_Sector"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForOilGasRefiningMarketingNonCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group, 
    ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_OilGasRefiningMarketingNonCIQ";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            Lehmans_Sector = (string)row["Lehmans_Sector"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForCommercialMortgage(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group, 
    ref string Property_Address, 
    ref string Loan_Number, 
    ref string Borrower_Name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            Property_Address = string.Empty;
            Loan_Number = string.Empty;
            Borrower_Name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_CommercialMortgage_v";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
            Property_Address = (string)row["Property_Address"];
            Loan_Number = (string)row["Loan_Number"];
            Borrower_Name = (string)row["Borrower_Name"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForProjFinance(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_ProjectFinance";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForNonBankFinancialCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_NonBankFinancial_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForNonBankFinancialNonCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_NonBankFinancial_non_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public static string GetScoredCardForInsuranceCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_Insurance_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForInsuranceNonCIQ(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_Insurance_non_ciq";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public static string GetScoredCardForRealEstateInvestment(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_RealEstateInvestments";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForUniversitySchoolHospital(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_UniversitySchoolHospital";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static string GetScoredCardForMortgageBonds(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_MortgageBonds";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        if (tb.Rows.Count == 0)
        {
            return "Can not find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}

public static string GetScoredCardForLeaseFinance(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    ref int ID, 
    ref string name, 
    ref string GICS_Sector, 
    ref string GICS_Industry_Group)
{
    try
    {
        // Check if the filename is a template
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        // Initialize DataSet and arguments
        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_LeaseFinance";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        // Set audit connection information
        objArguments.AuditConnectionDatabases = new string[1] { strConnAuditDB };
        objArguments.AuditConnectionUsers = new string[1] { user };
        objArguments.AuditConnectionPwds = new string[1] { password };

        // Fetch data
        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];

        // Check for data availability
        if (tb.Rows.Count == 0)
        {
            return "Cannot find PD scorecard information for " + filename;
        }
        else
        {
            // Extract information from the DataTable
            DataRow row = tb.Rows[0];
            ID = (int)row["Entity_ID"];
            name = (string)row["Entity_Name"];
            GICS_Sector = (string)row["GICS_Sector"];
            GICS_Industry_Group = (string)row["GICS_Industry_Group"];
        }

        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static bool IsATemplateFile(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename)
{
    try
    {
        // Initialize DataSet and arguments
        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "inv_scorecard_template_t";
        objArguments.FilterConditions[0] = $"scorecard_template_file_nm = '{filename}'";

        // Set audit connection information
        objArguments.AuditConnectionDatabases = new string[1] { strConnAuditDB };
        objArguments.AuditConnectionUsers = new string[1] { user };
        objArguments.AuditConnectionPwds = new string[1] { password };

        // Fetch data
        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];

        // Check if any rows were returned
        return tb.Rows.Count > 0;
    }
    catch (Exception ex)
    {
        // Return false if there was an exception, you may want to log this instead
        return false; 
    }
}



public static string GetScoredcardForGenericLargeCorporate(
    string appName, 
    string strConnDB, 
    string user, 
    string password, 
    string strConnAuditDB, 
    string sessionKey, 
    string filename, 
    out int ID, 
    out string name, 
    out string GICS_Sector, 
    out string GICS_Industry_Group, 
    out string Lehmans_Sector)
{
    try
    {
        // Initialize output variables
        ID = 0;
        name = string.Empty;
        GICS_Sector = string.Empty;
        GICS_Industry_Group = string.Empty;
        Lehmans_Sector = string.Empty;

        // Check if the file is a template
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            return "OK";
        }

        // Initialize DataSet and arguments
        DataSet ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName)
        {
            dsData = ds,
            ConnectionDatabase = strConnDB,
            ConnectionUser = user,
            ConnectionPwd = password,
            TableNames = new string[1],
            FilterConditions = new string[1]
        };
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_GenericLargeCorporate";
        objArguments.FilterConditions[0] = $"Filename = '{filename}'";

        // Set audit connection information
        objArguments.AuditConnectionDatabases = new string[1] { strConnAuditDB };
        objArguments.AuditConnectionUsers = new string[1] { user };
        objArguments.AuditConnectionPwds = new string[1] { password };

        // Fetch data
        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];

        // Check if any rows were returned
        if (tb.Rows.Count == 0)
        {
            return $"Can not find PD scorecard information for {filename}";
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
            GICS_Sector = Convert.ToString(row["GICS_Sector"]);
            GICS_Industry_Group = Convert.ToString(row["GICS_Industry_Group"]);
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}


public static void GetData(ref sGenericTableRequestArguments Arguments, string SessionKey, bool GetSecurity = false)
{
    int iCount;
    string strErrors = string.Empty;
    var objServices = new sGenericTable();
    ResultBody objResultBody;

    // Initialize GetSecurityInfo if not retrieving security
    if (!GetSecurity && Arguments.TableNames.Length > 0)
    {
        Arguments.GetSecurityInfo = new bool[Arguments.TableNames.Length];
    }

    // Execute the service function to get tables
    objResultBody = ServiceFunctions.Execute(objServices, "GetTables", Arguments, SessionKey);

    // Check for errors in the result
    if (objResultBody.ResultStatus.ReturnCode == "Error")
    {
        for (iCount = 0; iCount < objResultBody.ResultStatus.MessageCount; iCount++)
        {
            strErrors += objResultBody.ResultStatus.ResultMessage[iCount].MessageDescription + Environment.NewLine;
        }
        throw new ApplicationException(strErrors);
    }
}


public static void UpdateData(ref sGenericTableRequestArguments Arguments, string SessionKey, bool GetSecurity = false)
{
    int iCount;
    string strErrors = string.Empty;
    var objServices = new sGenericTable();
    ResultBody objResultBody;

    // Initialize GetSecurityInfo if not retrieving security
    if (!GetSecurity && Arguments.TableNames.Length > 0)
    {
        Arguments.GetSecurityInfo = new bool[Arguments.TableNames.Length];
    }

    // Execute the service function to update tables
    objResultBody = ServiceFunctions.Execute(objServices, "UpdateTables", Arguments, SessionKey, ServiceFunctions.TransactionLevel.DTSTransaction);

    // Check for errors in the result
    if (objResultBody.ResultStatus.ReturnCode == "Error")
    {
        for (iCount = 0; iCount < objResultBody.ResultStatus.MessageCount; iCount++)
        {
            strErrors += objResultBody.ResultStatus.ResultMessage[iCount].MessageDescription + Environment.NewLine;
        }
        throw new ApplicationException(strErrors);
    }
}


public static string GetScoredcardForSmallMediumEnterprises(string appName, string strConnDB, 
    string user, string password, string strConnAuditDB, 
    string sessionKey, string filename, 
    ref int ID, ref string name, ref string GICS_Sector, 
    ref string GICS_Industry_Group, ref string Lehmans_Sector)
{
    try
    {
        if (IsATemplateFile(appName, strConnDB, user, password, strConnAuditDB, sessionKey, filename))
        {
            ID = 0;
            name = string.Empty;
            Lehmans_Sector = string.Empty;
            GICS_Sector = string.Empty;
            GICS_Industry_Group = string.Empty;
            return "OK";
        }

        var ds = new DataSet();
        var objArguments = new sGenericTableRequestArguments(appName);
        
        objArguments.TableNames = new string[1];
        objArguments.FilterConditions = new string[1];
        objArguments.dsData = ds;
        objArguments.ConnectionDatabase = strConnDB;
        objArguments.ConnectionUser = user;
        objArguments.ConnectionPwd = password;
        objArguments.TableNames[0] = "cr_interface_PD_scorecard_predata_SmallMediumEnterprises";
        objArguments.FilterConditions[0] = "Filename = '" + filename + "'";

        objArguments.AuditConnectionDatabases = new string[1];
        objArguments.AuditConnectionUsers = new string[1];
        objArguments.AuditConnectionPwds = new string[1];
        objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
        objArguments.AuditConnectionUsers[0] = user;
        objArguments.AuditConnectionPwds[0] = password;

        GetData(objArguments, sessionKey);
        DataTable tb = ds.Tables[0];
        
        if (tb.Rows.Count == 0)
        {
            return "Cannot find PD scorecard information for " + filename;
        }
        else
        {
            DataRow row = tb.Rows[0];
            ID = Convert.ToInt32(row["Entity_ID"]);
            name = Convert.ToString(row["Entity_Name"]);
            Lehmans_Sector = Convert.ToString(row["Lehmans_Sector"]);
            GICS_Sector = Convert.ToString(row["GICS_Sector"]);
            GICS_Industry_Group = Convert.ToString(row["GICS_Industry_Group"]);
        }
        return "OK";
    }
    catch (Exception ex)
    {
        return ex.Message;
    }
}




