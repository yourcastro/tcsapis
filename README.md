using System;
using System.Data;
using System.Xml;
using IT.INV.Service.Base;

public class ClsPDScoreCardProcessor
{
    private int mFileID;
    private string mFilename;
    private string mUserID;
    private string mAppName;
    private string mStrConnDB;
    private string mStrDBUser;
    private string mStrDBPwd;
    private string mStrConnAuditDB;
    private string mStrSessionKey;
    private string mXML;

    public ClsPDScoreCardProcessor(int fileID, string filename, string userID, string appName, string strConnDB, string user, string password, string strConnAuditDB, string sessionKey, string xml)
    {
        mFileID = fileID;
        mFilename = filename;
        mUserID = userID;
        mAppName = appName;
        mStrConnDB = strConnDB;
        mStrDBUser = user;
        mStrDBPwd = password;
        mStrConnAuditDB = strConnAuditDB;
        mStrSessionKey = sessionKey;
        mXML = xml;
    }

    public void WriteDB()
    {
        string msg = "OK";
        try
        {
            if (ClsPDScoreCardFunctions.IsATemplateFile(mAppName, mStrConnDB, mStrDBUser, mStrDBPwd, mStrConnAuditDB, mStrSessionKey, mFilename))
            {
                return;
            }

            DataSet m_dsData = new DataSet();
            sGenericTableRequestArguments objArguments = new sGenericTableRequestArguments(mAppName)
            {
                dsData = m_dsData,
                ConnectionDatabase = mStrConnDB,
                ConnectionUser = mStrDBUser,
                ConnectionPwd = mStrDBPwd
            };

            objArguments.AuditConnectionDatabases = new string[] { mStrConnAuditDB };
            objArguments.AuditConnectionUsers = new string[] { mStrDBUser };
            objArguments.AuditConnectionPwds = new string[] { mStrDBPwd };
            objArguments.TableNames = new string[] { "inv_party_entity_scorecard_factors_t" };
            objArguments.FilterConditions = new string[] { "party_entity_scorecard_file_nm = '" + mFilename + "'" };

            string sessionKey = mStrSessionKey;
            ClsPDScoreCardFunctions.GetData(objArguments, sessionKey);
            DataTable tb = m_dsData.Tables[0];

            if (tb.Rows.Count > 0)
            {
                DataRow row = tb.Rows[0];
                SetRowValues(row);
                ClsPDScoreCardFunctions.UpdateData(objArguments, sessionKey);
                WriteScoreCardTable(Convert.ToInt32(row["party_entity_scorecard_factors_id"]));
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public void WriteScoreCardTable(int ScorecardFactorsID)
    {
        string msg = "OK";
        try
        {
            DataSet m_dsData = new DataSet();
            sGenericTableRequestArguments objArguments = new sGenericTableRequestArguments(mAppName)
            {
                dsData = m_dsData,
                ConnectionDatabase = mStrConnDB,
                ConnectionUser = mStrDBUser,
                ConnectionPwd = mStrDBPwd
            };

            objArguments.AuditConnectionDatabases = new string[] { mStrConnAuditDB };
            objArguments.AuditConnectionUsers = new string[] { mStrDBUser };
            objArguments.AuditConnectionPwds = new string[] { mStrDBPwd };
            objArguments.TableNames = new string[] { "inv_party_entity_scorecard_t" };
            objArguments.FilterConditions = new string[] { "party_entity_scorecard_factors_id = " + ScorecardFactorsID.ToString() };

            string sessionKey = mStrSessionKey;
            ClsPDScoreCardFunctions.GetData(objArguments, sessionKey);
            DataTable tb = m_dsData.Tables[0];

            if (tb.Rows.Count > 0)
            {
                DataRow row = tb.Rows[0];
                row["last_update_process_id"] = mUserID;
                row["last_update_dt_tm"] = DateTime.Now;
            }
            ClsPDScoreCardFunctions.UpdateData(objArguments, sessionKey);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    private void SetRowValues(DataRow row)
    {
        try
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(mXML);
            XmlNodeList nodes = xmlDoc.GetElementsByTagName("Data");
            if (nodes.Count > 0)
            {
                XmlNode nodeData = nodes.Item(0);
                if (nodeData.HasChildNodes)
                {
                    foreach (XmlNode node in nodeData.ChildNodes)
                    {
                        row[node.Name] = node.Attributes["value"].InnerText;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public string TestDBAccess()
    {
        string msg = "OK";
        try
        {
            DataSet m_dsData = new DataSet();
            sGenericTableRequestArguments objArguments = new sGenericTableRequestArguments(mAppName)
            {
                dsData = m_dsData,
                ConnectionDatabase = mStrConnDB,
                ConnectionUser = mStrDBUser,
                ConnectionPwd = mStrDBPwd
            };

            objArguments.AuditConnectionDatabases = new string[] { mStrConnAuditDB };
            objArguments.AuditConnectionUsers = new string[] { mStrDBUser };
            objArguments.AuditConnectionPwds = new string[] { mStrDBPwd };
            objArguments.TableNames = new string[] { "inv_pdscorecard_file_t" };
            objArguments.FilterConditions = new string[] { "pdscorecard_file_id = 4" };

            string sessionKey = mStrSessionKey;
            ClsPDScoreCardFunctions.GetData(objArguments, sessionKey);
            DataTable tb = m_dsData.Tables[0];

            if (tb.Rows.Count == 0)
            {
                DataRow row = tb.NewRow();
                row["pdscorecard_file_id"] = 4;
                row["excel_file_nm"] = "Test.xls";
                row["extract_data1"] = "Four";
                tb.Rows.Add(row);
            }
            else
            {
                DataRow row = tb.Rows[0];
                row["extract_data1"] = "One";
            }

            ClsPDScoreCardFunctions.UpdateData(objArguments, sessionKey);
            return msg;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }
}
