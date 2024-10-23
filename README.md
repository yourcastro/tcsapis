var objArguments = new sGenericTableRequestArguments(appName);
{
    ref var withBlock = ref objArguments;
    withBlock.TableNames = new string[1];
    withBlock.FilterConditions = new string[1];
    withBlock.dsData = ds;
    withBlock.ConnectionDatabase = strConnAuditDB;
    withBlock.ConnectionUser = user;
    withBlock.ConnectionPwd = password;
    withBlock.TableNames[0] = "Credit_Risk_Batch_Status_t";
    withBlock.FilterConditions[0] = "batch_process_cd = 'CANMTGPD'";

    withBlock.AuditConnectionDatabases = new string[1];
    withBlock.AuditConnectionUsers = new string[1];
    withBlock.AuditConnectionPwds = new string[1];
    withBlock.AuditConnectionDatabases[0] = strConnAuditDB;
    withBlock.AuditConnectionUsers[0] = user;
    withBlock.AuditConnectionPwds[0] = password;
}
