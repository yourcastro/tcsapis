using System;
using System.Data;
using System.Data.SqlClient;

public class DatabaseExecutor
{
    public void ExecuteSqlQuery(sGenericTableRequestArguments args)
    {
        // Build the connection string using the arguments
        var connectionString = $"Server={args.ConnectionDatabase};User Id={args.ConnectionUser};Password={args.ConnectionPwd};";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            // Build the SQL query
            string query = $"SELECT * FROM {args.TableNames[0]} WHERE {args.FilterConditions[0]}";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                // Execute the query and fill the DataSet
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(args.dsData);
                }
            }

            connection.Close();
        }

        // Logging or handling audit connections can be done here if needed
    }
}


Calling the above class

string appName = "YourApp";
    string strConnAuditDB = "YourDatabase";
    string user = "YourUsername";
    string password = "YourPassword";
    DataSet ds = new DataSet();

    // Initialize the arguments
    var objArguments = new sGenericTableRequestArguments(appName)
    {
        TableNames = new string[1],
        FilterConditions = new string[1],
        dsData = ds,
        ConnectionDatabase = strConnAuditDB,
        ConnectionUser = user,
        ConnectionPwd = password,
        AuditConnectionDatabases = new string[1],
        AuditConnectionUsers = new string[1],
        AuditConnectionPwds = new string[1]
    };

    objArguments.TableNames[0] = "Credit_Risk_Batch_Status_t";
    objArguments.FilterConditions[0] = "batch_process_cd = 'CANMTGPD'";
    objArguments.AuditConnectionDatabases[0] = strConnAuditDB;
    objArguments.AuditConnectionUsers[0] = user;
    objArguments.AuditConnectionPwds[0] = password;

    // Create an instance of the database executor
    var executor = new DatabaseExecutor();
    executor.ExecuteSqlQuery(objArguments);
