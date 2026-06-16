using Microsoft.Data.SqlClient;

namespace ExcelStage;

/// <summary>Connection-string construction and lightweight server browsing.</summary>
public static class Sql
{
    public static string BuildConnectionString(string server, string database) =>
        new SqlConnectionStringBuilder
        {
            DataSource = server,
            InitialCatalog = database,
            IntegratedSecurity = true,
            PersistSecurityInfo = false,
            Pooling = false,
            MultipleActiveResultSets = false,
            Encrypt = true,
            TrustServerCertificate = true,
            ConnectTimeout = 15
        }.ConnectionString;

    /// <summary>
    /// Returns the user (non-system) database names on the server, ordered by name.
    /// Connects to the <c>master</c> database to enumerate them.
    /// </summary>
    public static List<string> ListDatabases(string server)
    {
        var connectionString = BuildConnectionString(server, "master");

        using var connection = new SqlConnection(connectionString);
        connection.Open();

        const string sql = @"
            SELECT name
            FROM sys.databases
            WHERE database_id > 4              -- skip master/tempdb/model/msdb
              AND state = 0                    -- ONLINE only
              AND HAS_DBACCESS(name) = 1       -- only ones this login can use
            ORDER BY name;";

        using var command = new SqlCommand(sql, connection);
        using var reader = command.ExecuteReader();

        var databases = new List<string>();
        while (reader.Read())
        {
            databases.Add(reader.GetString(0));
        }

        return databases;
    }
}
