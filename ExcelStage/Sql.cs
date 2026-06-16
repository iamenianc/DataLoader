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

    // Server-level connection that does NOT pin a database, so it works even for
    // logins that cannot open 'master' (sys.databases is readable from any DB).
    private static string BuildServerConnectionString(string server) =>
        new SqlConnectionStringBuilder
        {
            DataSource = server,
            IntegratedSecurity = true,
            PersistSecurityInfo = false,
            Pooling = false,
            MultipleActiveResultSets = false,
            Encrypt = true,
            TrustServerCertificate = true,
            ConnectTimeout = 10
        }.ConnectionString;

    /// <summary>
    /// Returns the database names the current login can see on the server, ordered
    /// by name. Tries progressively simpler queries so it works across permission
    /// levels and SQL Server versions.
    /// </summary>
    public static List<string> ListDatabases(string server)
    {
        using var connection = new SqlConnection(BuildServerConnectionString(server));
        connection.Open();

        var queries = new[]
        {
            // Preferred: only online, user databases this login can actually use.
            @"SELECT name FROM sys.databases
              WHERE database_id > 4 AND state = 0 AND HAS_DBACCESS(name) = 1
              ORDER BY name;",
            // Fallback: all non-system databases.
            @"SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name;",
            // Last resort: everything sys.databases will return.
            @"SELECT name FROM sys.databases ORDER BY name;"
        };

        foreach (var sql in queries)
        {
            try
            {
                using var command = new SqlCommand(sql, connection);
                using var reader = command.ExecuteReader();

                var databases = new List<string>();
                while (reader.Read())
                {
                    databases.Add(reader.GetString(0));
                }

                return databases;
            }
            catch (SqlException)
            {
                // Permission/version issue with this query - try a simpler one.
            }
        }

        return new List<string>();
    }
}
