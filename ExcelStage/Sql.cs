using Microsoft.Data.SqlClient;

namespace ExcelStage;

/// <summary>Connection-string construction for the chosen server and database.</summary>
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
}
