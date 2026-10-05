using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    public enum SqlAuthenticationMode
    {
        WindowsCurrentUser,
        SqlLogin
    }

    // Connection options are an in-memory input. Passwords must not be written
    // to application settings, reports or diagnostic labels.
    public sealed class SqlConnectionOptions
    {
        public SqlConnectionOptions()
        {
            Server = String.Empty;
            Database = String.Empty;
            Authentication = SqlAuthenticationMode.WindowsCurrentUser;
            Username = String.Empty;
            Password = String.Empty;
            Encrypt = true;
            TrustServerCertificate = true;
        }

        public string Server { get; set; }
        public string Database { get; set; }
        public SqlAuthenticationMode Authentication { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public bool Encrypt { get; set; }
        public bool TrustServerCertificate { get; set; }

        public string SourceName { get { return Server + " / " + Database; } }

        public SqlConnectionOptions Clone()
        {
            return new SqlConnectionOptions {
                Server = Server,
                Database = Database,
                Authentication = Authentication,
                Username = Username,
                Password = Password,
                Encrypt = Encrypt,
                TrustServerCertificate = TrustServerCertificate
            };
        }

        public SqlConnectionStringBuilder CreateConnectionStringBuilder()
        {
            return CreateConnectionStringBuilder(Database);
        }

        internal SqlConnectionStringBuilder CreateConnectionStringBuilder(string database)
        {
            if (String.IsNullOrWhiteSpace(Server))
                throw new ArgumentException("SQL server must not be empty.", "Server");
            if (String.IsNullOrWhiteSpace(database))
                throw new ArgumentException("Database must not be empty.", "Database");
            if (Authentication != SqlAuthenticationMode.WindowsCurrentUser && Authentication != SqlAuthenticationMode.SqlLogin)
                throw new ArgumentException("Choose Windows authentication or a SQL login.", "Authentication");
            if (Authentication == SqlAuthenticationMode.SqlLogin && String.IsNullOrWhiteSpace(Username))
                throw new ArgumentException("A username is required for SQL authentication.", "Username");

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = Server;
            builder.InitialCatalog = database;
            builder.IntegratedSecurity = Authentication == SqlAuthenticationMode.WindowsCurrentUser;
            if (!builder.IntegratedSecurity)
            {
                builder.UserID = Username;
                builder.Password = Password ?? String.Empty;
            }
            builder.PersistSecurityInfo = false;
            builder.ConnectTimeout = 15;
            builder.ApplicationName = "SharePoint Database Explorer";
            builder.Encrypt = Encrypt;
            builder.TrustServerCertificate = TrustServerCertificate;
            return builder;
        }

        public override string ToString()
        {
            return SourceName;
        }
    }

    public static class SqlDatabaseDiscovery
    {
        internal static SqlConnectionStringBuilder CreateMasterConnectionStringBuilder(SqlConnectionOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            // Discovery does not depend on the chosen database already being
            // available and does not alter the caller's selection.
            return options.Clone().CreateConnectionStringBuilder("master");
        }
        public static List<string> GetAccessibleDatabases(SqlConnectionOptions options)
        {
            SqlConnectionStringBuilder builder = CreateMasterConnectionStringBuilder(options);
            List<string> databases = new List<string>();
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                connection.Open();
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandTimeout = 30;
                    command.CommandText = @"
SELECT name
FROM sys.databases
WHERE state = 0 AND HAS_DBACCESS(name) = 1
ORDER BY name;";
                    using (SqlDataReader reader = command.ExecuteReader())
                        while (reader.Read()) databases.Add(reader.GetString(0));
                }
            }
            return databases;
        }
    }
}
