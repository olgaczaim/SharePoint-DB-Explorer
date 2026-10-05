using System;
using Microsoft.Data.SqlClient;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    public static class SqlConnectionChecks
    {
        public static void Run()
        {
            TestWindowsDefaults();
            TestSqlCredentialsAndEscaping();
            TestCloningAndDiscovery();
            TestValidation();
            Console.WriteLine("PASS SQL connection options, authentication, escaping and credential-safe labels");
        }

        private static void TestWindowsDefaults()
        {
            SqlConnectionOptions options = new SqlConnectionOptions();
            Check(options.Server == String.Empty && options.Database == String.Empty, "A new connection must start without a preset server or database.");
            options.Server = "SQL";
            options.Database = "WSS_Content";
            options.Username = "ignored-login";
            options.Password = "ignored-secret";
            SqlConnectionStringBuilder builder = options.CreateConnectionStringBuilder();
            Check(builder.IntegratedSecurity && builder.UserID == String.Empty && builder.Password == String.Empty, "Windows authentication forwarded SQL credentials.");
            Check(builder.Encrypt && builder.TrustServerCertificate && !builder.PersistSecurityInfo, "Existing transport defaults or credential suppression changed.");
            Check(builder.ConnectTimeout == 15 && builder.ApplicationName == "SharePoint Database Explorer", "Connection behavior changed.");
            Check(!builder.ConnectionString.Contains(options.Password) && !builder.ConnectionString.Contains(options.Username), "Ignored credentials appeared in a Windows connection string.");
            SqlRepository original = new SqlRepository("SQL", "WSS_Content");
            SqlConnectionStringBuilder originalBuilder = new SqlConnectionStringBuilder(original.ConnectionString);
            Check(originalBuilder.IntegratedSecurity && originalBuilder.Encrypt && originalBuilder.TrustServerCertificate, "The original constructor changed authentication or transport defaults.");
            Check(original.SourceName == "SQL / WSS_Content", "The original source label changed.");
        }

        private static void TestSqlCredentialsAndEscaping()
        {
            SqlConnectionOptions options = new SqlConnectionOptions {
                Server = "SQL;Integrated Security=true;User ID=intruder",
                Database = "Content;TrustServerCertificate=false;Encrypt=false",
                Authentication = SqlAuthenticationMode.SqlLogin,
                Username = "reader;Integrated Security=true;Password=wrong",
                Password = "secret;User ID=attacker;\"quote\"='value'",
                Encrypt = true,
                TrustServerCertificate = false
            };
            SqlConnectionStringBuilder roundTrip = new SqlConnectionStringBuilder(options.CreateConnectionStringBuilder().ConnectionString);
            Check(roundTrip.DataSource == options.Server && roundTrip.InitialCatalog == options.Database, "Server or database text was interpreted as connection-string syntax.");
            Check(!roundTrip.IntegratedSecurity && roundTrip.UserID == options.Username && roundTrip.Password == options.Password, "Credential text was interpreted as connection-string syntax.");
            Check(roundTrip.Encrypt && !roundTrip.TrustServerCertificate && !roundTrip.PersistSecurityInfo, "Injected values overrode transport or credential settings.");
            SqlRepository repository = new SqlRepository(options);
            Check(repository.SourceName == options.Server + " / " + options.Database, "Source label did not identify the selected source.");
            Check(!repository.SourceName.Contains(options.Password) && !repository.SourceName.Contains(options.Username), "Credentials appeared in the source label.");
            Check(!options.ToString().Contains(options.Password) && !options.ToString().Contains(options.Username), "Credentials appeared in the options display text.");
            string connection = repository.ConnectionString;
            options.Server = "changed";
            options.Password = "changed";
            options.Authentication = SqlAuthenticationMode.WindowsCurrentUser;
            Check(repository.ConnectionString == connection && repository.SourceName != options.SourceName, "Repository connection settings changed after construction.");
        }

        private static void TestCloningAndDiscovery()
        {
            SqlConnectionOptions options = new SqlConnectionOptions {
                Server = "SQL\\instance",
                Database = "SelectedDatabase",
                Authentication = SqlAuthenticationMode.SqlLogin,
                Username = "read-only-login",
                Password = "private-test-password",
                Encrypt = false,
                TrustServerCertificate = false
            };
            SqlConnectionOptions clone = options.Clone();
            Check(!Object.ReferenceEquals(options, clone), "Clone reused the mutable options instance.");
            SqlConnectionStringBuilder copy = clone.CreateConnectionStringBuilder();
            Check(copy.DataSource == options.Server && copy.InitialCatalog == options.Database && copy.UserID == options.Username && copy.Password == options.Password && !copy.Encrypt && !copy.TrustServerCertificate, "Clone lost a connection setting.");
            clone.Authentication = SqlAuthenticationMode.WindowsCurrentUser;
            clone.Database = "OtherDatabase";
            clone.Encrypt = true;
            clone.TrustServerCertificate = true;
            SqlConnectionStringBuilder windows = clone.CreateConnectionStringBuilder();
            Check(windows.IntegratedSecurity && windows.UserID == String.Empty && windows.Password == String.Empty, "Authentication switching retained SQL credentials in the connection.");
            Check(options.Authentication == SqlAuthenticationMode.SqlLogin && options.Database == "SelectedDatabase" && !options.Encrypt && !options.TrustServerCertificate, "Changing a clone mutated the source options.");
            clone.Authentication = SqlAuthenticationMode.SqlLogin;
            Check(clone.CreateConnectionStringBuilder().Password == options.Password, "Switching back to SQL authentication lost the in-memory credential.");
            SqlConnectionStringBuilder master = SqlDatabaseDiscovery.CreateMasterConnectionStringBuilder(options);
            Check(master.InitialCatalog == "master" && master.DataSource == options.Server && master.UserID == options.Username && master.Password == options.Password, "Database discovery did not use master with the selected authentication.");
            Check(options.Database == "SelectedDatabase", "Discovery altered the selected database.");
            options.Database = String.Empty;
            Check(SqlDatabaseDiscovery.CreateMasterConnectionStringBuilder(options).InitialCatalog == "master", "Discovery incorrectly required a database selection.");
        }

        private static void TestValidation()
        {
            Throws<ArgumentNullException>(delegate { new SqlRepository((SqlConnectionOptions)null); });
            Throws<ArgumentNullException>(delegate { RecoverySession.OpenSql((SqlConnectionOptions)null); });
            Throws<ArgumentNullException>(delegate { SqlDatabaseDiscovery.GetAccessibleDatabases(null); });
            SqlConnectionOptions options = new SqlConnectionOptions { Server = " " };
            Throws<ArgumentException>(delegate { options.CreateConnectionStringBuilder(); });
            options.Server = "SQL";
            options.Database = " ";
            Throws<ArgumentException>(delegate { options.CreateConnectionStringBuilder(); });
            options.Database = "WSS_Content";
            options.Authentication = (SqlAuthenticationMode)999;
            Throws<ArgumentException>(delegate { options.CreateConnectionStringBuilder(); });
            options.Authentication = SqlAuthenticationMode.SqlLogin;
            options.Username = " ";
            options.Password = "never-report-this-password";
            Exception failure = Capture(delegate { options.CreateConnectionStringBuilder(); });
            Check(failure is ArgumentException && !failure.ToString().Contains(options.Password), "Validation reported a credential.");
        }

        private static Exception Capture(Action action)
        {
            try { action(); }
            catch (Exception failure) { return failure; }
            throw new Exception("An invalid connection setting was accepted.");
        }
        private static void Throws<T>(Action action) where T : Exception
        {
            Exception failure = Capture(action);
            if (!(failure is T)) throw new Exception("Expected " + typeof(T).Name + ", received " + failure.GetType().Name + ".");
        }
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }
    }
}
