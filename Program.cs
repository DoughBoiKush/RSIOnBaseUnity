using System;
using System.Net;
using Hyland.Unity;

namespace RSIOnBaseUnity
{
    public class Program
    {
        private static string appServerURL = System.Configuration.ConfigurationManager.AppSettings["appServerURL"].ToString();
        private static string username = System.Configuration.ConfigurationManager.AppSettings["username"].ToString();
        private static string password = System.Configuration.ConfigurationManager.AppSettings["password"].ToString();
        private static string dataSource = System.Configuration.ConfigurationManager.AppSettings["dataSource"].ToString();     

        private static Application Connect()
        {
            Application app = null;

            AuthenticationProperties authProps = Application.CreateOnBaseAuthenticationProperties(appServerURL, username, password, dataSource);
            authProps.LicenseType = LicenseType.Default;

            Console.Out.WriteLine("Attempting to make a connection...");

            try
            {
                app = Application.Connect(authProps);
            }
            catch (MaxLicensesException)
            {
                Console.Out.WriteLine("Error: All available licenses have been consumed.");
                Console.Out.WriteLine();
            }
            catch (SystemLockedOutException)
            {
                Console.Out.WriteLine("Error: The system is currently in lockout mode.");
                Console.Out.WriteLine();
            }
            catch (InvalidLoginException)
            {
                Console.Out.WriteLine("Error: Invalid Login Credentials.");
                Console.Out.WriteLine();
            }
            catch (UserAccountLockedException)
            {
                Console.Out.WriteLine("Error: This account has been temporarily locked.");
                Console.Out.WriteLine();
            }
            catch (AuthenticationFailedException)
            {
                Console.Out.WriteLine("Error: NT Authentication Failed.");
                Console.Out.WriteLine();
            }
            catch (MaxConcurrentLicensesException)
            {
                Console.Out.WriteLine("Error: All concurrent licenses for this user group have been consumed.");
                Console.Out.WriteLine();
            }
            catch (InvalidLicensingException)
            {
                Console.Out.WriteLine("Error: Invalid Licensing.");
                Console.Out.WriteLine();
            }

            if (app != null)
            {
                Console.Out.WriteLine("Connection Successful. Connection ID: " + app.SessionID);
                Console.Out.WriteLine();
            }

            return app;
        }

        private static void Disconnect(Application app)
        {
            Console.Out.WriteLine("Attempting to close connection...");

            try
            {
                app.Disconnect();
            }
            catch (SessionNotFoundException)
            {
                Console.Out.WriteLine("Error: Active session could not be found.");
                Console.Out.WriteLine();
            }
            finally
            {
                app.Dispose();
            }

            Console.Out.WriteLine("Connection closed.");
            Console.Out.WriteLine();
        }

        public static void Main(string[] args)
        {
            try
            {                                                       
                Application app = Connect();

                if (app != null)
                {
                    Disconnect(app);
                }
            }
            catch (WebException ex)
            {
                Console.Out.WriteLine("General network error: " + ex.Message);
                Console.Out.WriteLine();
            }
            catch (UnityAPIException ex)
            {
                Console.Out.WriteLine("General Unity API error: " + ex.Message);
                Console.Out.WriteLine();
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine("General error: " + ex.Message);
                Console.Out.WriteLine();
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey(true);
        }
    }
}
