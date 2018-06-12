using System;
using System.Collections.Generic;
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
        private static string documentTypeGroup = System.Configuration.ConfigurationManager.AppSettings["documentTypeGroup"].ToString();
        private static string documentType = System.Configuration.ConfigurationManager.AppSettings["documentType"].ToString();
        private static string fileType = System.Configuration.ConfigurationManager.AppSettings["fileType"].ToString();
                                             
        private static Application app;
        private static DocumentTypeGroup docTypeGroup;
                                                      
        private static void GetDocumentTypeGroup()
        {
            Console.Out.WriteLine("Attempting to Get Document type group: " + documentTypeGroup);

            docTypeGroup = app.Core.DocumentTypeGroups.Find(documentTypeGroup);
                         
            if (docTypeGroup == null)
            {
                throw new Exception(documentTypeGroup + " Document type group was not found.");
            }

            Console.Out.WriteLine(documentTypeGroup + " Document type group was found.");
        }
        private static void GetDocumentTypes()
        {
            Console.Out.WriteLine("Attempting to Get Document types for group: " + documentTypeGroup);

            foreach (var dt in docTypeGroup.DocumentTypes)
            {
                Console.Out.WriteLine("Document Type: " + dt.Name);
                foreach (var krt in dt.KeywordRecordTypes)
                {
                    //Console.Out.WriteLine("Keyword Record Types: " + krt.Name);
                    Console.Out.WriteLine("Keyword Types: ");
                    foreach (var kt in krt.KeywordTypes)
                    {
                        Console.Out.WriteLine(kt.Name + " : " + kt.ID);
                    }
                    Console.Out.WriteLine("");
                }    
            }

            Console.Out.WriteLine("Keyword Record Types: ");
            foreach (var krt in app.Core.KeywordRecordTypes)
            {
                Console.Out.WriteLine(krt.Name);
            }
            Console.Out.WriteLine("");
        }

        private static void ArchiveDocument()
        {
            Console.Out.WriteLine("Attempting to archive a document...");

            DocumentType docType = docTypeGroup.DocumentTypes.Find(documentType);
            if (docType == null)
            {
                throw new Exception("Document type was not found");
            }

            FileType fType = app.Core.FileTypes.Find(fileType);
            if (fType == null)
            {
                throw new Exception("File type was not found");
            }

            StoreNewDocumentProperties storeDocumentProperties = app.Core.Storage.CreateStoreNewDocumentProperties(docType, fType);
                                                                                
            //KeywordRecordType keywordRecordType = app.Core.KeywordRecordTypes.Find("RSI - Document Keyword Group");
            KeywordRecordType keywordRecordType = docType.KeywordRecordTypes[0];
            EditableKeywordRecord editableKeywordRecord = keywordRecordType.CreateEditableKeywordRecord();
                                                   
            foreach (var kt in keywordRecordType.KeywordTypes)
            {
                if (kt.Name.Equals("DLN"))
                {
                    editableKeywordRecord.AddKeyword(kt.CreateKeyword("06122018"));
                }
                if (kt.Name.Equals("Content Type"))
                {
                    editableKeywordRecord.AddKeyword(kt.CreateKeyword("RETURN"));
                }
            }                                                                     

            storeDocumentProperties.AddKeywordRecord(editableKeywordRecord); 

            storeDocumentProperties.DocumentDate = DateTime.Now;
            storeDocumentProperties.Comment = "RSI OnBase Unity Application";
            storeDocumentProperties.Options = StoreDocumentOptions.SkipWorkflow;
                                             
            List<string> fileList = new List<string>();
            fileList.Add(@"D:\GPSCode\RSIOnBaseUnity\Files\Sample.pdf");

            Document newDocument = app.Core.Storage.StoreNewDocument(fileList, storeDocumentProperties);

            Console.Out.WriteLine(string.Format("Document import was successful. New Document ID: {0}", newDocument.ID.ToString()));
            Console.Out.WriteLine();
        }

        private static void Connect()
        {
            app = null;

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
        }

        private static void Disconnect()
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
                Connect();
                if (app != null)
                {
                    GetDocumentTypeGroup();
                    if (docTypeGroup != null)
                    {
                        GetDocumentTypes();
                        ArchiveDocument();
                    }                     
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
            finally
            {
                if (app != null)
                {
                    Disconnect();
                }              
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey(true);
        }
    }
}
