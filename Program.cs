using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Hyland.Unity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace RSIOnBaseUnity
{
    public class Program
    {
        private static string APP_SERVER_URL = System.Configuration.ConfigurationManager.AppSettings["appServerURL"].ToString();
        private static string USERNAME = System.Configuration.ConfigurationManager.AppSettings["username"].ToString();
        private static string PASSWORD = System.Configuration.ConfigurationManager.AppSettings["password"].ToString();
        private static string DATA_SOURCE = System.Configuration.ConfigurationManager.AppSettings["dataSource"].ToString();
        private static string DOCUMENT_TYPE_GROUP = System.Configuration.ConfigurationManager.AppSettings["documentTypeGroup"].ToString();
        private static string DOCUMENT_TYPE = System.Configuration.ConfigurationManager.AppSettings["documentType"].ToString();
        private static string FILE_TYPE = System.Configuration.ConfigurationManager.AppSettings["fileType"].ToString();
        private static string CONTENT_UPLOAD = System.Configuration.ConfigurationManager.AppSettings["contentUpload"].ToString();
        private static string CONTENT_DOWNLOAD = System.Configuration.ConfigurationManager.AppSettings["contentDownload"].ToString();

        private static Application app;
        private static DocumentTypeGroup docTypeGroup;
        private static List<long> documentIdList = new List<long>();

        private static void GetDocumentTypeGroup()
        {
            Console.Out.WriteLine("Attempting to Get Document type group: " + DOCUMENT_TYPE_GROUP);

            docTypeGroup = app.Core.DocumentTypeGroups.Find(DOCUMENT_TYPE_GROUP);
                         
            if (docTypeGroup == null)
            {
                throw new Exception(DOCUMENT_TYPE_GROUP + " Document type group was not found.");
            }

            Console.Out.WriteLine(DOCUMENT_TYPE_GROUP + " Document type group was found.");
        }
        private static void GetDocumentTypes()
        {
            Console.Out.WriteLine("Attempting to Get Document types for group: " + DOCUMENT_TYPE_GROUP);
            Console.Out.WriteLine("");

            foreach (var dt in docTypeGroup.DocumentTypes)
            {
                Console.Out.WriteLine("Document Type: " + dt.Name + " : " + dt.ID);
                foreach (var krt in dt.KeywordRecordTypes)
                {                                                                
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

            string filePath = CONTENT_UPLOAD + "\\upload.json";
            if (File.Exists(filePath))
            {
                Console.Out.WriteLine("Content upload config file found: " + filePath);

                string inputJSON = File.ReadAllText(filePath);

                DocumentType docType = docTypeGroup.DocumentTypes.Find(DOCUMENT_TYPE);
                if (docType == null)
                {
                    throw new Exception("Document type was not found");
                }

                FileType fType = app.Core.FileTypes.Find(FILE_TYPE);
                if (fType == null)
                {
                    throw new Exception("File type was not found");
                }
                        
                KeywordRecordType keywordRecordType = docType.KeywordRecordTypes[0];
                        
                IList<JToken> contents = JToken.Parse(inputJSON)["contents"].Children().ToList();

                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();

                    string fileUploadPath = CONTENT_UPLOAD + "\\" + jContent.file;
                    if (File.Exists(fileUploadPath))
                    {
                        Console.Out.WriteLine("Content upload file found: " + fileUploadPath);
                        List<string> fileList = new List<string>();
                        fileList.Add(fileUploadPath);

                        StoreNewDocumentProperties storeDocumentProperties = app.Core.Storage.CreateStoreNewDocumentProperties(docType, fType);
                        foreach (var kt in keywordRecordType.KeywordTypes)
                        {
                            if (jContent.keywords.ContainsKey(kt.Name))
                            {
                                storeDocumentProperties.AddKeyword(kt.CreateKeyword(jContent.keywords[kt.Name]));
                            }
                        }

                        storeDocumentProperties.DocumentDate = DateTime.Now;
                        storeDocumentProperties.Comment = "RSI OnBase Unity Application";
                        storeDocumentProperties.Options = StoreDocumentOptions.SkipWorkflow;  

                        Document newDocument = app.Core.Storage.StoreNewDocument(fileList, storeDocumentProperties);

                        documentIdList.Add(newDocument.ID);
                        Console.Out.WriteLine(string.Format("Document import was successful. New Document ID: {0}", newDocument.ID.ToString()));
                    }
                    else
                    {
                        Console.Out.WriteLine("Content upload file not found: " + fileUploadPath);
                    }
                }
            }
            else
            {
                Console.Out.WriteLine("Content upload config file not found: " + filePath);
            }
                                                               
            Console.Out.WriteLine();
        }

        private static void DocumentLookup()
        {
            Console.Out.WriteLine("Attempting to find document...");

            foreach (var docID in documentIdList)
            {
                Document document = app.Core.GetDocumentByID(docID);
                if (document == null)
                {
                    throw new Exception("Document was not found");
                }

                Console.Out.WriteLine(string.Format("Document was retrieved successfully. Document Id: {0}", document.ID.ToString()));
            }
            Console.Out.WriteLine();
        }
        private static void ExecuteQuery()
        {
            Console.Out.WriteLine("Attempting to execute a document query...");

            DocumentType documentType = app.Core.DocumentTypes.Find(DOCUMENT_TYPE);
            if (documentType == null)
            {
                throw new Exception("Document type was not found");
            }

            DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
            documentQuery.AddDocumentType(documentType);

            DocumentList docList = documentQuery.Execute(long.MaxValue);

            Console.Out.WriteLine("Displaying first 10 documents returned.");
            Console.Out.WriteLine();

            int limit = (docList.Count < 10) ? docList.Count : 10;

            for (int x = 0; x < limit; x++)
            {
                Console.Out.WriteLine(string.Format("{0}. {1}", (x + 1).ToString(), docList[x].DateStored.ToShortDateString()));
            }
            Console.Out.WriteLine();
        }
        private static void DocumentQuery()
        {
            Console.Out.WriteLine("Attempting to execute a document query...");

            DocumentType documentType = app.Core.DocumentTypes.Find(DOCUMENT_TYPE);
            if (documentType == null)
            {
                throw new Exception("Document type was not found");
            }

            DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
            documentQuery.AddDisplayColumn(DisplayColumnType.DocumentDate);
            documentQuery.AddDocumentType(documentType);
            
            using (QueryResult queryResults = documentQuery.ExecuteQueryResults(10L))
            {
                Console.Out.WriteLine("Displaying first 10 documents returned.");

                foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                {
                    Console.Out.WriteLine(string.Format("Document ID {0} ({1} Display Column: {2})", queryResultItem.Document.ID.ToString(), queryResultItem.DisplayColumns.Count.ToString(), queryResultItem.DisplayColumns[0].Value.ToString()));
                }
            }
            Console.Out.WriteLine();
        }
        private static void KeywordQuery()
        {
            Console.Out.WriteLine("Attempting to execute a Keyword query...");
            string filePath = CONTENT_DOWNLOAD + "\\download.json";
            if (File.Exists(filePath))
            {
                Console.Out.WriteLine("Content download config file found: " + filePath);
                string inputJSON = File.ReadAllText(filePath);

                DocumentType documentType = app.Core.DocumentTypes.Find(DOCUMENT_TYPE);
                if (documentType == null)
                {
                    throw new Exception("Document type was not found");
                }

                DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
                documentQuery.AddDisplayColumn(DisplayColumnType.DocumentDate);
                documentQuery.AddDocumentType(documentType);

                KeywordRecordType keywordRecordType = documentType.KeywordRecordTypes[0];
                IList<JToken> contents = JToken.Parse(inputJSON)["contents"].Children().ToList();

                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();
                    foreach (var kt in keywordRecordType.KeywordTypes)
                    {
                        if (jContent.keywords.ContainsKey(kt.Name))
                        {
                            documentQuery.AddKeyword(kt.CreateKeyword(jContent.keywords[kt.Name]));
                        }
                    }      

                    using (QueryResult queryResults = documentQuery.ExecuteQueryResults(10L))
                    {
                        Console.Out.WriteLine("Displaying first 10 documents returned.");

                        foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                        {
                            Console.Out.WriteLine(string.Format("Document ID {0} ({1} Display Column: {2})", queryResultItem.Document.ID.ToString(), queryResultItem.DisplayColumns.Count.ToString(), queryResultItem.DisplayColumns[0].Value.ToString()));
                        }
                    }
                }
            }
            Console.Out.WriteLine();
        }
        private static void ExportDocument()
        {
            Console.Out.WriteLine("Attempting to export document...");

            foreach (var docID in documentIdList)
            {
                Document document = app.Core.GetDocumentByID(docID);
                if (document == null)
                {
                    throw new Exception("Document was not found");
                }

                using (DocumentLock documentLock = document.LockDocument())
                {
                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                    {
                        throw new Exception("Failed to lock document");
                    }

                    DefaultDataProvider defaultDataProvider = app.Core.Retrieval.Default;

                    using (PageData pageData = defaultDataProvider.GetDocument(document.DefaultRenditionOfLatestRevision))
                    {
                        using (Stream stream = pageData.Stream)
                        {
                            Utility.WriteStreamToFile(stream, string.Format(CONTENT_DOWNLOAD + @"\{0}.{1}", document.ID.ToString(), pageData.Extension));
                        }
                    }
                }

                Console.Out.WriteLine("Document export was successful.");
            }
            Console.Out.WriteLine();
        }

        private static void Connect()
        {
            app = null;

            AuthenticationProperties authProps = Application.CreateOnBaseAuthenticationProperties(APP_SERVER_URL, USERNAME, PASSWORD, DATA_SOURCE);
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
                        DocumentLookup();
                        ExecuteQuery();
                        DocumentQuery();
                        KeywordQuery();
                        ExportDocument();
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
