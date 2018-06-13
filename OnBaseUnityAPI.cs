using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Hyland.Unity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using log4net;

namespace RSIOnBaseUnity
{
    public class OnBaseUnityAPI
    {
        private static readonly ILog logger = LogManager.GetLogger(typeof(OnBaseUnityAPI));

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
            logger.Info("Attempting to Get Document Type Group: " + DOCUMENT_TYPE_GROUP);

            docTypeGroup = app.Core.DocumentTypeGroups.Find(DOCUMENT_TYPE_GROUP);
                         
            if (docTypeGroup == null)
            {
                throw new Exception("Document Type Group not found: " + DOCUMENT_TYPE_GROUP);
            }

            logger.Info("Document Type Group found: " + DOCUMENT_TYPE_GROUP);
        }
        private static void GetDocumentTypes()
        {
            logger.Info("Attempting to Get Document Types for Group: " + DOCUMENT_TYPE_GROUP);

            foreach (var dt in docTypeGroup.DocumentTypes)
            {
                logger.Info("Document Type: " + dt.Name + " (ID: " + dt.ID + ")");
                foreach (var krt in dt.KeywordRecordTypes)
                {                                                                
                    logger.Info("Keyword Types:");
                    foreach (var kt in krt.KeywordTypes)
                    {
                        logger.Info(kt.Name + " (ID: " + kt.ID + ")");
                    }                          
                }    
            }

            logger.Info("Keyword Record Types:");
            foreach (var krt in app.Core.KeywordRecordTypes)
            {
                logger.Info(krt.Name);
            }               
        } 
        private static void ArchiveDocument()
        {                                              
            logger.Info("Attempting to archive a document...");

            string filePath = CONTENT_UPLOAD + "\\upload.json";
            if (File.Exists(filePath))
            {
                logger.Info("Content upload config file found: " + filePath);

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
                documentIdList.Clear();

                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();

                    string fileUploadPath = CONTENT_UPLOAD + "\\" + jContent.file;
                    if (File.Exists(fileUploadPath))
                    {
                        logger.Info("Content upload file found: " + fileUploadPath);
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
                        logger.Info(string.Format("Document import was successful. New Document ID: {0}", newDocument.ID.ToString()));
                    }
                    else
                    {
                        logger.Info("Content upload file not found: " + fileUploadPath);
                    }
                }
            }
            else
            {
                logger.Info("Content upload config file not found: " + filePath);
            }
                                                               
            logger.Info("");
        }    
        private static void DocumentLookup()
        {
            logger.Info("Attempting to find document...");

            foreach (var docID in documentIdList)
            {
                Document document = app.Core.GetDocumentByID(docID);
                if (document == null)
                {
                    throw new Exception("Document was not found");
                }

                logger.Info(string.Format("Document was retrieved successfully. Document Id: {0}", document.ID.ToString()));
            }
            logger.Info("");
        }
        private static void ExecuteQuery()
        {
            logger.Info("Attempting to execute a document query...");

            DocumentType documentType = app.Core.DocumentTypes.Find(DOCUMENT_TYPE);
            if (documentType == null)
            {
                throw new Exception("Document type was not found");
            }

            DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
            documentQuery.AddDocumentType(documentType);

            DocumentList docList = documentQuery.Execute(long.MaxValue);

            logger.Info("Displaying first 10 documents returned.");
            logger.Info("");

            int limit = (docList.Count < 10) ? docList.Count : 10;
            documentIdList.Clear();

            for (int x = 0; x < limit; x++)
            {
                documentIdList.Add(docList[x].ID);
                logger.Info(string.Format("{0}. {1} {2}", (x + 1).ToString(), docList[x].ID, docList[x].DateStored.ToShortDateString())); 
            }
            logger.Info("");
        }
        private static void DocumentQuery()
        {
            logger.Info("Attempting to execute a document query...");

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
                logger.Info("Displaying first 10 documents returned.");
                documentIdList.Clear();

                foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                {       
                    documentIdList.Add(queryResultItem.Document.ID);
                    logger.Info(string.Format("Document ID {0} ({1} Display Column: {2})", queryResultItem.Document.ID.ToString(), queryResultItem.DisplayColumns.Count.ToString(), queryResultItem.DisplayColumns[0].Value.ToString()));
                }
            }
            logger.Info("");
        }
        private static void KeywordQuery()
        {
            logger.Info("Attempting to execute a Keyword query...");
            string filePath = CONTENT_DOWNLOAD + "\\download.json";
            if (File.Exists(filePath))
            {
                logger.Info("Content download config file found: " + filePath);
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
                documentIdList.Clear();

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
                        logger.Info("Displaying first 10 documents returned.");
                                                
                        foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                        {
                            documentIdList.Add(queryResultItem.Document.ID);
                            logger.Info(string.Format("Document ID {0} ({1} Display Column: {2})", queryResultItem.Document.ID.ToString(), queryResultItem.DisplayColumns.Count.ToString(), queryResultItem.DisplayColumns[0].Value.ToString()));
                        }
                    }
                }
            }
            logger.Info("");
        }
        private static void ExportDocument()
        {
            logger.Info("Attempting to export document...");

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

                logger.Info("Document export was successful.");
            }
            logger.Info("");
        }       
        private static void Connect()
        {
            app = null;

            AuthenticationProperties authProps = Application.CreateOnBaseAuthenticationProperties(APP_SERVER_URL, USERNAME, PASSWORD, DATA_SOURCE);
            authProps.LicenseType = LicenseType.Default;

            logger.Info("Attempting to make a connection...");

            try
            {
                app = Application.Connect(authProps);
            }
            catch (MaxLicensesException)
            {
                logger.Info("Error: All available licenses have been consumed."); 
            }
            catch (SystemLockedOutException)
            {
                logger.Info("Error: The system is currently in lockout mode.");   
            }
            catch (InvalidLoginException)
            {
                logger.Info("Error: Invalid Login Credentials.");                 
            }
            catch (UserAccountLockedException)
            {
                logger.Info("Error: This account has been temporarily locked.");  
            }
            catch (AuthenticationFailedException)
            {
                logger.Info("Error: NT Authentication Failed.");                  
            }
            catch (MaxConcurrentLicensesException)
            {
                logger.Info("Error: All concurrent licenses for this user group have been consumed.");
            }
            catch (InvalidLicensingException)
            {
                logger.Info("Error: Invalid Licensing.");                                             
            }

            if (app != null)
            {
                logger.Info("Connection Successful. Connection ID: " + app.SessionID);                
            }   
        }  
        private static void Disconnect()
        {
            logger.Info("Attempting to close connection...");

            try
            {
                app.Disconnect();
            }
            catch (SessionNotFoundException)
            {
                logger.Info("Error: Active session could not be found.");   
            }
            finally
            {
                app.Dispose();
            }

            logger.Info("Connection closed.");                              
        }    
        public static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            try
            {
                Connect();
                if (app != null)
                {
                    GetDocumentTypeGroup();
                    if (docTypeGroup != null)
                    {
                        logger.Info("Enter 0 stopping the application.");
                        logger.Info("Enter 1 for Getting list of Document Types with Keywords.");
                        logger.Info("Enter 2 for Uploading Content to OnBase using upload.json file.");
                        logger.Info("Enter 3 for Keyword Query using download.json file.");
                        logger.Info("Enter 4 for Execute Query.");
                        logger.Info("Enter 5 for Document Query.");
                        logger.Info("Enter 6 for Document Lookup (depends on options from 2 to 5 to get list of document ID's).");
                        logger.Info("Enter 7 for Downloading Content from OnBase (depends on options from 2 to 5 to get list of document ID's).");

                        int inputKey = -1;
                        Int32.TryParse(Console.ReadLine(), out inputKey);

                        while (inputKey != 0)
                        {                               
                            switch (inputKey)
                            {
                                case 0:
                                    break;
                                case 1:
                                    GetDocumentTypes();
                                    break;
                                case 2:
                                    ArchiveDocument();
                                    break;
                                case 3:
                                    KeywordQuery();
                                    break;
                                case 4:
                                    ExecuteQuery();
                                    break;
                                case 5:
                                    DocumentQuery();
                                    break;
                                case 6:
                                    DocumentLookup();
                                    break;
                                case 7:
                                    ExportDocument();
                                    break;
                            }

                            logger.Info("Enter 0 stopping the application.");
                            logger.Info("Enter 1 for Getting list of Document Types with Keywords.");
                            logger.Info("Enter 2 for Uploading Content to OnBase using upload.json file.");
                            logger.Info("Enter 3 for Keyword Query using download.json file.");
                            logger.Info("Enter 4 for Execute Query.");
                            logger.Info("Enter 5 for Document Query.");
                            logger.Info("Enter 6 for Document Lookup (depends on options from 2 to 5 to get list of document ID's).");
                            logger.Info("Enter 7 for Downloading Content from OnBase (depends on options from 2 to 5 to get list of document ID's).");

                            Int32.TryParse(Console.ReadLine(), out inputKey);
                        }   
                    }
                }                    
            }
            catch (WebException ex)
            {
                logger.Info("General network error: " + ex.Message);      
            }
            catch (UnityAPIException ex)
            {
                logger.Info("General Unity API error: " + ex.Message);    
            }
            catch (Exception ex)
            {
                logger.Info("General error: " + ex.Message);              
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
