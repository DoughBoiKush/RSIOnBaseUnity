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
        private static string ARCHIVE_DIRECTORY = System.Configuration.ConfigurationManager.AppSettings["archiveDirectory"].ToString();
        private static string CONTENT_DOWNLOAD = System.Configuration.ConfigurationManager.AppSettings["contentDownload"].ToString();
        private static string QUERY_FILE = System.Configuration.ConfigurationManager.AppSettings["queryFile"].ToString();
        private static string REINDEX_FILE = System.Configuration.ConfigurationManager.AppSettings["reindexFile"].ToString();

        private static Application app;
        private static DocumentTypeGroup docTypeGroup;
        private static List<long> documentIdList = new List<long>();

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
                        string usage = "\n----------------------------------------------------- \n" +
                            "Select Option: \n" +
                            "1 : Get Config Info \n" +
                            "2 : Query Documents (using query.json) \n" +
                            "3 : Get Document Data (after #2 option) \n" +
                            "4 : Archive Document (using archive.json) \n" +
                            "5 : Reindex (using reindex.json) \n" +
                            "0 : Exit \n" +
                            "-----------------------------------------------------";

                        logger.Info(usage);
                        int inputKey = -1;
                        Int32.TryParse(Console.ReadLine(), out inputKey);

                        while (inputKey != 0)
                        {
                            switch (inputKey)
                            {
                                case 0:
                                    break;
                                case 1:
                                    GetConfigInfo();
                                    break;
                                case 2:
                                    QueryDocuments();
                                    break;
                                case 3:
                                    GetDocumentData();
                                    break;
                                case 4:
                                    ArchiveDocument();
                                    break;
                                case 5:
                                    Reindex();
                                    break; 
                            }

                            logger.Info(usage);
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
        private static void GetConfigInfo()
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
                logger.Info("");
            }

            logger.Info("Keyword Record Types:");
            foreach (var krt in app.Core.KeywordRecordTypes)
            {
                logger.Info(krt.Name);
            }
        }
        private static void QueryDocuments()
        {
            logger.Info("Attempting to execute a document query...");  
            if (File.Exists(QUERY_FILE))
            {
                logger.Info("Query File found: " + QUERY_FILE);
                string inputJSON = File.ReadAllText(QUERY_FILE);

                IList<JToken> contents = JToken.Parse(inputJSON)["contents"].Children().ToList();
                documentIdList.Clear();
                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();

                    DocumentType documentType = app.Core.DocumentTypes.Find(jContent.documentType);

                    if (documentType == null)
                    {
                        throw new Exception("Document type was not found");
                    }

                    DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
                    documentQuery.AddDisplayColumn(DisplayColumnType.DocumentDate);
                    documentQuery.AddDocumentType(documentType);

                    KeywordRecordType keywordRecordType = documentType.KeywordRecordTypes[0];
                    
                    foreach (var kt in keywordRecordType.KeywordTypes)
                    {
                        if (jContent.keywords.ContainsKey(kt.Name))
                        {
                            documentQuery.AddKeyword(kt.CreateKeyword(jContent.keywords[kt.Name]));
                        }
                    }

                    using (QueryResult queryResults = documentQuery.ExecuteQueryResults(long.MaxValue))
                    {
                        logger.Info("Documents returned:");

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
        private static void GetDocumentData()
        {
            logger.Info("Attempting to Get Document Data...");

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

                logger.Info("Document ID: " + docID + " export was successful.");
            }
            logger.Info("");
        }
        private static void ArchiveDocument()
        {
            logger.Info("Attempting to archive documents...");

            string filePath = ARCHIVE_DIRECTORY + "\\archive.json";
            if (File.Exists(filePath))
            {
                logger.Info("Archive config file found: " + filePath);    
                string inputJSON = File.ReadAllText(filePath);

                IList<JToken> contents = JToken.Parse(inputJSON)["contents"].Children().ToList(); 

                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();

                    DocumentType docType = docTypeGroup.DocumentTypes.Find(jContent.documentType);
                    if (docType == null)
                    {
                        throw new Exception("Document type was not found");
                    }

                    FileType fType = app.Core.FileTypes.Find(jContent.fileTypes[0]);
                    if (fType == null)
                    {
                        throw new Exception("File type was not found");
                    }

                    KeywordRecordType keywordRecordType = docType.KeywordRecordTypes[0];  
                               
                    string fileUploadPath = ARCHIVE_DIRECTORY + "\\" + jContent.file;
                    if (File.Exists(fileUploadPath))
                    {
                        logger.Info("Archive document found: " + fileUploadPath);
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
                                                                    
                        logger.Info(string.Format("Document import was successful. New Document ID: {0}", newDocument.ID.ToString()));
                    }
                    else
                    {
                        logger.Info("Archive document file not found: " + fileUploadPath);
                    }
                }
            }
            else
            {
                logger.Info("Archive config file not found: " + filePath);
            }

            logger.Info("");
        }
        private static void Reindex()
        {
            logger.Info("Attempting to re-index document by updating a keyword...");
                                                  
            if (File.Exists(REINDEX_FILE))
            {
                logger.Info("Archive config file found: " + REINDEX_FILE);
                string inputJSON = File.ReadAllText(REINDEX_FILE);

                IList<JToken> contents = JToken.Parse(inputJSON)["contents"].Children().ToList();

                foreach (JToken jToken in contents)
                {
                    Content jContent = jToken.ToObject<Content>();

                    DocumentType documentType = docTypeGroup.DocumentTypes.Find(jContent.documentType);
                    if (documentType == null)
                    {
                        throw new Exception("Document type was not found");
                    }                                       

                    long documentId = long.Parse(jContent.documentID);
                    Document document = app.Core.GetDocumentByID(documentId, DocumentRetrievalOptions.LoadKeywords);
                    if (document == null)
                    {
                        throw new Exception("Document was not found");
                    }

                    foreach (var kr in document.KeywordRecords)
                    {
                        foreach (var keyword in kr.Keywords)
                        {
                            foreach (var inputKeyword in jContent.keywords)
                            {
                                if (keyword.KeywordType.Name.Equals(inputKeyword.Key))
                                {                                                                                           
                                    using (DocumentLock documentLock = document.LockDocument())
                                    {
                                        if (documentLock.Status != DocumentLockStatus.LockObtained)
                                        {
                                            throw new Exception("Failed to lock document");
                                        }

                                        KeywordModifier keyModifier = document.CreateKeywordModifier();                                                               
                                        Keyword keywordToModify = keyword.KeywordType.CreateKeyword(inputKeyword.Value);                                                               
                                        keyModifier.UpdateKeyword(keyword, keywordToModify);  
                                        keyModifier.ApplyChanges();                           
                                    }
                                }
                            }
                        }
                    }

                    logger.Info(string.Format("Keyword was successfully updated. Document Id: {0}", jContent.documentID));
                }
            }

            logger.Info("");
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
                logger.Info("");
            }

            logger.Info("Keyword Record Types:");
            foreach (var krt in app.Core.KeywordRecordTypes)
            {
                logger.Info(krt.Name);
            }               
        } 
        private static void ArchiveDocument_Orig()
        {                                              
            logger.Info("Attempting to archive a document...");

            string filePath = ARCHIVE_DIRECTORY + "\\archive.json";
            if (File.Exists(filePath))
            {
                logger.Info("Content upload config file found: " + filePath);

                string inputJSON = File.ReadAllText(filePath);

                DocumentType docType = docTypeGroup.DocumentTypes.Find("RETURN - RETURN");
                if (docType == null)
                {
                    throw new Exception("Document type was not found");
                }

                FileType fType = app.Core.FileTypes.Find("PDF");
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

                    string fileUploadPath = ARCHIVE_DIRECTORY + "\\" + jContent.file;
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

            DocumentType documentType = app.Core.DocumentTypes.Find("RETURN - RETURN");
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

            DocumentType documentType = app.Core.DocumentTypes.Find("RETURN - RETURN");
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

                DocumentType documentType = app.Core.DocumentTypes.Find("RETURN - RETURN");
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
    }
}
