using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Hyland.Unity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using log4net;
using System.Net.Http;

namespace RSIOnBaseUnity
{
    public class UnityBroker
    {
        private static readonly ILog logger = LogManager.GetLogger(typeof(UnityBroker));

        public static Application Connect(string url, string username, string password, string datasource, string domain)
        {
            Application app = null;

            try
            {
                if (!String.IsNullOrEmpty(domain) && !String.IsNullOrEmpty(username) && !String.IsNullOrEmpty(password))
                {
                    DomainAuthenticationProperties authProps = Application.CreateDomainAuthenticationProperties(url, datasource);
                    logger.Info("Impersonation : " + domain + " " + username + " " + password);
                    using (new Impersonation(domain, username, password))
                    {
                        string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                        logger.Info("Attempting to make a Domain connection using Principal.WindowsIdentity.GetCurrent().Name: " + userName);

                        string vUserName2 = Environment.UserDomainName + "\\" + Environment.UserName;
                        logger.Info("Environment.UserDomainName\\Environment.UserName: " + vUserName2);

                        app = Application.Connect(authProps);
                    }                      
                }
                else if (String.IsNullOrEmpty(username) && String.IsNullOrEmpty(password))
                {
                    DomainAuthenticationProperties authProps = Application.CreateDomainAuthenticationProperties(url, datasource);

                    string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    logger.Info("Attempting to make a Domain connection using Principal.WindowsIdentity.GetCurrent().Name: " + userName);

                    string vUserName2 = Environment.UserDomainName + "\\" + Environment.UserName;
                    logger.Info("Environment.UserDomainName\\Environment.UserName: " + vUserName2);        

                    app = Application.Connect(authProps);
                }
                else
                {
                    AuthenticationProperties authProps = Application.CreateOnBaseAuthenticationProperties(url, username, password, datasource);
                    authProps.LicenseType = LicenseType.Default;
                    logger.Info("Attempting to make a OnBase User connection...");
                    app = Application.Connect(authProps);
                }
            }
            catch (MaxLicensesException ex)
            {
                logger.Error("Error: All available licenses have been consumed.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (SystemLockedOutException ex)
            {
                logger.Error("Error: The system is currently in lockout mode.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (InvalidLoginException ex)
            {
                logger.Error("Error: Invalid Login Credentials.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (AuthenticationFailedException ex)
            {
                logger.Error("Error: NT Authentication Failed.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (MaxConcurrentLicensesException ex)
            {
                logger.Error("Error: All concurrent licenses for this user group have been consumed.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (InvalidLicensingException ex)
            {
                logger.Error("Error: Invalid Licensing.");
                logger.Error(ex.Message, ex);
                throw ex;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
                throw ex;
            }

            if (app != null)
            {
                logger.Info("Connection Successful. Connection ID: " + app.SessionID);
            }
            return app;
        } 
        public static void Disconnect(Application app)
        {
            logger.Info("Attempting to close connection...");

            try
            {
                app.Disconnect();
            }
            catch (SessionNotFoundException)
            {
                logger.Error("Error: Active session could not be found.");
            }
            finally
            {
                app.Dispose();
            }

            logger.Info("Connection closed.");
        }
        
        public static void GetConfigInfo(Application app, string documentsDir)
        {
            logger.Info("Attempting to Get Config Info");

            List<RSIDocumentTypeGroup> rsiDTGs = new List<RSIDocumentTypeGroup>();
            foreach (var dtg in app.Core.DocumentTypeGroups)
            {
                logger.Info("Document Type Group: " + dtg.Name + " (ID: " + dtg.ID + ")");
                RSIDocumentTypeGroup rsiDTG = new RSIDocumentTypeGroup();
                rsiDTG.ID = dtg.ID.ToString();
                rsiDTG.Name = dtg.Name;              
                List<RSIDocumentType> rsiDocTypes = new List<RSIDocumentType>();

                foreach (var dt in dtg.DocumentTypes)
                {
                    logger.Info("Document Type: " + dt.Name + " (ID: " + dt.ID + ")");
                    RSIDocumentType rsiDT = new RSIDocumentType();
                    rsiDT.ID = dt.ID.ToString();
                    rsiDT.Name = dt.Name;            
                    List<RSIKeywordType> rsiKeyTypes = new List<RSIKeywordType>();

                    foreach (var krt in dt.KeywordRecordTypes)
                    {
                        logger.Info("Keyword Types:");
                        foreach (var kt in krt.KeywordTypes)
                        {   
                            logger.Info(kt.Name + " (ID: " + kt.ID + " Type: " + kt.DataType.ToString()  + ")");
                            RSIKeywordType rsiKeywordType = new RSIKeywordType();
                            rsiKeywordType.ID = kt.ID.ToString();
                            rsiKeywordType.Name = kt.Name;
                            rsiKeywordType.Type = kt.DataType.ToString();
                            rsiKeywordType.Required = kt.FullFieldRequired.ToString();
                            rsiKeyTypes.Add(rsiKeywordType);
                        }
                    }
                    rsiDT.KeywordTypes = rsiKeyTypes;
                    rsiDocTypes.Add(rsiDT);          
                    logger.Info("");
                }

                rsiDTG.DocumentTypes = rsiDocTypes;
                rsiDTGs.Add(rsiDTG);
            }

            JArray jArray = (JArray)JToken.FromObject(rsiDTGs);
            string filePath = documentsDir + "\\config.json";
            using (StreamWriter file = File.CreateText(filePath))
            using (JsonTextWriter writer = new JsonTextWriter(file))
            {
                jArray.WriteTo(writer);
            }

            logger.Info("Keyword Record Types:");
            foreach (var krt in app.Core.KeywordRecordTypes)
            {
                logger.Info(krt.Name);
            }

            logger.Info("");
            logger.Info("File Types:");
            foreach (var ft in app.Core.FileTypes)
            {
                logger.Info(ft.Name);
            }


        }
        public static List<long> QueryDocuments(Application app, string documentsDir)
        {
            logger.Info("Attempting to execute a document query...");
            List<long> documentIdList = new List<long>();
            string filePath = documentsDir + "\\query.json";
            if (File.Exists(filePath))
            {
                logger.Info("Query File found: " + filePath);
                string inputJSON = File.ReadAllText(filePath);

                IList<JToken> jTokens = JToken.Parse(inputJSON)["contents"].Children().ToList();
                foreach (JToken jToken in jTokens)
                {
                    Content content = jToken.ToObject<Content>();

                    DocumentType documentType = app.Core.DocumentTypes.Find(content.documentType);

                    if (documentType == null)
                    {
                        throw new Exception("Document type was not found");
                    }

                    DocumentQuery documentQuery = app.Core.CreateDocumentQuery();
                    documentQuery.AddDisplayColumn(DisplayColumnType.AuthorName);
                    documentQuery.AddDisplayColumn(DisplayColumnType.DocumentDate);
                    documentQuery.AddDisplayColumn(DisplayColumnType.ArchivalDate);
                    documentQuery.AddDocumentType(documentType);

                    KeywordRecordType keywordRecordType = documentType.KeywordRecordTypes[0];
                    
                    foreach (var kt in keywordRecordType.KeywordTypes)
                    {
                        if (content.keywords.ContainsKey(kt.Name))
                        {   
                            switch (kt.DataType)
                            {
                                case KeywordDataType.AlphaNumeric:
                                    documentQuery.AddKeyword(kt.CreateKeyword(content.keywords[kt.Name]));
                                    break;
                                case KeywordDataType.Currency:
                                case KeywordDataType.SpecificCurrency:
                                case KeywordDataType.Numeric20:
                                    documentQuery.AddKeyword(kt.CreateKeyword(decimal.Parse(content.keywords[kt.Name])));
                                    break;
                                case KeywordDataType.Date:
                                case KeywordDataType.DateTime:
                                    documentQuery.AddKeyword(kt.CreateKeyword(DateTime.Parse(content.keywords[kt.Name]))); 
                                    break;
                                case KeywordDataType.FloatingPoint:
                                    documentQuery.AddKeyword(kt.CreateKeyword(double.Parse(content.keywords[kt.Name])));   
                                    break;
                                case KeywordDataType.Numeric9:
                                    documentQuery.AddKeyword(kt.CreateKeyword(long.Parse(content.keywords[kt.Name])));     
                                    break;
                            }
                        }
                    }

                    using (QueryResult queryResults = documentQuery.ExecuteQueryResults(long.MaxValue))
                    {                   
                        logger.Info("Number of " + content.documentType + " Documents Found: " + queryResults.QueryResultItems.Count().ToString());  
                        foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                        {
                            documentIdList.Add(queryResultItem.Document.ID);
                            logger.Info(string.Format("Document ID: {0}", queryResultItem.Document.ID.ToString()));
                            logger.Info(string.Format("Author: {0}, Document Date: {1}, Archival Date: {2}", queryResultItem.DisplayColumns[0].Value.ToString(), DateTime.Parse(queryResultItem.DisplayColumns[1].Value.ToString()).ToShortDateString(), DateTime.Parse(queryResultItem.DisplayColumns[2].Value.ToString()).ToShortDateString()));

                            foreach (var keyword in queryResultItem.Document.KeywordRecords[0].Keywords)
                            {
                                logger.Info(keyword.KeywordType.Name + " : " + keyword.Value.ToString());
                            }
                            logger.Info("");
                        }
                    }
                }
            }
            logger.Info("");
            return documentIdList;
        }
        public static void GetDocumentData(Application app, List<long> documentIdList, string documentsDir)
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
                            Utility.WriteStreamToFile(stream, string.Format(documentsDir + @"\{0}.{1}", document.ID.ToString(), pageData.Extension));
                        }
                    }
                }

                logger.Info("Document ID: " + docID + " export was successful.");
            }
            logger.Info("");
        }
        public static void ArchiveDocument(Application app, string documentsDir, string documentTypeGroup)
        {
            logger.Info("Attempting to archive documents...");
            try
            {                                       
                string filePath = documentsDir + "\\archive.json";
                if (File.Exists(filePath))
                {
                    logger.Info("Archive config file found: " + filePath);
                    string inputJSON = File.ReadAllText(filePath);

                    logger.Info("Attempting to Get Document Type Group: " + documentTypeGroup);
                    DocumentTypeGroup docTypeGroup = app.Core.DocumentTypeGroups.Find(documentTypeGroup);
                    if (docTypeGroup == null)
                    {
                        throw new Exception("Document Type Group not found: " + documentTypeGroup);
                    }
                    logger.Info("Document Type Group found: " + documentTypeGroup);

                    IList<JToken> jTokens = JToken.Parse(inputJSON)["contents"].Children().ToList();
                    foreach (JToken jToken in jTokens)
                    {
                        Content content = jToken.ToObject<Content>();
                        logger.Info("");
                        logger.Info("Attempting to Get Document Type: " + content.documentType);
                        DocumentType docType = docTypeGroup.DocumentTypes.Find(content.documentType);
                        if (docType == null)
                        {
                            throw new Exception("Document type was not found");
                        }

                        logger.Info("Attempting to Get File Type: " + content.fileTypes[0]);
                        FileType fType = app.Core.FileTypes.Find(content.fileTypes[0]);
                        if (fType == null)
                        {
                            throw new Exception("File type was not found");
                        }

                        KeywordRecordType keywordRecordType = docType.KeywordRecordTypes[0];

                        string fileUploadPath = documentsDir + "\\" + content.file;
                        if (File.Exists(fileUploadPath))
                        {
                            logger.Info("Archive document found: " + fileUploadPath);
                            List<string> fileList = new List<string>();
                            fileList.Add(fileUploadPath);

                            StoreNewDocumentProperties storeDocumentProperties = app.Core.Storage.CreateStoreNewDocumentProperties(docType, fType);
                            foreach (var kt in keywordRecordType.KeywordTypes)
                            {
                                if (content.keywords.ContainsKey(kt.Name))
                                {
                                    logger.Info("Add Keyword: " + kt.Name + " = " + content.keywords[kt.Name]);

                                    switch (kt.DataType)
                                    {
                                        case KeywordDataType.AlphaNumeric:
                                            storeDocumentProperties.AddKeyword(kt.CreateKeyword(content.keywords[kt.Name]));
                                            break;
                                        case KeywordDataType.Currency:
                                        case KeywordDataType.SpecificCurrency:
                                        case KeywordDataType.Numeric20:
                                            storeDocumentProperties.AddKeyword(kt.CreateKeyword(decimal.Parse(content.keywords[kt.Name])));
                                            break;
                                        case KeywordDataType.Date:
                                        case KeywordDataType.DateTime:                     
                                            storeDocumentProperties.AddKeyword(kt.CreateKeyword(DateTime.Parse(content.keywords[kt.Name])));
                                            break;
                                        case KeywordDataType.FloatingPoint:                        
                                            storeDocumentProperties.AddKeyword(kt.CreateKeyword(double.Parse(content.keywords[kt.Name])));
                                            break;     
                                        case KeywordDataType.Numeric9:                 
                                            storeDocumentProperties.AddKeyword(kt.CreateKeyword(long.Parse(content.keywords[kt.Name])));
                                            break;
                                    }                                     
                                }
                            }

                            storeDocumentProperties.DocumentDate = DateTime.Now;
                            storeDocumentProperties.Comment = "RSI OnBase Unity Application";
                            storeDocumentProperties.Options = StoreDocumentOptions.SkipWorkflow;

                            Document newDocument = app.Core.Storage.StoreNewDocument(fileList, storeDocumentProperties);

                            logger.Info(string.Format("Document Archive was successful. New Document ID: {0}", newDocument.ID.ToString()));
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
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }                                

            logger.Info("");
        }
        public static void Reindex(Application app, string documentsDir)
        {
            logger.Info("Attempting to re-index document by updating a keyword...");
            try
            {
                string filePath = documentsDir + "\\reindex.json";
                if (File.Exists(filePath))
                {
                    logger.Info("Re-index config file found: " + filePath);
                    string inputJSON = File.ReadAllText(filePath);

                    IList<JToken> jTokens = JToken.Parse(inputJSON)["contents"].Children().ToList(); 
                    foreach (JToken jToken in jTokens)
                    {
                        Content content = jToken.ToObject<Content>();

                        logger.Info("Attempting to Get Document ID: " + content.documentID);
                        long documentId = long.Parse(content.documentID);
                        Document document = app.Core.GetDocumentByID(documentId, DocumentRetrievalOptions.LoadKeywords);
                        if (document == null)
                        {
                            throw new Exception("Document was not found");
                        }

                        var keywords = document.KeywordRecords[0].Keywords.Where(x => content.keywords.Keys.Contains(x.KeywordType.Name));
                    
                        using (DocumentLock documentLock = document.LockDocument())
                        {
                            if (documentLock.Status != DocumentLockStatus.LockObtained)
                            {
                                throw new Exception("Failed to lock document");
                            }

                            KeywordModifier keyModifier = document.CreateKeywordModifier();

                            foreach (var keyword in keywords)
                            {
                                logger.Info("Update Keyword: " + content.keywords[keyword.KeywordType.Name]);
                                Keyword keywordToModify = null;

                                switch (keyword.KeywordType.DataType)
                                {
                                    case KeywordDataType.AlphaNumeric:
                                        keywordToModify = keyword.KeywordType.CreateKeyword(content.keywords[keyword.KeywordType.Name]);
                                        break;
                                    case KeywordDataType.Currency:
                                    case KeywordDataType.SpecificCurrency:
                                    case KeywordDataType.Numeric20:
                                        keywordToModify = keyword.KeywordType.CreateKeyword(decimal.Parse(content.keywords[keyword.KeywordType.Name]));
                                        break;
                                    case KeywordDataType.Date:
                                    case KeywordDataType.DateTime:
                                        keywordToModify = keyword.KeywordType.CreateKeyword(DateTime.Parse(content.keywords[keyword.KeywordType.Name]));
                                        break;
                                    case KeywordDataType.FloatingPoint:
                                        keywordToModify = keyword.KeywordType.CreateKeyword(double.Parse(content.keywords[keyword.KeywordType.Name]));
                                        break;
                                    case KeywordDataType.Numeric9:
                                        keywordToModify = keyword.KeywordType.CreateKeyword(long.Parse(content.keywords[keyword.KeywordType.Name]));
                                        break;
                                }
                                                                                                                                           
                                keyModifier.UpdateKeyword(keyword, keywordToModify);
                            }

                            keyModifier.ApplyChanges();
                        }

                        logger.Info(string.Format("Keyword was successfully updated. Document Id: {0}", content.documentID));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }

            logger.Info("");
        }

        /*
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

            string filePath = DOCUMENTS_DIRECTORY + "\\archive.json";
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

                    string fileUploadPath = DOCUMENTS_DIRECTORY + "\\" + jContent.file;
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
            string filePath = DOCUMENTS_DIRECTORY + "\\download.json";
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
                            Utility.WriteStreamToFile(stream, string.Format(DOCUMENTS_DIRECTORY + @"\{0}.{1}", document.ID.ToString(), pageData.Extension));
                        }
                    }
                }

                logger.Info("Document export was successful.");
            }
            logger.Info("");
        } 
        */
    }
}
