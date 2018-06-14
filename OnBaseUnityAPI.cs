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
        private static string DOCUMENTS_DIRECTORY = System.Configuration.ConfigurationManager.AppSettings["documentsDirectory"].ToString();      

        private static Application app;
        private static DocumentTypeGroup docTypeGroup;
        private static List<long> documentIdList = new List<long>();

        public static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            try
            {
                app = UnityBroker.Connect(APP_SERVER_URL, USERNAME, PASSWORD, DATA_SOURCE);
                if (app != null)
                {
                    docTypeGroup = UnityBroker.GetDocumentTypeGroup(app, DOCUMENT_TYPE_GROUP);
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
                                    UnityBroker.GetConfigInfo(app, DOCUMENT_TYPE_GROUP, docTypeGroup);
                                    break;
                                case 2:
                                    documentIdList = UnityBroker.QueryDocuments(app, DOCUMENTS_DIRECTORY);
                                    break;
                                case 3:
                                    UnityBroker.GetDocumentData(app, documentIdList, DOCUMENTS_DIRECTORY);
                                    break;
                                case 4:
                                    UnityBroker.ArchiveDocument(app, DOCUMENTS_DIRECTORY, docTypeGroup);
                                    break;
                                case 5:
                                    UnityBroker.Reindex(app, DOCUMENTS_DIRECTORY);
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
                    UnityBroker.Disconnect(app);
                }
            }
        }                
    }
}
