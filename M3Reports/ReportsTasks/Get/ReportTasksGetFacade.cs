using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;
using System.IO;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    public class ReportTasksGetFacade : M3IPClient.M3UserSession, M3IPClient.IAsyncRequestFacade
    {
        ReportTask reportTask;
        ReportTasksGet reportTasksGet = new ReportTasksGet();

        public ReportTasksGetFacade(string ip, int port, string login, string password, string reportTaskGetJSON)
            : base (ip, port, login, password, reportTaskGetJSON)
        {
            this.reportTask = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportTask>(reportTaskGetJSON);

            this.connection.ReadEvent += this.IPRead;
        }

        public void DoStuff()
        {
            if (this.signin.info.isError == 0)
            {
                this.connection.Write(Queries.ReportTaskGet(this.reportTask));

                this.ewh.Reset();
                this.ewh.WaitOne();

                if (this.signin.info.isError != 0)
                {
                    this.reportTasksGet.info.isError = 1;
                }
            }
            else
            {
                this.reportTasksGet.info.isError = 1;
            }

            this.connection.Disconnect();
        }

        public string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.reportTasksGet.info);
        }

        private void IPRead(string message, bool complit)
        {
            XmlNode messageNode;
            XmlDocument xmlDocument = new XmlDocument();

            try
            {
                xmlDocument.LoadXml(message);

                messageNode = xmlDocument.SelectSingleNode("Message");

                this.requestName = messageNode.SelectSingleNode("Request/./@name").InnerText;

                switch (this.requestName)
                {
                    case "CWebReportTaskGetInfo":
                        this.reportTasksGet.ParseMessage(messageNode);
                        break;
                }
            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                M3Utils.Log.Instance.Info("ReportTasksGetFacade.IPRead(...) exception:");
                M3Utils.Log.Instance.Info(exp.Message);
                M3Utils.Log.Instance.Info(exp.Source);
                M3Utils.Log.Instance.Info(exp.StackTrace);
            }
            finally
            {
                this.ewh.Set();
            }
        }
    }
}