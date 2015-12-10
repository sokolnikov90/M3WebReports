using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

namespace M3Reports
{
    public class ReportTaskChangeFacade : M3IPClient.M3UserSession, M3IPClient.IAsyncRequestFacade
    {
        ReportTask reportTask = new ReportTask();
        ReportTaskChange reportTaskChange = new ReportTaskChange();

        public ReportTaskChangeFacade(string ip, int port, string login, string password, string reportTaskChangeJSON)
            : base(ip, port, login, password)
        {
            this.reportTask = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportTask>(reportTaskChangeJSON);

            this.connection.ReadEvent += this.IPRead;
        }

        public void DoStuff()
        {
            if (this.signin.info.isError == 0)
            {
                this.connection.Write(Queries.ReportTaskChange(this.reportTask));

                this.ewh.Reset();
                this.ewh.WaitOne();

                if (this.signin.info.isError != 0) this.reportTaskChange.info.isError = 1;
            }
            else
            {
                this.reportTaskChange.info.isError = 1;
            }

            this.connection.Disconnect();
            
        }

        public string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.reportTaskChange.info);
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
                    case "CWebReportTaskChangeInfo":
                        this.reportTaskChange.ParseMessage(messageNode);
                        break;
                }
            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                M3Utils.Log.Instance.Info("ReportTaskChangeFacade.IPRead(...) exception:");
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