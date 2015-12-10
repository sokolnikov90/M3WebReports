using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

namespace M3Reports
{
    public class ReportTaskDeleteFacade : M3IPClient.M3UserSession, M3IPClient.IAsyncRequestFacade
    {
        ReportTaskDelete reportTaskDelete = new ReportTaskDelete();

        List<string> taskIds;

        public ReportTaskDeleteFacade(string ip, int port, string login, string password, List<string> taskIds)
            : base(ip, port, login, password)
        {
            this.taskIds = taskIds;

            this.connection.ReadEvent += this.IPRead;
        }

        public void DoStuff()
        {
            if (this.signin.info.isError == 0)
            {
                this.connection.Write(Queries.ReportTaskDelete(this.taskIds));

                this.ewh.Reset();
                this.ewh.WaitOne();

                if (this.signin.info.isError != 0) this.reportTaskDelete.info.isError = 1;
            }
            else
            {
                this.reportTaskDelete.info.isError = 1;
            }

            this.connection.Disconnect();
        }

        public string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.reportTaskDelete.info);
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
                   case "CWebReportTaskDeleteInfo":
                        this.reportTaskDelete.ParseMessage(messageNode);
                        break;
                }
            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                M3Utils.Log.Instance.Info("ReportTaskDeleteFacade.IPRead(...) exception:");
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