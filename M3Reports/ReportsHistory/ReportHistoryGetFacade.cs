using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

namespace M3Reports
{
    using M3Utils;

    public class ReportHistoryGetFacade : M3IPClient.M3UserSession, M3IPClient.IAsyncRequestFacade
    {
        ReportHistoryGet reportHistoryGet = new ReportHistoryGet();

        ReportHistory reportHistory;

        public ReportHistoryGetFacade(string ip, int port, string login, string password, string getJSON)
            : base(ip, port, login, password, getJSON)
        {
            this.reportHistory = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportHistory>(getJSON);

            this.connection.ReadEvent += this.IPRead;
        }

        public void DoStuff()
        {
            if (this.signin.info.isError == 0)
            {
                this.connection.Write(Queries.ReportHistoryGet(this.reportHistory));

                this.ewh.Reset();
                this.ewh.WaitOne();

                if (this.signin.info.isError != 0) this.reportHistoryGet.info.isError = 1;
            }
            else
            {
                this.reportHistoryGet.info.isError = 1;
            }

            this.connection.Disconnect();
        }

        public string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.reportHistoryGet.info);
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
                    case "CWebReportHistoryGetInfo":
                        this.reportHistoryGet.ParseMessage(messageNode);
                        break;
                }
            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                Log.Instance.Info("ReportHistoryGetFacade.IPRead(...) exception:");
                Log.Instance.Info(exp.Message);
                Log.Instance.Info(exp.Source);
                Log.Instance.Info(exp.StackTrace);
            }
            finally
            {
                this.ewh.Set();
            }
        }
    }
}