using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    using System.Globalization;

    public class ReportMessagesGetFacade : ReportDataProvider
    {
        public ReportMessagesGetFacade(string ip, int port, string login, string password, string getJSON)
            : base(ip, port, login, password, getJSON)
        {
            this.report = new ReportMessagesGet();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(getJSON);
        }

        public override void DoStuff()
        {
            if (this.signin.info.isError != 0)
                goto cleanup;

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(Queries.QueryMessageReport(this.report.Info.atmsId, this.report.Info.from, this.report.Info.to), this.ewh);

            this.connection.Disconnect();

            Thread.CurrentThread.CurrentCulture = Thread.CurrentThread.CurrentUICulture = ReportBuilder.GetCultureInfo(this.report.Info.languageCode);

            switch (this.report.Info.format)
            {
                case "xlsx":
                    this.report.MakeAnExcel();
                    break;
            }

        cleanup:
            if (this.signin.info.isError != 0)
                this.report.Info.isError = 1;
        }

        public override string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.report.Data.MessageHistoryGet.info);
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CMessageHistoryInfo":
                    this.report.Data.MessageHistoryGet = new MessageHistoryGet();
                    this.report.Data.MessageHistoryGet.ParseMessage(messageNode);
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}