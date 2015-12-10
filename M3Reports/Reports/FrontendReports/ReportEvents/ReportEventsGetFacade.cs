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

    public class ReportEventsGetFacade : ReportDataProvider
    {
        private new ReportEvents report;

        public ReportEventsGetFacade(string ip, int port, string login, string password, string eventsGetJSON)
            : base(ip, port, login, password, eventsGetJSON)
        {
            base.report = this.report = new ReportEvents();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(eventsGetJSON);
        }

        public override void DoStuff()
        {
            if (this.signin.info.isError != 0)
                goto cleanup;

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(Queries.QueryEventsHistoryGet(this.report.Info.from, this.report.Info.to, this.report.Info.atmsId), this.ewh);

            var evtIds = this.report.eventItems.Select(eventItem => eventItem.Id).Distinct();

            this.connection.Write(Queries.QueryGetEvents(evtIds));

            this.connection.Disconnect();

            Thread.CurrentThread.CurrentCulture = Thread.CurrentThread.CurrentUICulture = ReportBuilder.GetCultureInfo(this.report.Info.languageCode);

            switch (this.report.Info.format)
            {
                case "xlsx":
                    this.report.MakeAnExcel();
                    break;
            }

            cleanup:
                if (this.signin.info.isError != 0) this.report.Info.isError = 1;
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CEventsInfo":
                    this.report.ParseEvents(messageNode);
                    break;
                case "CEvtsInfo":
                    this.report.ParseEventsDescriptions(messageNode);
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}