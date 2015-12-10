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

    public class ReportIncidentsPivotTableFacade : ReportDataProvider
    {
        public ReportIncidentsPivotTableFacade(string ip, int port, string login, string password, string reportRunJSON)
            : base(ip, port, login, password, reportRunJSON)
        {
            this.report = new ReportIncidentsPivotTable();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(reportRunJSON);
        }

        public override void DoStuff()
        {
            if (this.signin.info.isError != 0)
                goto cleanup;

            for (int i = 0; i < this.report.Info.incidents.Count; i++)
            {
                this.report.Info.incidents[i].comments = HttpUtility.UrlDecode(this.report.Info.incidents[i].comments);
            }

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.incidents.Select(x => x.atmId).Distinct().ToList()));

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode,"Statuses"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Users"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "ResponsibleFor"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "UserRoles"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "IncidentsRules"), this.ewh);

            this.connection.Disconnect();

            this.report.Info.from = this.report.Info.to = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

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
    }
}