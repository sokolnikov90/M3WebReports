using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

using M3Atms;
using M3Incidents;
using M3Dictionaries;
using System.IO;

namespace M3Reports
{
    public class ReportSummaryIncidentsOpenFacade : ReportDataProvider
    {
        public ReportSummaryIncidentsOpenFacade(string ip, int port, string login, string password, string eventsGetJSON)
            : base(ip, port, login, password, eventsGetJSON)
        {
            this.report = new ReportSummaryIncidentsOpen();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(eventsGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Atms.Queries.QueryAtmGroups(new [] { this.signin.info.userId }), this.ewh);

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Statuses"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "UserRoles"), this.ewh);

            this.report.Data.QueryIncident = new IncidentGet();
            this.report.Data.QueryIncident.from = DateTime.Parse(this.report.Info.from).AddMonths(-3).ToString("yyyy-MM-dd HH:mm:ss");
            this.report.Data.QueryIncident.to = this.report.Info.to;
            this.report.Data.QueryIncident.atmIds = this.report.Info.atmsId;
            this.report.Data.QueryIncident.typeIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Types select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.userRoleIds = (from item in this.report.Data.DictionariesGet.UserRoles where item.appType == "M3Web" select item.id.ToString()).ToList();
            this.report.Data.QueryIncident.criticalType = (BankName == "RNCB") ? "critical" : this.report.Info.criticalType;

            switch (this.report.Info.type)
            {
                case "SummaryIncidentsOpen":
                    // Статусы инцидентов - все, кроме "Открыт" и "Отменен"
                    this.report.Data.QueryIncident.statusIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Statuses where ((item.isClosed != "0") && (item.text != "Отменен")) select item.id.ToString()).ToArray());
                    break;
                case "SummaryIncidentsOpenWorking":
                    // Статусы инцидентов - только рабочие
                    this.report.Data.QueryIncident.statusIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Statuses where (item.isClosed == "1") select item.id.ToString()).ToArray());
                    break;
            }

            this.connection.Write(M3Incidents.Queries.IncidentsGet(this.report.Data.QueryIncident), this.ewh);
        }
    }
}