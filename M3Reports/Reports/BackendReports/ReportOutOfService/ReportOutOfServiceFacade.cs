using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;

using M3Atms;
using M3Incidents;
using M3Dictionaries;

namespace M3Reports
{
    using M3Utils;

    public class ReportOutOfServiceFacade : ReportDataProvider
    {
        public ReportOutOfServiceFacade(string ip, int port, string login, string password, string incidentsGetJSON)
            : base(ip, port, login, password, incidentsGetJSON)
        {
            this.report = new ReportOutOfService();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(incidentsGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Statuses"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "UserRoles"), this.ewh);

            this.report.Data.QueryIncident = new IncidentGet();
            this.report.Data.QueryIncident.from = this.report.Info.from;
            this.report.Data.QueryIncident.to = this.report.Info.to;
            this.report.Data.QueryIncident.statusIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.info.statuses.data where ((item.isClosed == "0") || (item.isClosed != "1")) select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.atmIds = this.report.Info.atmsId;
            this.report.Data.QueryIncident.typeIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.info.types.data select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.userRoleIds = (from item in this.report.Data.DictionariesGet.info.userRoles.data where item.appType == "M3Web" select item.id.ToString()).ToList();
            this.report.Data.QueryIncident.criticalType = (BankName == "RNCB") ? "critical" : this.report.Info.criticalType;

            this.connection.Write(M3Incidents.Queries.IncidentsGet(this.report.Data.QueryIncident));

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "IncidentsRules"), this.ewh);            

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "DevicesTypes"), this.ewh);

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(M3Atms.Queries.QueryAtmGroups(new [] {Convert.ToInt32(this.report.Info.userId)}));
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CUsersGroupsAtmsInfo":
                    this.report.Data.AtmGroupsGet.ParseMessage(messageNode);
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}
