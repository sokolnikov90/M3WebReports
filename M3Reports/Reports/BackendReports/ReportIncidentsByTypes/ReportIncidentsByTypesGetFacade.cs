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
    public class ReportIncidentsByTypesGetFacade : ReportDataProvider
    {
        public ReportIncidentsByTypesGetFacade(string ip, int port, string login, string password, string incidentsGetJSON)
            : base(ip, port, login, password, incidentsGetJSON)
        {
            this.report = new ReportIncidentsByTypes();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(incidentsGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Statuses"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Users"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "UserRoles"), this.ewh);   

            this.report.Data.QueryIncident = new IncidentGet();
            this.report.Data.QueryIncident.from = DateTime.Parse(this.report.Info.from).AddMonths(-1).ToString("yyyy-MM-dd HH:mm:ss");
            this.report.Data.QueryIncident.to = this.report.Info.to;
            this.report.Data.QueryIncident.statusIds = String.Join(", ", 
                (from item in this.report.Data.DictionariesGet.Statuses 
                 where ((item.id != 1) && (item.id != 3) && (Convert.ToInt32(item.isClosed) == this.GetIsClosed()))
                 select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.atmIds = this.report.Info.atmsId;
            this.report.Data.QueryIncident.typeIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Types select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.userRoleIds = (from item in this.report.Data.DictionariesGet.UserRoles where item.appType == "M3Web" select item.id.ToString()).ToList();
            this.report.Data.QueryIncident.criticalType = (BankName == "RNCB") ? "critical" : this.report.Info.criticalType;

            this.connection.Write(M3Incidents.Queries.IncidentsGet(this.report.Data.QueryIncident), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "IncidentsRules"), this.ewh);  

            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);
        }

        private int GetIsClosed()
        {
            int isClosed;

            switch (this.report.Info.type)
            {
                case "IncidentsHistoryCurrent":
                    {
                        isClosed = 1;
                    }
                    break;
                case "IncidentsHistoryRange":
                    {
                        isClosed = 2;
                    }
                    break;
                default:
                    {
                        isClosed = 0;
                    }
                    break;
            }

            return isClosed;
        }
    }
}