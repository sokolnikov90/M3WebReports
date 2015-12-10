using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

namespace M3Reports
{
    public class ReportUsersGetFacade : ReportDataProvider
    {
        private new ReportUsers report;

        public ReportUsersGetFacade(string ip, int port, string login, string password, string usersGetJSON)
            : base(ip, port, login, password, usersGetJSON)
        {
            base.report = this.report = new ReportUsers();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(usersGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(Queries.QueryUsersHistoryGet(this.report.Info.from, this.report.Info.to,
                new List<string>() { this.report.Info.userId }), this.ewh);

            this.connection.Write(Queries.QueryGetFunctions(), this.ewh);

            this.connection.Write(Queries.QueryGetUsers(), this.ewh);
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CAuditHistoryInfo":
                    this.report.ParseUserActions(messageNode);
                    break;
                case "CFunctionsQuery":
                    this.report.ParseFuncDescriptions(messageNode);
                    break;
                case "CUsersQuery":
                    this.report.ParseUserDescriptions(messageNode);
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}