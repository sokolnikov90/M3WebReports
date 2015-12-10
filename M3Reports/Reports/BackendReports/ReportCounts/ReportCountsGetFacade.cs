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
    using M3Utils;

    public class ReportCountsGetFacade : ReportDataProvider
    {
        private new readonly ReportCounts report;

        private string executeFunctionName;

        public ReportCountsGetFacade(string ip, int port, string login, string password, string countsGetJSON)
            : base(ip, port, login, password, countsGetJSON)
        {
            base.report = this.report = new ReportCounts();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(countsGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            IEnumerable<string> atmIdsD912 = (from data in this.report.Data.AtmInfo
                                              where data.TreeId == "1"
                                              select data.Id);

            IEnumerable<string> atmIdsNDC = (from data in this.report.Data.AtmInfo
                                             where data.TreeId == "2"
                                             select data.Id);

            this.executeFunctionName = "GetAtmsCounts";
            IEnumerable<string> countsAttributesName = StringHelper.GenerateAtributes(
                new[] { "Cassete_{0}_Currency", "Cassete_{0}_Value", "Total_RemainCass_Pos{0}", "Total_LoadCass_Pos{0}" },
                startIndex: 1, endIndex: 4);
            this.report.Data.AtmCounts = new List<CountsGet.AtmCountsData>();

            if (!atmIdsD912.IsNullOrEmpty())
            {
                this.connection.Write(M3Atms.Queries.QueryAttrsValueByNameQuery(atmIdsD912, "1", countsAttributesName), this.ewh);
            }

            if (!atmIdsNDC.IsNullOrEmpty())
            {
                this.connection.Write(M3Atms.Queries.QueryAttrsValueByNameQuery(atmIdsNDC, "2", countsAttributesName), this.ewh);
            }

            this.executeFunctionName = "GetAtmsBNACounts";
            IEnumerable<string> BNAAttributesName = new[] { "BNACounts" };
            this.report.Data.AtmBNACounts = new List<BNACountsGet.AtmBNACountsData>();

            if (!atmIdsD912.IsNullOrEmpty())
            {
                this.connection.Write(M3Atms.Queries.QueryAttrsValueByNameQuery(atmIdsD912, "1", BNAAttributesName), this.ewh);
            }

            if (!atmIdsNDC.IsNullOrEmpty())
            {
                this.connection.Write(M3Atms.Queries.QueryAttrsValueByNameQuery(atmIdsNDC, "2", BNAAttributesName), this.ewh);
            }
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "DMGroupAttrs":
                    switch (this.executeFunctionName)
                    {
                        case "GetAtmsCounts":
                            this.report.Data.AtmCountsGet = new CountsGet();
                            this.report.Data.AtmCountsGet.ParseMessage(messageNode);
                            if (this.report.Data.AtmCountsGet.info.isError == 0)
                            {
                                this.report.Data.AtmCounts.AddRange(this.report.Data.AtmCountsGet.info.data);
                            }
                            break;
                        case "GetAtmsBNACounts":
                            this.report.Data.AtmBNACountsGet = new BNACountsGet();
                            this.report.Data.AtmBNACountsGet.ParseMessage(messageNode);
                            if (this.report.Data.AtmBNACountsGet.info.isError == 0)
                            {
                                this.report.Data.AtmBNACounts.AddRange(this.report.Data.AtmBNACountsGet.info.data);
                            }
                            break;
                    }
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}