using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;
using System.Xml.Linq;

using M3Atms;
using M3Incidents;
using M3Dictionaries;
using System.Web.Configuration;

namespace M3Reports
{
    using M3IPClient;

    public class ReportAvailabilitiesGetFacade : ReportDataProvider
    {
        private new readonly ReportAvailabilities report;

        public ReportAvailabilitiesGetFacade(string ip, int port, string login, string password, string availsGetJSON)
            : base(ip, port, login, password, availsGetJSON)
        {
            base.report = this.report = new ReportAvailabilities();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(availsGetJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Statuses"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Users"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "UserRoles"), this.ewh);

            if (BankName == "RNCB") this.connection.Write(M3Atms.Queries.QueryAtmGroups(new [] { this.signin.info.userId }), this.ewh);

            this.report.Data.QueryIncident = new IncidentGet();
            this.report.Data.QueryIncident.from = DateTime.Parse(this.report.Info.from).AddMonths(-1).ToString("yyyy-MM-dd HH:mm:ss");
            this.report.Data.QueryIncident.to = this.report.Info.to;
            this.report.Data.QueryIncident.statusIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Statuses where item.text != "Отменен" select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.atmIds = this.report.Info.atmsId;
            this.report.Data.QueryIncident.typeIds = String.Join(", ", (from item in this.report.Data.DictionariesGet.Types select item.id.ToString()).ToArray());
            this.report.Data.QueryIncident.userRoleIds = (from item in this.report.Data.DictionariesGet.UserRoles where item.appType == "M3Web" select item.id.ToString()).ToList();
            this.report.Data.QueryIncident.criticalType = (BankName == "RNCB") ? "critical" : "all";

            this.connection.Write(M3Incidents.Queries.IncidentsGet(this.report.Data.QueryIncident), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "DevicesTypes"), this.ewh);

            this.connection.Write(M3Dictionaries.Queries.DictionaryGet(this.report.Info.languageCode, "Types"), this.ewh);

            if (BankName == "Absolute")
            {
                XDocument settDoc = XDocument.Load(AppDomain.CurrentDomain.BaseDirectory + @"bin/M3Reports/ReportAvailabilityColumns.xml");
                string attrId = settDoc.Element("Columns").Element("AttrSettings").Element("AtmMode").Value;
                string InService = settDoc.Element("Columns").Element("AttrSettings").Element("AtmMode").Attribute("InService").Value;
                string OutOfService = settDoc.Element("Columns").Element("AttrSettings").Element("AtmMode").Attribute("OutOfService").Value;

                this.connection.Write(Queries.QueryGetAttrHistory(this.report.Data.QueryIncident.from, this.report.Data.QueryIncident.to, this.report.Data.QueryIncident.atmIds, attrId), this.ewh);

                this.report.SearchAvailsForAbsolute();
            }
            else
            {
                this.report.SearchAvails();
            }               
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CAtmsInfo":
                    this.report.Data.AtmInfoGet = new InfoGet();
                    this.report.Data.AtmInfoGet.ParseMessage(messageNode);

                    foreach (Info atm in this.report.Data.AtmInfo)
                    {
                        if ((atm.CashDispense != null)&&(atm.CashDispense == "1"))
                        {
                            if ((atm.CashAccept != null)&&(atm.CashAccept == "1"))
                            {
                                this.report.hasBNAandDispenser++;
                                continue;
                            }
                            this.report.hasDispenser++;
                            continue;
                        }

                        if ((atm.CashAccept != null)&&(atm.CashAccept == "1"))
                        {
                            this.report.hasBna++;
                        }
                    }
                    break;
                case "CWebIncidentsGetInfo":
                    var incidents = new M3Incidents.GetAll();
                    incidents.ParseMessage(messageNode);

                    this.report.Data.IncidentsGet = new M3Incidents.GetAll(new List<Incident>());

                    if (incidents.incidentInfo.isError == 0)
                    {
                        foreach (Incident incident in incidents.incidentInfo.data)
                        {
                            if ((incident.timeClosed == string.Empty) || (DateTime.Parse(incident.timeClosed) > this.report.FromDate))
                            {
                                this.report.Data.Incidents.Add(incident);
                            }
                        }
                    }
                    else
                    {
                        this.report.Data.IncidentsGet.incidentInfo.isError = 1;
                    }
                    break;
                case "CAttributeHistoryInfo":
                    this.report.Data.Incidents.AddRange(this.GetIncidentsFromAttrHistory(messageNode.InnerText, this.report.AtmsId));
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }

        private List<Incident> GetIncidentsFromAttrHistory(string msg, List<string> atmsId)
        {
            List<Incident> result = new List<Incident>();

            XDocument doc = XDocument.Parse(msg);

            var items = doc.Element("Message").Element("Request").Elements("Item").OrderBy(m => m.Element("DTime").Value).ToList();
            foreach (var atm in atmsId)
            {
                var itemsForOneAtm = (from item in items
                                      where item.Element("AtmId").Value == atm
                                      orderby DateTime.Parse(item.Element("DTime").Value)
                                      select item).ToList();

                for (int i = 0; i < itemsForOneAtm.Count; i++)
                {
                    if (itemsForOneAtm[i].Element("AttrValue").Value == "13" && itemsForOneAtm[i + 1].Element("AttrValue").Value == "12")
                    {
                        result.Add(new Incident()
                        {
                            timeCreated = itemsForOneAtm[i].Element("DTime").Value,
                            timeClosed = itemsForOneAtm[i + 1].Element("DTime").Value,
                            deviceTypeId = "InOutService",
                            atmId = atm
                        });
                    }
                }
            }
            return result;
        }
    }
}