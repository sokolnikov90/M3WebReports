using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;

namespace M3Reports
{
    public class ReportTaskCreate
    {
        public struct Info
        {
            public int isError;
            public string status;
            public ReportTask data;
        }

        public
            Info info = new Info();

        public void ParseMessage(XmlNode messageNode)
        {
            this.info.isError = 0;
            
            try
            {
                this.info.status = messageNode.SelectSingleNode("Request/Status").InnerText.Trim();

                if (this.info.status == "SUCCESS")
                {
                    this.info.data = new ReportTask();

                    this.info.data.id = Convert.ToInt32(messageNode.SelectSingleNode("Request/Id").InnerText.Trim());
                    this.info.data.userId = messageNode.SelectSingleNode("Request/UserId").InnerText.Trim();
                    this.info.data.type = messageNode.SelectSingleNode("Request/Type").InnerText.Trim();
                    this.info.data.name = messageNode.SelectSingleNode("Request/Name").InnerText.Trim();
                    this.info.data.description = messageNode.SelectSingleNode("Request/Description").InnerText.Trim();
                    this.info.data.timeCreated = messageNode.SelectSingleNode("Request/TimeCreated").InnerText.Trim();
                    this.info.data.active = messageNode.SelectSingleNode("Request/Active").InnerText.Trim();
                    this.info.data.repeatMon = messageNode.SelectSingleNode("Request/RepeatMon").InnerText.Trim();
                    this.info.data.repeatTue = messageNode.SelectSingleNode("Request/RepeatTue").InnerText.Trim();
                    this.info.data.repeatWed = messageNode.SelectSingleNode("Request/RepeatWed").InnerText.Trim();
                    this.info.data.repeatThu = messageNode.SelectSingleNode("Request/RepeatThu").InnerText.Trim();
                    this.info.data.repeatFri = messageNode.SelectSingleNode("Request/RepeatFri").InnerText.Trim();
                    this.info.data.repeatSat = messageNode.SelectSingleNode("Request/RepeatSat").InnerText.Trim();
                    this.info.data.repeatSun = messageNode.SelectSingleNode("Request/RepeatSun").InnerText.Trim();
                    this.info.data.repeatMonth = messageNode.SelectSingleNode("Request/RepeatMonth").InnerText.Trim();
                    this.info.data.repeatYear = messageNode.SelectSingleNode("Request/RepeatYear").InnerText.Trim();
                    this.info.data.runTime = messageNode.SelectSingleNode("Request/RunTime").InnerText.Trim();
                    this.info.data.fromTime = messageNode.SelectSingleNode("Request/FromTime").InnerText.Trim();
                    this.info.data.toTime = messageNode.SelectSingleNode("Request/ToTime").InnerText.Trim();
                    this.info.data.interval = messageNode.SelectSingleNode("Request/Interval").InnerText.Trim();
                    this.info.data.format = messageNode.SelectSingleNode("Request/Format").InnerText.Trim();
                    this.info.data.userForId = messageNode.SelectSingleNode("Request/UserForId").InnerText.Trim();
                    this.info.data.criticalType = messageNode.SelectSingleNode("Request/CriticalType").InnerText.Trim();

                    XmlNodeList atmsIdsNodes = messageNode.SelectNodes("Request/AtmsIds/Id");

                    this.info.data.atmsIds = new List<string>();

                    for (int i = 0; i < atmsIdsNodes.Count; i++)
                    {
                        this.info.data.atmsIds.Add(atmsIdsNodes[i].InnerText.Trim());
                    }

                    XmlNodeList groupsIdsNodes = messageNode.SelectNodes("Request/GroupsIds/Id");

                    this.info.data.groupsIds = new List<string>();

                    for (int i = 0; i < groupsIdsNodes.Count; i++)
                    {
                        this.info.data.groupsIds.Add(groupsIdsNodes[i].InnerText.Trim());
                    }

                    XmlNodeList usersIdsNodes = messageNode.SelectNodes("Request/UsersIds/Id");

                    this.info.data.usersIds = new List<string>();

                    for (int i = 0; i < usersIdsNodes.Count; i++)
                    {
                        this.info.data.usersIds.Add(usersIdsNodes[i].InnerText.Trim());
                    }

                    if (this.info.data.type == "IncidentsHistoryCurrent") this.info.data.fromTime = this.info.data.toTime;
                }
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportTaskCreate() exeption: " + exception.Message);
                this.info.isError = 1;
            }
        }
    }
}