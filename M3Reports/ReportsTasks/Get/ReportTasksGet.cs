using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;
using System.Text;

namespace M3Reports
{
    public class ReportTasksGet
    {
        public struct Info
        {
            public int isError;
            public List<ReportTask> data;
        }

        public
            Info info = new Info();

        public void ParseMessage(XmlNode messageNode)
        {
            this.info.isError = 0;

            try
            {
                StringBuilder stringBuilder = new StringBuilder();

                stringBuilder.Append("<Message>");
                stringBuilder.Append(M3Utils.IOHelper.UnzipBytes(Convert.FromBase64String(messageNode.SelectSingleNode("Request/ZippedData").InnerText)));
                stringBuilder.Append("</Message>");

                XDocument xDocument = XDocument.Load(new StringReader(stringBuilder.ToString()));
                
                List<ReportTask> data = (from reportTask in xDocument.Root.Elements("ReportsTasks").Elements("ReportTask")
                                         select new ReportTask()
                                         {
                                             id = Convert.ToInt32(reportTask.Element("Id").Value.Trim()),
                                             userId = reportTask.Element("UserId").Value.Trim(),
                                             type = reportTask.Element("Type").Value.Trim(),
                                             active = reportTask.Element("Active").Value.Trim(),
                                             repeatMon = reportTask.Element("RepeatMon").Value.Trim(),
                                             repeatTue = reportTask.Element("RepeatTue").Value.Trim(),
                                             repeatWed = reportTask.Element("RepeatWed").Value.Trim(),
                                             repeatThu = reportTask.Element("RepeatThu").Value.Trim(),
                                             repeatFri = reportTask.Element("RepeatFri").Value.Trim(),
                                             repeatSat = reportTask.Element("RepeatSat").Value.Trim(),
                                             repeatSun = reportTask.Element("RepeatSun").Value.Trim(),
                                             repeatMonth = reportTask.Element("RepeatMonth").Value.Trim(),
                                             repeatYear = reportTask.Element("RepeatYear").Value.Trim(),
                                             name = reportTask.Element("Name").Value.Trim(),
                                             description = reportTask.Element("Description").Value.Trim(),
                                             timeCreated = reportTask.Element("TimeCreated").Value.Trim(),
                                             runTime = reportTask.Element("RunTime").Value.Trim(),
                                             fromTime = reportTask.Element("FromTime").Value.Trim(),
                                             toTime = reportTask.Element("ToTime").Value.Trim(),
                                             interval = reportTask.Element("Interval").Value.Trim(),
                                             format = reportTask.Element("Format").Value.Trim(),
                                             userForId = reportTask.Element("UserForId").Value.Trim(),
                                             criticalType = reportTask.Element("CriticalType").Value.Trim(),
                                             atmsIds = new List<string>(from atmId in reportTask.Element("AtmsIds").Elements("Id")
                                                                        select atmId.Value.Trim()),
                                             groupsIds = new List<string>(from atmId in reportTask.Element("GroupsIds").Elements("Id")
                                                                          select atmId.Value.Trim()),
                                             usersIds = new List<string>(from userId in reportTask.Element("UsersIds").Elements("Id")
                                                                          select userId.Value.Trim())
                                         }).ToList();

                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i].type == "IncidentsHistoryCurrent")
                        data[i].fromTime = data[i].toTime;
                }

                this.info.data = data.ToList();
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportTaskGet() exeption: " + exception.Message);
                this.info.isError = 1;
            }
        }
    }
}