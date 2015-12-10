using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;

namespace M3Reports
{
    public class ReportHistoryGet
    {
        public struct Info
        {
            public int isError;
            public List<ReportHistory> data;
        }

        public
            Info info = new Info();

        public void ParseMessage(XmlNode messageNode)
        {
            this.info.isError = 0;

            try
            {
                XDocument xDocument = XDocument.Load(XmlReader.Create(new StringReader(messageNode.InnerXml)));

                IEnumerable<ReportHistory> data = from reportTask in xDocument.Root.Elements("Tasks").Elements("Task")
                                                  select new ReportHistory()
                                                  {
                                                      id = reportTask.Element("Id").Value.Trim(),
                                                      taskId = reportTask.Element("TaskId").Value.Trim(),
                                                      type = reportTask.Element("Type").Value.Trim(),
                                                      userId = reportTask.Element("UserId").Value.Trim(),
                                                      name = reportTask.Element("Name").Value.Trim(),
                                                      description = reportTask.Element("Description").Value.Trim(),
                                                      status = reportTask.Element("Status").Value.Trim(),
                                                      runTime = reportTask.Element("RunTime").Value.Trim(),
                                                      fromTime = reportTask.Element("FromTime").Value.Trim(),
                                                      toTime = reportTask.Element("ToTime").Value.Trim(),
                                                      path = reportTask.Element("Path").Value.Trim(),
                                                      userForId = reportTask.Element("UserForId").Value.Trim(),
                                                      atmsIds = new List<string>(from atmId in reportTask.Elements("AtmsIds").Elements("Id")
                                                                                 select atmId.Value.Trim()),
                                                      groupsIds = new List<string>(from groupId in reportTask.Elements("GroupsIds").Elements("Id")
                                                                                   select groupId.Value.Trim())
                                                  };

                this.info.data = data.ToList();
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportHistoryGet() exeption: " + exception.Message);
                this.info.isError = 1;
            }
        }
    }
}