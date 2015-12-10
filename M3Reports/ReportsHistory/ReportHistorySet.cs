using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Web;

namespace M3Reports
{
    using M3IPClient;

    public class ReportHistorySet
    {
        public struct ReportHistoryInfo : IDataInfo
        {
            public int isError { get; set; }
            public int id;
            public string tableName;
        }

        public
            ReportHistoryInfo reportHistoryInfo = new ReportHistoryInfo();

        public void ParseMessage(XmlNode messageNode)
        {
            this.reportHistoryInfo.isError = 0;

            try
            {
                var status = messageNode.SelectSingleNode("Request/Status").InnerText.Trim();

                if (status == "SUCCESS")
                {
                    this.reportHistoryInfo.id = Convert.ToInt32(messageNode.SelectSingleNode("Request/Id").InnerText.Trim());
                    this.reportHistoryInfo.tableName = messageNode.SelectSingleNode("Request/TableName").InnerText.Trim();
                }
                else
                {
                    this.reportHistoryInfo.isError = 1;
                }
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportHistorySet() exeption: " + exception.Message);
                this.reportHistoryInfo.isError = 1;
            }
        }
    }
}