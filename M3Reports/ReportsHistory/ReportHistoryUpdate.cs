using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Web;

namespace M3Reports
{
    using M3IPClient;

    public class ReportHistoryUpdate
    {
        public struct ReportHistoryInfo : IDataInfo
        {
            public int isError { get; set; }
            public string status;
        }

        public
            ReportHistoryInfo reportHistoryInfo = new ReportHistoryInfo();

        public void ParseMessage(XmlNode messageNode)
        {
            this.reportHistoryInfo.isError = 0;

            try
            {
                this.reportHistoryInfo.status = messageNode.SelectSingleNode("Request/Status").InnerText.Trim();
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportHistorySet() exeption: " + exception.Message);
                this.reportHistoryInfo.isError = 1;
            }
        }
    }
}