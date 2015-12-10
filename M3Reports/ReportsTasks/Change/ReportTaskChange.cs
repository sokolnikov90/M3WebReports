using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Web;

namespace M3Reports
{
    public class ReportTaskChange
    {
        public struct Info
        {
            public int isError;
            public string status;
        }

        public
            Info info = new Info();

        public void ParseMessage(XmlNode messageNode)
        {
            this.info.isError = 0;

            try
            {
                this.info.status = messageNode.SelectSingleNode("Request/Status").InnerText.Trim();
            }
            catch (Exception exception)
            {
                M3Utils.Log.Instance.Info("ReportTaskChange() exeption: " + exception.Message);
                this.info.isError = 1;
            }
        }
    }
}