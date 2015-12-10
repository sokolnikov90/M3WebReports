using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace M3Reports
{
    public class ReportHistory
    {
        public string id { get; set; }
        public string taskId { get; set; }
        public string type { get; set; }
        public string userId { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string status { get; set; }
        public string runTime { get; set; }
        public string fromTime { get; set; }
        public string toTime { get; set; }
        public string path { get; set; }
        public string userForId { get; set; }
        public List<string> atmsIds { get; set; }
        public List<string> groupsIds { get; set; }
    }
}