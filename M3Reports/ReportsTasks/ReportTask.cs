using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace M3Reports
{
    public class ReportTask
    {
        public int id { get; set; }
        public string type { get; set; }
        public string userId { get; set; }
        public string active { get; set; }
        public string repeatMon { get; set; }
        public string repeatTue { get; set; }
        public string repeatWed { get; set; }
        public string repeatThu { get; set; }
        public string repeatFri { get; set; }
        public string repeatSat { get; set; }
        public string repeatSun { get; set; }
        public string repeatMonth { get; set; }
        public string repeatYear { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string format { get; set; }
        public string fromTime { get; set; }
        public string toTime { get; set; }
        public string runTime { get; set; }
        public string interval { get; set; }
        public string timeCreated { get; set; }
        public string userForId { get; set; }
        public string criticalType { get; set; }
        public string languageCode { get; set; }

        public List<string> ids { get; set; }
        
        public List<string> atmsIds { get; set; }
        public List<string> groupsIds { get; set; }
        public List<string> usersIds { get; set; }

        public List<ReportTaskItem> items { get; set; }
    }

    public class ReportTaskItem
    {
        public string name { get; set; }
        public string type { get; set; }
        public string value { get; set; }
    }
}