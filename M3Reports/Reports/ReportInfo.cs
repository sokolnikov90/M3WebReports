namespace M3Reports
{
    using System.Collections.Generic;
    using M3Incidents;
    using M3Dictionaries;

    using M3IPClient;

    /// <summary>
    /// Класс содержит данные, пришедшие в JSON
    /// </summary>
    public class ReportInfo : IDataInfo
    {
        public int isError { get; set; }
        public string taskId { get; set; }
        public string userId { get; set; }
        public string type { get; set; }
        public string runTime { get; set; }
        public string from { get; set; }
        public string to { get; set; }
        public string path { get; set; }
        public string format { get; set; }
        public string criticalType { get; set; }
        public string languageCode { get; set; }

        public List<string> atmsId { get; set; }
        public List<string> atmsGroupsId { get; set; }
        public List<Incident> incidents { get; set; }

        public Response Response
        {
            get
            {
                return new Response
                {
                    isError = this.isError,
                    from = this.from,
                    to = this.to,
                    path = this.path
                };
            }
        }
    }

    public struct Response : IDataInfo
    {
        public int isError { get; set; }
        public string from { get; set; }
        public string to { get; set; }
        public string path { get; set; }
    }
}
