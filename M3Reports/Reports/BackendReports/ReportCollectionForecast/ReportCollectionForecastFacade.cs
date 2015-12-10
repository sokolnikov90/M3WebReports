namespace M3Reports
{
    using System;
    using System.Xml;

    using M3Atms;

    using M3Dictionaries;

    public class ReportCollectionForecastFacade : ReportDataProvider
    {
        public ReportCollectionForecastFacade(string ip, int port, string login, string password, string getJSON)
            : base(ip, port, login, password, getJSON)
        {
            this.report = new ReportCollectionForecast();
            this.report.Info = Newtonsoft.Json.JsonConvert.DeserializeObject<ReportInfo>(getJSON);
        }

        protected override void SendDataQueries()
        {
            this.connection.Write(M3Atms.Queries.QueryAtmInfo(this.report.Info.atmsId), this.ewh);

            this.connection.Write(M3Atms.Queries.QueryAtmWithdrawHystory(this.report.Info.atmsId), this.ewh);

            DateTime today = DateTime.Today;  //new DateTime(2015, 07, 25);
            DateTime toDateTime = today.AddDays(13);

            this.report.Info.from = today.ToString("yyyy-MM-dd");
            this.report.Info.to = toDateTime.ToString("yyyy-MM-dd");
        }

        protected override void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CAtmsWithdrawHistoryInfo":
                    this.report.Data.WithdrawHistoryGet = new WithdrawHistoryGet();
                    this.report.Data.WithdrawHistoryGet.ParseMessage(messageNode);
                    break;
                default:
                    base.ParseMessage(messageNode);
                    break;
            }
        }
    }
}
