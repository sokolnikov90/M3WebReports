namespace M3Reports
{
    using System.Globalization;

    /// <summary>
    /// Базовый класс для построение отчёта
    /// </summary>
    public abstract class ReportBuilder
    {
        // Данные, пришедшие в JSON
        internal ReportInfo Info { get; set; }

        // Данные, полученные от WP-сервиса
        internal ReportData Data
        {
            get
            {
                return this.data;
            }
        }

        private ReportData data = new ReportData();

        internal abstract void MakeAnExcel();

        public static CultureInfo GetCultureInfo(string languageCode)
        {
            CultureInfo cultureInfo;

            switch (languageCode.ToUpper())
            {
                case "EN":
                    cultureInfo = CultureInfo.CreateSpecificCulture("en-US");
                    break;
                case "RU":
                default:
                    cultureInfo = CultureInfo.CreateSpecificCulture("ru-RU");
                    break;
            }

            return cultureInfo;
        }
    }
}