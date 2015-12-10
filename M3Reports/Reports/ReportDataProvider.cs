namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Xml;
    using System.Xml.Linq;

    using M3Atms;

    using M3Dictionaries;

    using M3IPClient;

    /// <summary>
    /// Базовый класс получения данных для построения отчётов.
    /// В наследнике нужно переопределить:
    ///     SendDataQueries() - отправляет запросы к WP
    ///     ParseMessage() - парсит запросы от WP
    /// </summary>
    public abstract class ReportDataProvider : M3UserSession, IAsyncRequestFacade
    {
        /// <summary>
        /// Парсинг XML-файлов, содержащих колонки Excel отчётов
        /// </summary>
        /// <param name="xmlPath"></param>
        /// <returns></returns>
        public static List<ReportColumns> ParseXML(string xmlPath)
        {
            XDocument xmlDocument = XDocument.Load(AppDomain.CurrentDomain.BaseDirectory + xmlPath);

            var columns =
                xmlDocument.Root.Elements("Column")
                    .Select(
                        column =>
                        new ReportColumns()
                        {
                            localtion = column.Element("Location").Value,
                            name = column.Element("Name").Value,
                            title = ReportsSource.ResourceManager.GetString(column.Element("Name").Value),
                            width = Convert.ToDouble(column.Element("Width").Value)
                        });
            return columns.ToList();
        }

        protected ReportBuilder report;

        protected ReportDataProvider(string ip, int port, string login, string password, string getJSON)
            : base(ip, port, login, password, getJSON)
        {
            this.connection.ReadEvent += this.IPRead;
        }

        /// <summary>
        /// Template Method. Подключение к WP, получение данных и построение отчёта.
        /// Обязательно переопределить для FrontendReports (т.к. не требуется запись в БД о построении отчёта)
        /// </summary>
        public virtual void DoStuff()
        {
            if (ReportData.ExistInfo(this.signin.info))
            {
                this.connection.Write(Queries.ReportHistorySet(this.report), this.ewh);

                if (ReportData.ExistInfo(this.report.Data.ReportHistorySet.reportHistoryInfo))
                {
                    if (String.IsNullOrEmpty(this.report.Info.languageCode)) this.report.Info.languageCode = "RU";

                    try
                    {
                        this.SendDataQueries();
                    }
                    catch (Exception exp)
                    {
                        M3Utils.Log.Instance.Info(this + ".SendDataQueries(...) exception:");
                        M3Utils.Log.Instance.Info(exp.Message);
                        M3Utils.Log.Instance.Info(exp.Source);
                        M3Utils.Log.Instance.Info(exp.StackTrace);
                    }

                    Thread.CurrentThread.CurrentCulture = Thread.CurrentThread.CurrentUICulture =  ReportBuilder.GetCultureInfo(this.report.Info.languageCode);

                    switch (this.report.Info.format)
                    {
                        case "xlsx":
                            this.report.MakeAnExcel();
                            break;
                    }
                }
                else
                {
                    this.report.Info.isError = 1;                    
                }
            }
            else
            {
                this.report.Info.isError = 1;
            }

            this.connection.Write(Queries.ReportHistoryUpdate(this.report.Data.ReportHistorySet, this.report.Info.isError, this.report.Info.path), this.ewh);

            if (!ReportData.ExistInfo(this.report.Data.ReportHistoryUpdate.reportHistoryInfo))
            {
                this.report.Info.isError = 1;                
            }

            this.connection.ReadEvent -= this.IPRead;
            this.connection.Disconnect();
        }

        /// <summary>
        /// Чтение сообщений от WP.
        /// </summary>
        private void IPRead(string message, bool complit)
        {
            M3Utils.Log.Instance.Info(this + ".IPRead(): " + message);

            XmlDocument xmlDocument = new XmlDocument();

            try
            {
                xmlDocument.LoadXml(message);

                var messageNode = xmlDocument.SelectSingleNode("Message");

                this.requestName = messageNode.SelectSingleNode("Request/./@name").InnerText;

                this.ParseMessage(messageNode);

            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                M3Utils.Log.Instance.Info(this + ".IPRead(...) exception:");
                M3Utils.Log.Instance.Info(exp.Message);
                M3Utils.Log.Instance.Info(exp.Source);
                M3Utils.Log.Instance.Info(exp.StackTrace);
            }
            finally
            {
                this.ewh.Set();
            }
        }

        /// <summary>
        /// Отправка ответа клиенту
        /// </summary>
        /// <returns></returns>
        public virtual string Response()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this.report.Info.Response);
        }

        /// <summary>
        /// Отправляет запросы на получение данных от WP
        /// Переопределить в наследнике для запроса данных, необходимых для конкретного отчёта
        /// </summary>
        protected virtual void SendDataQueries()
        {

        }

        /// <summary>
        /// Парсинг сообщений от WP.
        /// Переопределить в наследнике для кастомного парсинга сообщений
        /// или для парсинга "редких" запросов
        /// </summary>
        /// <param name="messageNode"></param>
        protected virtual void ParseMessage(XmlNode messageNode)
        {
            switch (this.requestName)
            {
                case "CWebDictionaryGetInfo":
                    var dictionaries = this.report.Data.DictionariesGet = this.report.Data.DictionariesGet ?? new GetAll();
                    switch (messageNode.SelectSingleNode("Request").Attributes["type"].InnerText)
                    {
                        case "Statuses":
                            dictionaries.statuses = new GetStatuses();
                            dictionaries.statuses.ParseMessage(messageNode);
                            break;
                        case "Types":
                            dictionaries.types = new GetTypes();
                            dictionaries.types.ParseMessage(messageNode);
                            break;
                        case "Users":
                            dictionaries.users = new GetUsers();
                            dictionaries.users.ParseMessage(messageNode);
                            dictionaries.info.users.userId = this.signin.info.userId;
                            dictionaries.info.users.roleId = this.signin.info.roleId;
                            break;
                        case "UserRoles":
                            dictionaries.userRoles = new GetUserRoles();
                            dictionaries.userRoles.ParseMessage(messageNode);
                            break;
                        case "DevicesTypes":
                            dictionaries.devicesTypes = new GetDevicesTypes();
                            dictionaries.devicesTypes.ParseMessage(messageNode);
                            break;
                        case "IncidentsRules":
                            dictionaries.incidentsRules = new GetIncidentsRules();
                            dictionaries.incidentsRules.ParseForReport(messageNode);
                            break;
                        case "ResponsibleFor":
                            dictionaries.responsibleFor = new GetResponsibleFor();
                            dictionaries.responsibleFor.ParseMessage(messageNode);
                            break;
                    }
                    break;
                case "CWebReportHistorySetInfo":
                    this.report.Data.ReportHistorySet = new ReportHistorySet();
                    this.report.Data.ReportHistorySet.ParseMessage(messageNode);
                    break;
                case "CWebReportHistoryUpdateInfo":
                    this.report.Data.ReportHistoryUpdate = new ReportHistoryUpdate();
                    this.report.Data.ReportHistoryUpdate.ParseMessage(messageNode);
                    break;
                case "CAtmsInfo":
                    this.report.Data.AtmInfoGet = new InfoGet();
                    this.report.Data.AtmInfoGet.ParseMessage(messageNode);
                    break;
                case "CUsersGroupsAtmsInfo":
                    this.report.Data.AtmGroupsGet = new GroupsGet();
                    this.report.Data.AtmGroupsGet.ParseMessage(messageNode);
                    break;
                case "CWebIncidentsGetInfo":
                    this.report.Data.IncidentsGet = new M3Incidents.GetAll();
                    this.report.Data.IncidentsGet.ParseMessage(messageNode);
                    break;
            }
        }
    }
}