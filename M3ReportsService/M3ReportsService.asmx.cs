using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

using M3Reports;
using M3IPClient;

namespace M3ReportsService
{
    public class M3ReportsService : AsyncWebService
    {
        [WebMethod(Description = "Создание задания для составления отчета.")]
        public IAsyncResult BeginReportTaskCreate(string ip, int port, string login, string password,  string reportTaskCreateJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportTaskCreateFacade), ip, port, login, password, reportTaskCreateJSON};

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Создание задания для составления отчета.")]
        public string EndReportTaskCreate(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Получение заданий отчетов.")]
        public IAsyncResult BeginReportTasksGet(string ip, int port, string login, string password, string reportTaskGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportTasksGetFacade), ip, port, login, password, reportTaskGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Получение заданий отчетов.")]
        public string EndReportTasksGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Изменение существующего задания отчета")]
        public IAsyncResult BeginReportTaskChange(string ip, int port, string login, string password, string reportTaskChangeJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportTaskChangeFacade), ip, port, login, password, reportTaskChangeJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Изменение существующего задания отчета")]
        public string EndReportTaskChange(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Удаление задания отчета")]
        public IAsyncResult BeginReportTaskDelete(string ip, int port, string login, string password, List<string> taskIds, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportTaskDeleteFacade), ip, port, login, password, taskIds };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Удаление задания отчета")]
        public string EndReportTaskDelete(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Получение списка выполненных заданий")]
        public IAsyncResult BeginReportHistoryGet(string ip, int port, string login, string password, string reportHistoryGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportHistoryGetFacade), ip, port, login, password, reportHistoryGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Получение списка выполненных заданий")]
        public string EndReportHistoryGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по сообщениям.")]
        public IAsyncResult BeginReportMessagesGet(string ip, int port, string login, string password, string messagesGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportMessagesGetFacade), ip, port, login, password, messagesGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по сообщениям.")]
        public string EndReportMessagesGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по транзакциям.")]
        public IAsyncResult BeginReportTransactionsGet(string ip, int port, string login, string password, string transactionsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportTransactionsGetFacade), ip, port, login, password, transactionsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по транзакциям.")]
        public string EndReportTransactionsGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет о доступности.")]
        public IAsyncResult BeginReportAvailabilitiesGet(string ip, int port, string login, string password, string availsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportAvailabilitiesGetFacade), ip, port, login, password, availsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет о доступности.")]
        public string EndReportAvailabilitiesGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет об инцидентах. По типу инцидента.")]
        public IAsyncResult BeginReportIncidentsByTypesGet(string ip, int port, string login, string password, string incidentsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportIncidentsByTypesGetFacade), ip, port, login, password, incidentsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет об инцидентах. По типу инцидента.")]
        public string EndReportIncidentsByTypesGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет об инцидентах. По устройству.")]
        public IAsyncResult BeginReportIncidentsByDevicesGet(string ip, int port, string login, string password, string incidentsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportIncidentsByDevicesGetFacade), ip, port, login, password, incidentsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет об инцидентах. По устройству.")]
        public string EndReportIncidentsByDevicesGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет службы мониторинга.")]
        public IAsyncResult BeginReportMonitoringGet(string ip, int port, string login, string password, string incidentsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportMonitoringFacade), ip, port, login, password, incidentsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет службы мониторинга.")]
        public string EndReportMonitoringGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет о простаивающих банкоматах.")]
        public IAsyncResult BeginReportOutOfServiceGet(string ip, int port, string login, string password, string incidentsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportOutOfServiceFacade), ip, port, login, password, incidentsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет о простаивающих банкоматах.")]
        public string EndReportOutOfServiceGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет об остатках наличности.")]
        public IAsyncResult BeginReportCountsGet(string ip, int port, string login, string password, string countsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportCountsGetFacade), ip, port, login, password, countsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет об остатках наличности.")]
        public string EndReportCountsGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по  событиям.")]
        public IAsyncResult BeginReportEventsGet(string ip, int port, string login, string password, string eventsGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportEventsGetFacade), ip, port, login, password, eventsGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по  событиям.")]
        public string EndReportEventsGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по  действиям пользователя.")]
        public IAsyncResult BeginReportUsersGet(string ip, int port, string login, string password, string usersGetJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportUsersGetFacade), ip, port, login, password, usersGetJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по  действиям пользователя.")]
        public string EndReportUsersGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по сводной таблице инцидентов.")]
        public IAsyncResult BeginReportIncidentsPivotTableGet(string ip, int port, string login, string password, string reportRunJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportIncidentsPivotTableFacade), ip, port, login, password, reportRunJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по сводной таблице инцидентов.")]
        public string EndReportIncidentsPivotTableGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по открытым инцидентам.")]
        public IAsyncResult BeginReportIncidentsOpenGet(string ip, int port, string login, string password, string reportRunJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportIncidentsOpenFacade), ip, port, login, password, reportRunJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по открытым инцидентам.")]
        public string EndReportIncidentsOpenGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Сводный отчет по инцидентам.")]
        public IAsyncResult BeginReportSummaryIncidentsOpenGet(string ip, int port, string login, string password, string reportRunJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportSummaryIncidentsOpenFacade), ip, port, login, password, reportRunJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Сводный отчет по инцидентам.")]
        public string EndReportSummaryIncidentsOpenGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет о зарегистрированных банкоматах.")]
        public IAsyncResult BeginReportAllAtmsGet(string ip, int port, string login, string password, string reportRunJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportAllAtmsFacade), ip, port, login, password, reportRunJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет о зарегистрированных банкоматах.")]
        public string EndReportAllAtmsGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }

        [WebMethod(Description = "Отчет по прогнозу инкассации.")]
        public IAsyncResult BeginReportCollectionForecastGet(string ip, int port, string login, string password, string reportRunJSON, AsyncCallback cb, object state)
        {
            object[] taskState = { typeof(ReportCollectionForecastFacade), ip, port, login, password, reportRunJSON };

            return StubCallBack.BeginInvoke(taskState, cb, null);
        }

        [WebMethod(Description = "Отчет по прогнозу инкассации.")]
        public string EndReportCollectionForecastGet(IAsyncResult call)
        {
            return StubCallBack.EndInvoke(call);
        }
    }
}