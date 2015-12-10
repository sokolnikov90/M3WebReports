using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using System.Xml;

namespace M3WebService
{
    public class ReportOutOfServiceFacade : IPClientDecorator
    {
        AtmInfoGet atmInfoGet = new AtmInfoGet();      
        AtmGroupsGet atmGroupget = new AtmGroupsGet();
        IDictionaryGet dictionaryGet;

        Report report = new Report();
        ReportOutOfService reportOutOfService = new ReportOutOfService();
        ReportHistorySet reportHistorySet = new ReportHistorySet();

        IncidentGet queryIncident = new IncidentGet();
        IncidentsGet incidentsGet = new IncidentsGet();

        DictionariesGet dictionariesGet = new DictionariesGet();

        string status;

        public ReportOutOfServiceFacade(string ip, int port, string login, string password, string status, string incidentsGetJSON)
        {
            this.status = status;

            connection.ReadEvent += new IPClient.ReadDelegate(IPRead);

            autorization.info.isError = Connect(ip, port, login, password);

            report = new JavaScriptSerializer().Deserialize<Report>(incidentsGetJSON);
        }

        public void DoStuff()
        {
            if (autorization.info.isError == 0)
            {
                connection.Write(queries.ReportHistorySet(report));

                ewh.Reset();
                ewh.WaitOne();

                if (reportHistorySet.reportHistoryInfo.status == "SUCCESS")
                {
                    reportOutOfService.Init(report);
                    connection.Write(queries.QueryDictionaryGet("Statuses"));

                    ewh.Reset();
                    ewh.WaitOne();

                    if (dictionaryGet != null && dictionaryGet.GetType().ToString() == "M3WebService.DictionaryGetStatuses")
                        dictionariesGet.info.statuses = (DictionaryGetStatuses.Info)dictionaryGet.Info;

                    connection.Write(queries.QueryDictionaryGet("Types"));

                    ewh.Reset();
                    ewh.WaitOne();

                    if (dictionaryGet != null && dictionaryGet.GetType().ToString() == "M3WebService.DictionaryGetTypes")
                        dictionariesGet.info.types = (DictionaryGetTypes.Info)dictionaryGet.Info;

                    connection.Write(queries.QueryDictionaryGet("Users"));

                    ewh.Reset();
                    ewh.WaitOne();

                    if (dictionaryGet != null && dictionaryGet.GetType().ToString() == "M3WebService.DictionaryGetUsers")
                    {
                        dictionariesGet.info.users = (DictionaryGetUsers.Info)dictionaryGet.Info;

                        dictionariesGet.info.users.userId = autorization.info.userId;
                        dictionariesGet.info.users.role = autorization.info.role;
                    }

                    connection.Write(queries.QueryDictionaryGet("UserRoles"));

                    ewh.Reset();
                    ewh.WaitOne();

                    if (dictionaryGet != null && dictionaryGet.GetType().ToString() == "M3WebService.DictionaryGetUserRoles")
                        dictionariesGet.info.userRoles = (DictionaryGetUserRoles.Info)dictionaryGet.Info;

                    connection.Write(queries.QueryAtmInfo(report.atmsId));

                    ewh.Reset();
                    ewh.WaitOne();

                    queryIncident.from = report.from;
                    queryIncident.to = report.to;
                    queryIncident.statusIds = String.Join(", ", (from item in dictionariesGet.info.statuses.data where ((Convert.ToInt32(item.isClosed) == 0) || (Convert.ToInt32(item.isClosed) == 1)) select item.id).ToArray());
                    queryIncident.atmIds = report.atmsId;
                    queryIncident.typeIds = String.Join(", ", (from item in dictionariesGet.info.types.data select item.id.ToString()).ToArray());
                    queryIncident.userRoleIds = (from item in dictionariesGet.info.userRoles.data where item.appType == "M3Web" select item.id).ToList();

                    connection.Write(queries.QueryIncidentsGet(queryIncident));

                    ewh.Reset();
                    ewh.WaitOne();    
                    
                    connection.Write(queries.QueryDictionaryGet("DevicesTypes"));
                    ewh.Reset();
                    ewh.WaitOne();

                    if (dictionaryGet != null && dictionaryGet.GetType().ToString() == "M3WebService.DictionaryGetDevicesTypes")
                        dictionariesGet.info.devicesTypes = (DictionaryGetDevicesTypes.Info)dictionaryGet.Info;

                    connection.Write(queries.QueryAtmGroups(new List<string>() {report.userId}));
                    ewh.Reset();
                    ewh.WaitOne();
                    reportOutOfService.groups = SearchGroupNames(report.atmsGroupsId,atmGroupget.info);
                    reportOutOfService.groupTree = atmGroupget.info;   

                    connection.Write(queries.QueryAtmInfo(report.atmsId));
                    ewh.Reset();
                    ewh.WaitOne();

                    reportOutOfService.ActualIncidents = incidentsGet.incidentInfo.data;
                    reportOutOfService.AtmInfoLst = atmInfoGet.info.data;
                    reportOutOfService.dictionariesInfo = dictionariesGet.info;
                   // reportOutOfService.deviceTypes = (DictionaryGetDevicesTypes)dictionaryGet;
                    reportOutOfService.CreateFailuresDict();
                    switch (report.format)
                    {
                        case "xlsx":
                            reportOutOfService.MakeAnExcel();
                            break;
                    }

                    if (reportHistorySet.reportHistoryInfo.isError != 0)
                        reportOutOfService.reportMonitoringInfo.isError = 1;

                    connection.Write(queries.ReportHistoryUpdate(reportHistorySet, reportOutOfService.reportMonitoringInfo.isError, reportOutOfService.reportMonitoringInfo.path));

                    ewh.Reset();
                    ewh.WaitOne();
                }
            }
            else
            {
                reportOutOfService.reportMonitoringInfo.isError = 1;
            }

            connection.Disconnect();
        }

        private void IPRead(string message, bool complit)
        {
            M3WebService.logger.Info("<-IPRead(): " + message);

            XmlNode messageNode;
            XmlDocument xmlDocument = new XmlDocument();

            try
            {
                xmlDocument.LoadXml(message);

                messageNode = xmlDocument.SelectSingleNode("Message");

                requestName = messageNode.SelectSingleNode("Request/./@name").InnerText;

                switch (requestName)
                {
                    case "CAdminError":
                        autorization.ParseMessage(messageNode);
                        break;
                    case "CWebReportHistorySetInfo":
                        reportHistorySet.ParseMessage(messageNode);
                        break;
                    case "CWebIncidentsGetInfo":
                        incidentsGet.ParseMessage(messageNode);
                        break;
                    case "CAtmsInfo":
                        atmInfoGet.ParseMessage(messageNode);
                        break;
                    case "CWebDictionaryGetInfo":
                        switch (messageNode.SelectSingleNode("Request").Attributes["type"].InnerText)
                        {
                            case "Statuses":
                                dictionaryGet = new DictionaryGetStatuses();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "Types":
                                dictionaryGet = new DictionaryGetTypes();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "Users":
                                dictionaryGet = new DictionaryGetUsers();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "UserRoles":
                                dictionaryGet = new DictionaryGetUserRoles();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "DevicesTypes":
                                dictionaryGet = new DictionaryGetDevicesTypes();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "IncidentsRules":
                                dictionaryGet = new DictionaryGetIncidentsRules();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                            case "ResponsibleFor":
                                dictionaryGet = new DictionaryGetResponsibleFor();
                                dictionaryGet.ParseMessage(messageNode);
                                break;
                        }
                        break;

                    case "CUsersGroupsAtmsInfo":
                        atmGroupget.ParseMessage(messageNode);
                        break;

                   
                }
            }
            catch (Exception exception)
            {
                M3WebService.logger.Info("IPRead() exeption: " + exception.Message);

                autorization.info.isError = 1;
            }
            finally
            {
                ewh.Set();
            }
        }

        public string Response()
        {
            return new JavaScriptSerializer().Serialize(reportOutOfService.reportMonitoringInfo);
        }

        private Dictionary<string, AtmGroupsGet.AtmGroup> SearchGroupNames(List<string> GroupIDs, AtmGroupsGet.Info info)
        {
            Dictionary<string, AtmGroupsGet.AtmGroup> groups = new Dictionary<string, AtmGroupsGet.AtmGroup>();
            foreach(string id in GroupIDs)
            {
                groups.Add(id,recursiveSearch(id, info.usersGroup[0].atmGroups));
            }
 
            return groups;
        }

        private AtmGroupsGet.AtmGroup recursiveSearch(string id, List<AtmGroupsGet.AtmGroup> info)
        {
            var result = new AtmGroupsGet.AtmGroup();
            result.name = string.Empty;
            foreach (var group in info)
            {
                if (group.id.ToString() == id)
                    return group;
                if (!group.atmGroups.IsNullOrEmpty())
                {
                    result = recursiveSearch(id, group.atmGroups);
                    if(  result.name != string.Empty)
                      return result;
                    
                }
            }
            return result;
        }


    }
}