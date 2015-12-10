using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using M3Atms;
using M3Incidents;
using M3Dictionaries;
using System.Text;

namespace M3Reports
{
    public static class Queries
    {
        public static string QueryMessageReport(IEnumerable<string> atmIds, string startDate, string endDate)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CMessageHistoryGet\">");
            stringBuilder.AppendFormat("<From>{0}</From>", startDate);
            stringBuilder.AppendFormat("<To>{0}</To>", endDate);
            stringBuilder.Append("<Atms>");
            
            foreach (var atm in atmIds)
                stringBuilder.AppendFormat("<Atm>{0}</Atm>", atm);

            stringBuilder.Append("</Atms>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }


        public static string ReportTaskCreate(ReportTask reportTask)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportTaskCreate\">");
            stringBuilder.AppendFormat("<UserId>{0}</UserId>", reportTask.userId);
            stringBuilder.AppendFormat("<Type>{0}</Type>", reportTask.type);
            stringBuilder.AppendFormat("<Name>{0}</Name>", reportTask.name);
            stringBuilder.AppendFormat("<Description>{0}</Description>", reportTask.description);
            stringBuilder.AppendFormat("<TimeCreated>{0}</TimeCreated>", reportTask.timeCreated);
            stringBuilder.AppendFormat("<Active>{0}</Active>", reportTask.active);
            stringBuilder.AppendFormat("<RepeatMon>{0}</RepeatMon>", reportTask.repeatMon);
            stringBuilder.AppendFormat("<RepeatTue>{0}</RepeatTue>", reportTask.repeatTue);
            stringBuilder.AppendFormat("<RepeatWed>{0}</RepeatWed>", reportTask.repeatWed);
            stringBuilder.AppendFormat("<RepeatThu>{0}</RepeatThu>", reportTask.repeatThu);
            stringBuilder.AppendFormat("<RepeatFri>{0}</RepeatFri>", reportTask.repeatFri);
            stringBuilder.AppendFormat("<RepeatSat>{0}</RepeatSat>", reportTask.repeatSat);
            stringBuilder.AppendFormat("<RepeatSun>{0}</RepeatSun>", reportTask.repeatSun);
            stringBuilder.AppendFormat("<RepeatMonth>{0}</RepeatMonth>", reportTask.repeatMon);
            stringBuilder.AppendFormat("<RepeatYear>{0}</RepeatYear>", reportTask.repeatYear);
            stringBuilder.AppendFormat("<RunTime>{0}</RunTime>", reportTask.runTime);
            stringBuilder.AppendFormat("<FromTime>{0}</FromTime>", reportTask.fromTime);
            stringBuilder.AppendFormat("<ToTime>{0}</ToTime>", reportTask.toTime);
            stringBuilder.AppendFormat("<Interval>{0}</Interval>", reportTask.interval);
            stringBuilder.AppendFormat("<Format>{0}</Format>", reportTask.format);
            stringBuilder.AppendFormat("<UserForId>{0}</UserForId>", reportTask.userForId);
            stringBuilder.AppendFormat("<CriticalType>{0}</CriticalType>", reportTask.criticalType);
            stringBuilder.AppendFormat("<LanguageCode>{0}</LanguageCode>", reportTask.languageCode);
            stringBuilder.Append("<AtmsIds>");
            for (int i = 0; i < reportTask.atmsIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.atmsIds[i]);
            }
            stringBuilder.Append("</AtmsIds>");
            stringBuilder.Append("<GroupsIds>");
            for (int i = 0; i < reportTask.groupsIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.groupsIds[i]);
            }
            stringBuilder.Append("</GroupsIds>");
            stringBuilder.Append("<UsersIds>");
            for (int i = 0; i < reportTask.usersIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.usersIds[i]);
            }
            stringBuilder.Append("</UsersIds>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportTaskGet(ReportTask reportTask)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportTaskGet\">");
            stringBuilder.AppendFormat("<Active>{0}</Active>", reportTask.active);
            stringBuilder.Append("<ZipResponse>true</ZipResponse>");
            stringBuilder.Append("<UsersIds>");
            for (int i = 0; i < reportTask.usersIds.Count; i++)
            {
                stringBuilder.AppendFormat("<UserId>{0}</UserId>", reportTask.usersIds[i]);
            }
            stringBuilder.Append("</UsersIds>");
            stringBuilder.Append("<Ids>");
            for (int i = 0; i < reportTask.ids.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.ids[i]);
            }
            stringBuilder.Append("</Ids>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportTaskChange(ReportTask reportTask)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportTaskChange\">");
            stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.id);
            stringBuilder.AppendFormat("<UserId>{0}</UserId>", reportTask.userId);
            stringBuilder.Append("<Items>");
            for (int i = 0; i < reportTask.items.Count; i++)
            {
                stringBuilder.Append("<Item>");
                stringBuilder.AppendFormat("<Name>{0}</Name>", reportTask.items[i].name);
                stringBuilder.AppendFormat("<Type>{0}</Type>", reportTask.items[i].type);
                stringBuilder.AppendFormat("<Value>{0}</Value>", reportTask.items[i].value);
                stringBuilder.Append("</Item>");
            }
            stringBuilder.Append("</Items>");

            stringBuilder.Append("<AtmsIds>");
            for (int i = 0; i < reportTask.atmsIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.atmsIds[i]);
            }
            stringBuilder.Append("</AtmsIds>");

            stringBuilder.Append("<GroupsIds>");
            for (int i = 0; i < reportTask.groupsIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.groupsIds[i]);
            }
            stringBuilder.Append("</GroupsIds>");

            stringBuilder.Append("<UsersIds>");
            for (int i = 0; i < reportTask.usersIds.Count; i++)
            {
                stringBuilder.AppendFormat("<Id>{0}</Id>", reportTask.usersIds[i]);
            }
            stringBuilder.Append("</UsersIds>");

            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportTaskDelete(IEnumerable<string> taskIds)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportTaskDelete\">");
            stringBuilder.Append("<TasksIds>");

            foreach (var task in taskIds)
                stringBuilder.AppendFormat("<Id>{0}</Id>", task);

            stringBuilder.Append("</TasksIds>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportHistorySet(ReportBuilder report)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportHistorySet\">");
            stringBuilder.AppendFormat("<TaskId>{0}</TaskId>", report.Info.taskId);
            stringBuilder.AppendFormat("<RunTime>{0}</RunTime>", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            stringBuilder.AppendFormat("<FromTime>{0}</FromTime>", report.Info.from);
            stringBuilder.AppendFormat("<ToTime>{0}</ToTime>", report.Info.to);
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportHistoryUpdate(ReportHistorySet reportHistorySet, int isError, string path)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportHistoryUpdate\">");
            stringBuilder.AppendFormat("<Id>{0}</Id>", reportHistorySet.reportHistoryInfo.id);
            stringBuilder.AppendFormat("<TableName>{0}</TableName>", reportHistorySet.reportHistoryInfo.tableName);
            stringBuilder.AppendFormat("<Status>{0}</Status>", isError);
            stringBuilder.AppendFormat("<Path>{0}</Path>", path);
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string ReportHistoryGet(ReportHistory reportHistory)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CWebReportHistoryGet\">");
            stringBuilder.AppendFormat("<UserId>{0}</UserId>", reportHistory.userId);
            stringBuilder.AppendFormat("<From>{0}</From>", reportHistory.fromTime);
            stringBuilder.AppendFormat("<To>{0}</To>", reportHistory.toTime);
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }


        public static string QueryCasesHistoryGet(string from, string to, IEnumerable<string> atmIds)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CCasesQuery\">");
            stringBuilder.AppendFormat("<From>{0}</From>", from);
            stringBuilder.AppendFormat("<To>{0}</To>", to);
            stringBuilder.Append("<Atms>");

            foreach (var atm in atmIds)
                stringBuilder.AppendFormat("<Atm>{0}</Atm>", atm);

            stringBuilder.Append("</Atms>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string QueryEventsHistoryGet(string from, string to, IEnumerable<string> atmIds)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CEventsQuery\">");
            stringBuilder.AppendFormat("<From>{0}</From>", from);
            stringBuilder.AppendFormat("<To>{0}</To>", to);
            stringBuilder.Append("<Atms>");

            foreach (var atm in atmIds)
                stringBuilder.AppendFormat("<Atm>{0}</Atm>", atm);

            stringBuilder.Append("</Atms>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }


        public static string QueryUsersHistoryGet(string from, string to, IEnumerable<string> users)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CAuditHistoryGet\">");
            stringBuilder.AppendFormat("<From>{0}</From>", from);
            stringBuilder.AppendFormat("<To>{0}</To>", to);
            stringBuilder.Append("<Users>");

            foreach (var user in users)
                stringBuilder.AppendFormat("<User>{0}</User>", user);

            stringBuilder.Append("</Users>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string QueryGetEvents(IEnumerable<string> evts)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CEvtsQuery\">");
            stringBuilder.Append("<Ids>");

            foreach (var evt in evts)
                stringBuilder.AppendFormat("<Id>{0}</Id>", evt);

            stringBuilder.Append("</Ids>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string QueryGetFunctions()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CFunctionsQuery\">");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string QueryGetUsers()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CUsersQuery\">");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }

        public static string QueryGetAttrHistory(string from, string to, IEnumerable<string> atmIds, string attrId)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"CAttributeHistoryGet\">");
            stringBuilder.AppendFormat("<From>{0}</From>", from);
            stringBuilder.AppendFormat("<To>{0}</To>", to);
            stringBuilder.Append("<Atms>");

            foreach (var atm in atmIds)
                stringBuilder.AppendFormat("<Atm>{0}</Atm>", atm);

            stringBuilder.Append("</Atms>");
            stringBuilder.AppendFormat("<Attributes><AttrId>{0}</AttrId></Attributes>", attrId);
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }
    }
}