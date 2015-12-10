namespace M3Reports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;

    using M3Atms;

    using M3Dictionaries;

    using M3Incidents;

    using M3IPClient;

    /// <summary>
    /// Класс содержит данные, полученные от сервиса WP
    /// </summary>
    public class ReportData
    {
        public static bool ExistInfo<TInfo>(TInfo info) where TInfo : IDataInfo
        {
            if ((info != null) && (info.isError == 0))
                return true;

            return false;
        }

        public M3Dictionaries.GetAll.Info DictionariesInfo
        {
            get
            {
                if (ExistInfo(this.DictionariesGet.info))
                    return this.DictionariesGet.info; //AllAtms

                throw new NullReferenceException("this.DictionariesGet.info.isError != 0");
            }
        }

        public List<Info> AtmInfo
        {
            get
            {
                if (ExistInfo(this.AtmInfoGet.info))
                    return this.AtmInfoGet.info.data;

                throw new NullReferenceException("this.atmInfoGet.info.isError != 0");
            }
        }

        public List<GroupsGet.AtmGroup> AtmGroups
        {
            get
            {
                if (ExistInfo(this.AtmGroupsGet.info))
                    return this.AtmGroupsGet.info.usersGroup[0].atmGroups; //AllAtms

                throw new NullReferenceException("this.AtmGroupsGet.info.isError != 0");                
            }
        }

        public List<Incident> Incidents
        {
            get
            {
                if (ExistInfo(this.IncidentsGet.incidentInfo))
                    return this.IncidentsGet.incidentInfo.data;

                throw new NullReferenceException("this.IncidentsGet.incidentInfo.isError != 0");    
            }
        }

        public ReportHistorySet ReportHistorySet;
        public ReportHistoryUpdate ReportHistoryUpdate;

        public M3Dictionaries.GetAll DictionariesGet;

        public InfoGet AtmInfoGet;
        public GroupsGet AtmGroupsGet;

        public IncidentGet QueryIncident;
        public M3Incidents.GetAll IncidentsGet;

        public CountsGet AtmCountsGet;
        public BNACountsGet AtmBNACountsGet;
        public List<CountsGet.AtmCountsData> AtmCounts;
        public List<BNACountsGet.AtmBNACountsData> AtmBNACounts;

        public WithdrawHistoryGet WithdrawHistoryGet;
        public MessageHistoryGet MessageHistoryGet;     
    }
}
