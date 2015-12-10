using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace M3Reports
{
    public class ReportTransactionItem
    {
        public string atmId { get; set; }        // Id банкомата
        public string dateTime { get; set; }     // Время транзакции
        public string amount { get; set; }       // Количество
        public string okb { get; set; }          // Буфер
        public string sum { get; set; }          // Сумма
        public string state { get; set; }        // Стейт
        public string currency { get; set; }     // Валюта
        public string dispense { get; set; }     // Выдано
        public string dispCass1 { get; set; }    // Выдано из 1-ой кассеты
        public string dispCass2 { get; set; }    // Выдано из 2-ой кассеты
        public string dispCass3 { get; set; }    // Выдано из 3-ой кассеты
        public string dispCass4 { get; set; }    // Выдано из 4-ой кассеты
        public string track2Data { get; set; }   // Карта
    }
}