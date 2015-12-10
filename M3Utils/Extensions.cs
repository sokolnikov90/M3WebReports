using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace M3Utils
{
    public static class Extensions
    {
        public static void ForEach<T>(this IEnumerable<T> ie, Action<T> action)
        {
            foreach (var i in ie)
            {
                action(i);
            }
        }

        public static bool IsNullOrEmpty<T>(this IEnumerable<T> items)
        {
            return items == null || !items.Any();
        }

        public static DayOfWeekRus DayOfWeekRus(this DateTime dt)
        {
            return (DayOfWeekRus)dt.DayOfWeek;
        }
    }

    public enum DayOfWeekRus
    {
        Воскресенье = 0,
        Понедельник = 1,
        Вторник = 2,
        Среда = 3,
        Четверг = 4,
        Пятница = 5,
        Суббота = 6
    }
}
