using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace M3IPClient
{
    public static class Queries
    {
        public static string Login(string login, string password)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("<Message>");
            stringBuilder.Append("<Request name=\"COperLogin\">");
            stringBuilder.AppendFormat("<Login>{0}</Login>", login);
            stringBuilder.AppendFormat("<Password>{0}</Password>", password);
            stringBuilder.Append("<AppType>M3Web</AppType>");
            stringBuilder.Append("</Request>");
            stringBuilder.Append("</Message>");

            return stringBuilder.ToString();
        }
    }
}