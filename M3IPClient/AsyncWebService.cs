using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Services;
using System.Diagnostics;

namespace M3IPClient
{
    abstract public class AsyncWebService : WebService
    {
        protected static readonly Func<object[], string> StubCallBack = WebMethodCallBack;

        private static string WebMethodCallBack(object[] webMethodState)
        {
            string result = string.Empty;

            try
            {
                IAsyncRequestFacade report = CreateM3UserSession(webMethodState);
                report.DoStuff();
                result = report.Response();
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(
                    String.Join(Environment.NewLine, new []
                    {
                        "WebMethodCallBack(...) exception:",
                        exp.Message,
                        exp.Source,
                        exp.StackTrace
                    }));
            }

            return result;
        }

        private static IAsyncRequestFacade CreateM3UserSession(object[] webMethodState)
        {
            Type typeOfObject = (Type)webMethodState[0];
            object[] constructorParams = webMethodState.Skip(1).ToArray();

            var ipClient = Activator.CreateInstance(typeOfObject, constructorParams) as IAsyncRequestFacade;

            if (ipClient == null)
            {
                throw new NullReferenceException("ipClient");
            }

            return ipClient;
        }
    }
}
