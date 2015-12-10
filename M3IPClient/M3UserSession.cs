using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Web;

namespace M3IPClient
{
    using System.Web.Configuration;

    public abstract class M3UserSession
    {
        private static string bankName;

        public static string BankName
        {
            get
            {
                if (bankName == null)
                {
                    bankName = "DefaultBank";

                    try
                    {
                        bankName = WebConfigurationManager.AppSettings["BankName"];
                    }
                    catch (Exception ex)
                    {
                        M3Utils.Log.Instance.Info("Get BankName error: " + ex.ToString());
                    }                 
                }

                return bankName;
            }
        }

        public IPClient connection = new IPClient();

        public EventWaitHandle ewh = new EventWaitHandle(false, EventResetMode.AutoReset);

        public SignIn signin = new SignIn();

        public string requestName;

        public M3UserSession(string ip, int port, string login, string password)
        {
            this.signin.info.isError = this.Connect(ip, port, login, password);
        }

        public M3UserSession(string ip, int port, string login, string password, string getJSON): this(ip, port, login, password)
        {
            M3Utils.Log.Instance.Info(this.GetType() + ".GetJSON: " + getJSON);
        }

        private int Connect(string ip, int port, string login, string password)
        {
            password = M3Utils.CryptographyHelper.CryptPassword8(password);

            try
            {
                this.connection.ReadEvent += this.IPRead;

                if (this.connection.Connect(ip, port))
                {
                    this.connection.Write(Queries.Login(login, password));

                    this.ewh.Reset();
                    this.ewh.WaitOne();

                    this.connection.ReadEvent -= this.IPRead;
                    
                    return 0;
                }
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info(
                    "IPClientDecorator.Connect(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }

            this.connection.ReadEvent -= this.IPRead;

            return 1;
        }

        private void IPRead(string message, bool complit)
        {
            XmlDocument xmlDocument = new XmlDocument();

            try
            {
                xmlDocument.LoadXml(message);

                var messageNode = xmlDocument.SelectSingleNode("Message");

                this.requestName = messageNode.SelectSingleNode("Request/./@name").InnerText;

                switch (this.requestName)
                {
                    case "CAdminError":
                        this.signin.ParseMessage(messageNode);
                        break;
                }
            }
            catch (Exception exp)
            {
                this.signin.info.isError = 1;

                M3Utils.Log.Instance.Info(
                    "IPClientDecorator.IPRead(...) exception:",
                    exp.Message,
                    exp.Source,
                    exp.StackTrace);
            }
            finally
            {
                this.ewh.Set();
            }
        }
    }
}