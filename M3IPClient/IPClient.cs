using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Net;
using System.Threading;

namespace M3IPClient
{
    public class IPClient
    {
        public delegate void ReadDelegate(string message, bool complit);
        public event ReadDelegate ReadEvent;

        Thread thread;

        private TcpClient tcpClient;
        private NetworkStream networkStream;

        public bool Connect(string ip, int port)
        {
            try
            {
                Disconnect();

                AutoResetEvent connectDone = new AutoResetEvent(false);

                this.tcpClient = new TcpClient();

                this.tcpClient.BeginConnect(ip, port,
                    delegate(IAsyncResult ar)
                        {
                            try
                            {
                                tcpClient.EndConnect(ar);
                            }
                            catch { }

                            connectDone.Set();
                        }, this.tcpClient);

                if (!connectDone.WaitOne())
                    throw new Exception();

                networkStream = this.tcpClient.GetStream();

                thread = new Thread(WaitRequest);

                thread.IsBackground = true;
                thread.Name = "ConnectThread";

                thread.Start();

                return true;
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info("IPClientDecorator.Connect(string ip, int port) exception:");
                M3Utils.Log.Instance.Info(exp.Message);
                M3Utils.Log.Instance.Info(exp.Source);
                M3Utils.Log.Instance.Info(exp.StackTrace);
            }

            return false;
        }

        public void Disconnect()
        {
            try
            {
                if (networkStream != null)
                    networkStream.Close();

                if (this.tcpClient != null)
                    this.tcpClient.Close();

                if (thread != null)
                    thread.Abort();
            }
            catch{}
        }

        private void WaitRequest()
        {
            StringBuilder dataReceivedBuilder = new StringBuilder();

            string buffer = "",
                   dataReceived = "";

            byte[] bytes = new byte[65535];

            try
            {
                while (true)
                {
                    int numberOfBytes = networkStream.Read(bytes, 0, bytes.Length);

                    if (numberOfBytes <= 0)
                        break;

                    buffer = Encoding.Default.GetString(bytes, 0, numberOfBytes);

                    dataReceivedBuilder.Append(buffer);

                    if (dataReceivedBuilder.Length >= 10)
                    {
                        buffer = buffer.Insert(0, dataReceivedBuilder.ToString(dataReceivedBuilder.Length - 10, 10));
                    }
                    else
                    {
                        if (dataReceivedBuilder.Length > 0)
                        {
                            buffer = buffer.Insert(0, dataReceivedBuilder.ToString());
                        }
                    }

                    if (buffer.Contains("</Message>"))
                    {
                        dataReceived = dataReceivedBuilder.ToString();

                        try
                        {
                            bool bExit = false;

                            do
                            {
                                int index = dataReceived.IndexOf("</Message>");

                                if (index > -1)
                                {
                                    string message = dataReceived.Substring(0, index + 10);

                                    dataReceived = dataReceived.Remove(0, index + 10);

                                    if (ReadEvent != null)
                                        ReadEvent.Invoke(message, true);
                                }
                                else
                                {
                                    bExit = true;
                                }
                            }
                            while (!bExit);

                            dataReceivedBuilder.Length = 0;
                            dataReceivedBuilder.Append(dataReceived);
                        }
                        catch { }
                    }
                }

            }
            catch { }
        }

        public void Write(string buffer)
        {
            try
            {
                networkStream.Write(Encoding.Default.GetBytes(buffer), 0, buffer.Length);
            }
            catch { }
        }

        public void Write(string buffer, EventWaitHandle ewh)
        {
            try
            {
                networkStream.Write(Encoding.Default.GetBytes(buffer), 0, buffer.Length);
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info("IPClient.Write(string buffer, EventWaitHandle ewh) exception:");
                M3Utils.Log.Instance.Info(exp.Message);
                M3Utils.Log.Instance.Info(exp.Source);
                M3Utils.Log.Instance.Info(exp.StackTrace);
            }
            finally
            {
                ewh.Reset();
                ewh.WaitOne();
            }
        }

        public void Write(byte[] buffer)
        {
            try
            {
                networkStream.Write(buffer, 0, buffer.Length);
            }
            catch { }
        }

        public void Write(byte[] buffer, EventWaitHandle ewh)
        {
            try
            {
                networkStream.Write(buffer, 0, buffer.Length);
            }
            catch (Exception exp)
            {
                M3Utils.Log.Instance.Info("IPClient.Write(byte[] buffer, EventWaitHandle ewh) exception:");
                M3Utils.Log.Instance.Info(exp.Message);
                M3Utils.Log.Instance.Info(exp.Source);
                M3Utils.Log.Instance.Info(exp.StackTrace);
            }
            finally
            {
                ewh.Reset();
                ewh.WaitOne();
            }
        }
    }
}