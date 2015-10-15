using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Web;

namespace KGSA
{
    public class WebServer
    {
        private readonly HttpListener _listener = new HttpListener();
        private readonly Func<HttpListenerRequest, string> _responderMethod;
        private AppSettings appConfig;
        public KgsaServer server;

        public WebServer(string[] prefixes, Func<HttpListenerRequest, string> method)
        {
            try
            {
                if (!HttpListener.IsSupported)
                    throw new NotSupportedException("Needs Windows XP SP2, Server 2003 or later.");

                if (prefixes == null || prefixes.Length == 0)
                    throw new ArgumentException("prefixes");

                if (method == null)
                    throw new ArgumentException("method");

                foreach (string s in prefixes)
                    _listener.Prefixes.Add(s);

                _responderMethod = method;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public WebServer(Func<HttpListenerRequest, string> method, params string[] prefixes) : this(prefixes, method) { }

        public bool IsOnline()
        {
            try
            {
                if (_listener != null)
                    return _listener.IsListening;
                else
                    return false;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        private bool CheckIdentity(HttpListenerContext ctx)
        {
            if (appConfig.webserverRequireSimpleAuthentication)
            {
                HttpListenerBasicIdentity identity = (HttpListenerBasicIdentity)ctx.User.Identity;
                var aes = new SimpleAES();
                if (appConfig.webserverPassword.Length > 0)
                {
                    string pass = aes.DecryptString(appConfig.webserverPassword);
                    if (identity.Name == appConfig.webserverUser && identity.Password == pass)
                        return true;
                }
            }
            else
                return true;

            ctx.Response.StatusCode = 401;
            return false;
        }

        public void Run()
        {
            ThreadPool.QueueUserWorkItem((o) =>
            {
                try
                {
                    while (_listener.IsListening)
                    {
                        ThreadPool.QueueUserWorkItem((c) =>
                        {
                            var ctx = c as HttpListenerContext;
                            try
                            {
                                if (CheckIdentity(ctx)) // Check credentials
                                {
                                    server.ProcessRequest(ctx);
                                }
                            }
                            catch { } // suppress any exceptions
                            finally
                            {
                                // always close the stream
                                ctx.Response.OutputStream.Close();
                            }
                        }, _listener.GetContext());
                    }
                }
                catch { } // suppress any exceptions
            });
        }

        public void Start()
        {
            try
            {
                if (appConfig.webserverRequireSimpleAuthentication)
                    _listener.AuthenticationSchemes = AuthenticationSchemes.Basic;
                _listener.Start();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void Stop()
        {
            _listener.Stop();
            _listener.Close();
        }

        public void Settings(AppSettings app, KgsaServer serverObj)
        {
            this.appConfig = app;
            this.server = serverObj;
        }
    }
}
