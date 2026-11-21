using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using NetOffice.PowerPointApi;

namespace BetterWPP.Server
{
    public enum BindMode
    {
        /// <summary>仅本机 127.0.0.1</summary>
        Loopback,
        /// <summary>局域网 0.0.0.0，同网段可访问</summary>
        Lan
    }

    /// <summary>
    /// 单个 PowerPoint 实例对应的本地 HTTP 服务。
    /// 只操作本实例的 ActivePresentation，不跨窗口；每个加载项实例持有一份。
    /// </summary>
    public sealed class HttpServer : IDisposable
    {
        private readonly Application _application;
        private HttpListener _listener;
        private Thread _thread;
        private bool _running;
        private int _port;
        private BindMode _mode;
        private string _lastError;

        public bool IsRunning { get { return _running; } }
        public int Port { get { return _port; } }
        public BindMode Mode { get { return _mode; } }
        public string LastError { get { return _lastError; } }

        public HttpServer(Application application)
        {
            _application = application;
            _mode = BindMode.Loopback;
        }

        public void SetMode(BindMode mode)
        {
            _mode = mode;
        }

        /// <summary>启动 HttpListener，绑定随机端口。</summary>
        public bool Start()
        {
            if (_running) return true;
            _lastError = null;
            try
            {
                _port = GetRandomPort();
                string prefix = _mode == BindMode.Lan
                    ? "http://*:" + _port + "/"
                    : "http://127.0.0.1:" + _port + "/";

                _listener = new HttpListener();
                _listener.Prefixes.Add(prefix);
                _listener.Start();
                _running = true;
                _thread = new Thread(ListenLoop) { IsBackground = true, Name = "BetterWPP-Http" };
                _thread.Start();
                return true;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                _running = false;
                return false;
            }
        }

        public void Stop()
        {
            _running = false;
            try { if (_listener != null) _listener.Stop(); } catch { }
            try { if (_thread != null && _thread.IsAlive) _thread.Join(800); } catch { }
            try { if (_listener != null) _listener.Close(); } catch { }
            _listener = null;
            _thread = null;
        }

        /// <summary>对外展示的访问地址（局域网模式取本机 IPv4）。</summary>
        public string[] GetUrls()
        {
            if (!_running) return new string[0];
            var list = new List<string>();
            list.Add("http://127.0.0.1:" + _port + "/");
            if (_mode == BindMode.Lan)
            {
                string lanIp = GetLanIpv4();
                if (!string.IsNullOrEmpty(lanIp)) list.Add("http://" + lanIp + ":" + _port + "/");
            }
            return list.ToArray();
        }

        private static int GetRandomPort()
        {
            var l = new TcpListener(IPAddress.Loopback, 0);
            l.Start();
            int port = ((IPEndPoint)l.LocalEndpoint).Port;
            l.Stop();
            return port;
        }

        private static string GetLanIpv4()
        {
            try
            {
                string host = Dns.GetHostName();
                foreach (IPAddress ip in Dns.GetHostAddresses(host))
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork &&
                        !IPAddress.IsLoopback(ip)) return ip.ToString();
                }
            }
            catch { }
            return null;
        }

        private void ListenLoop()
        {
            while (_running)
            {
                HttpListenerContext ctx = null;
                try
                {
                    ctx = _listener.GetContext();
                    Handle(ctx);
                }
                catch (HttpListenerException) { break; }
                catch (Exception ex)
                {
                    _lastError = ex.Message;
                }
                finally
                {
                    if (ctx != null) try { ctx.Response.Close(); } catch { }
                }
            }
        }

        private void Handle(HttpListenerContext ctx)
        {
            string method = ctx.Request.HttpMethod;
            string path = ctx.Request.Url.AbsolutePath.TrimEnd('/');
            if (string.IsNullOrEmpty(path)) path = "/";

            try
            {
                var router = new Router(_application, this);
                byte[] body = router.Dispatch(method, path, ctx.Request);
                string contentType = router.ContentType;
                Write(ctx, body, contentType, 200);
            }
            catch (HttpApiException api)
            {
                string json = Json.Obj(new Dictionary<string, string>
                {
                    { "ok", Json.Bool(false) },
                    { "error", Json.Str(api.Message) }
                });
                Write(ctx, Encoding.UTF8.GetBytes(json), "application/json; charset=utf-8", api.Status);
            }
            catch (Exception ex)
            {
                string json = Json.Obj(new Dictionary<string, string>
                {
                    { "ok", Json.Bool(false) },
                    { "error", Json.Str(ex.Message) }
                });
                Write(ctx, Encoding.UTF8.GetBytes(json), "application/json; charset=utf-8", 500);
            }
        }

        private static void Write(HttpListenerContext ctx, byte[] body, string contentType, int status)
        {
            ctx.Response.StatusCode = status;
            ctx.Response.ContentType = contentType;
            ctx.Response.ContentLength64 = body.Length;
            ctx.Response.OutputStream.Write(body, 0, body.Length);
        }

        public void Dispose() { Stop(); }
    }

    internal sealed class HttpApiException : Exception
    {
        public int Status { get; private set; }
        public HttpApiException(int status, string message) : base(message) { Status = status; }
    }
}
