using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using NetOffice.PowerPointApi;
using BetterWPP.Server;

namespace BetterWPP.Server
{
    /// <summary>
    /// 路由分发：按 METHOD + PATH 调对应处理器。
    /// 第一批：health/status/setup/页数/当前页，验证骨架。
    /// </summary>
    internal sealed class Router
    {
        private readonly Application _app;
        private readonly HttpServer _server;
        public string ContentType { get; private set; }

        public Router(Application app, HttpServer server)
        {
            _app = app;
            _server = server;
            ContentType = "application/json; charset=utf-8";
        }

        public byte[] Dispatch(string method, string path, HttpListenerRequest request)
        {
            // GET 路由表
            if (method == "GET")
            {
                switch (path)
                {
                    case "/api/health": return Health();
                    case "/api/status": return Status();
                    case "/api/setup": return Setup();
                    case "/api/slides/count": return SlideCount();
                    case "/api/slides/active": return ActiveSlide();
                }
            }
            throw new HttpApiException(404, "not found: " + method + " " + path);
        }

        private Presentation CurrentPresentation()
        {
            var pres = _app.ActivePresentation;
            if (pres == null) throw new HttpApiException(409, "no active presentation");
            return pres;
        }

        private byte[] Health()
        {
            var f = new Dictionary<string, string>
            {
                { "ok", Json.Bool(true) },
                { "service", Json.Str("BetterWPP") },
                { "time", Json.Str(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")) }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] Status()
        {
            var f = new Dictionary<string, string>
            {
                { "running", Json.Bool(_server.IsRunning) },
                { "port", _server.Port.ToString() },
                { "mode", Json.Str(_server.Mode.ToString()) },
                { "urls", Json.Str(string.Join(",", _server.GetUrls())) }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] Setup()
        {
            var pres = CurrentPresentation();
            var ps = pres.PageSetup;
            var f = new Dictionary<string, string>
            {
                { "width", Json.Num(ps.SlideWidth) },
                { "height", Json.Num(ps.SlideHeight) },
                { "orientation", Json.Str(ps.SlideOrientation.ToString()) },
                { "presentation", Json.Str(pres.Name) },
                { "fullName", Json.Str(pres.FullName) }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] SlideCount()
        {
            var pres = CurrentPresentation();
            var f = new Dictionary<string, string>
            {
                { "count", pres.Slides.Count.ToString() },
                { "presentation", Json.Str(pres.Name) }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] ActiveSlide()
        {
            var pres = CurrentPresentation();
            int idx = 1;
            try
            {
                var win = _app.ActiveWindow;
                if (win != null && win.View != null)
                {
                    object slideObj = win.View.Slide;
                    if (slideObj is NetOffice.PowerPointApi.Slide slide) idx = slide.SlideIndex;
                }
            }
            catch { }
            var f = new Dictionary<string, string>
            {
                { "activeSlideIndex", idx.ToString() },
                { "count", pres.Slides.Count.ToString() }
            };
            return Utf8(Json.Obj(f));
        }

        private static byte[] Utf8(string json) { return Encoding.UTF8.GetBytes(json); }
    }
}
