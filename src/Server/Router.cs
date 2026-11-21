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
    /// GET 只读采集；POST 写操作在 Writer 中实现。
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
            var segments = Split(path);

            if (method == "GET") return Get(segments, request);
            if (method == "POST") return Post(segments, request);
            throw new HttpApiException(405, "method not allowed: " + method);
        }

        private static string[] Split(string path)
        {
            var list = new List<string>();
            foreach (var seg in path.Split('/'))
            {
                if (!string.IsNullOrEmpty(seg)) list.Add(Uri.UnescapeDataString(seg));
            }
            return list.ToArray();
        }

        // ---------------- GET ----------------
        private byte[] Get(string[] s, HttpListenerRequest request)
        {
            // /api/health
            if (Match(s, "api", "health")) return Health();
            // /api/status
            if (Match(s, "api", "status")) return Status();

            var pres = Inspector.Current(_app);

            // /api/setup  (20)
            if (Match(s, "api", "setup")) return Utf8(Json.Obj(Inspector.Setup(pres)));
            // /api/slides/count  (1)
            if (Match(s, "api", "slides", "count")) return Utf8(Json.Obj(new Dictionary<string, string>
            {
                { "count", pres.Slides.Count.ToString() }
            }));
            // /api/slides  (1)
            if (Match(s, "api", "slides")) return Utf8(Json.Obj(Inspector.Slides(pres)));
            // /api/slides/active  (3)
            if (Match(s, "api", "slides", "active")) return ActiveSlide(pres);

            // /api/slides/{i}/thumbnail  (2)
            if (s.Length == 4 && s[0] == "api" && s[1] == "slides" && s[3] == "thumbnail")
                return Thumbnail(pres, s[2], request);

            // /api/slides/{i}/shapes  (4)
            if (s.Length == 4 && s[0] == "api" && s[1] == "slides" && s[3] == "shapes")
                return Utf8(Json.Obj(Inspector.Shapes(pres, s[2])));

            // /api/slides/{i}/transition  (7)
            if (s.Length == 4 && s[0] == "api" && s[1] == "slides" && s[3] == "transition")
                return Utf8(Json.Obj(Inspector.Transition(pres, s[2])));

            // /api/slides/{i}/timeline  (8,9)
            if (s.Length == 4 && s[0] == "api" && s[1] == "slides" && s[3] == "timeline")
                return Utf8(Json.Obj(Inspector.Timeline(pres, s[2])));

            // /api/slides/{i}/shapes/{id}  (5)
            if (s.Length == 5 && s[0] == "api" && s[1] == "slides" && s[3] == "shapes")
                return Utf8(Json.Obj(Inspector.ShapeDetail(pres, s[2], s[4])));

            // /api/slides/{i}/shapes/{id}/media?prop=  (6) 图片/视频回二进制，其余回 JSON
            if (s.Length == 6 && s[0] == "api" && s[1] == "slides" && s[3] == "shapes" && s[5] == "media")
                return ShapeMedia(pres, s[2], s[4], request);

            // /api/animations/all  (10 内部: 全量动画特效枚举值)
            if (Match(s, "api", "animations", "all"))
                return EnumList("PpEntryEffect", typeof(NetOffice.PowerPointApi.Enums.PpEntryEffect));
            // /api/transitions/all  (11 内部: 全量切换特效枚举值)
            if (Match(s, "api", "transitions", "all"))
                return EnumList("PpEntryEffect", typeof(NetOffice.PowerPointApi.Enums.PpEntryEffect));
            // /api/animations/effects  (10 内部: 全量动画特效枚举值)
            if (Match(s, "api", "animations", "effects"))
                return EnumList("MsoAnimEffect", typeof(NetOffice.PowerPointApi.Enums.MsoAnimEffect));

            // /api/fonts  (12) /api/shapes  (13) — COM 不直接枚举，占位
            if (Match(s, "api", "fonts")) return NotImpl("fonts require PPTX package parsing");
            if (Match(s, "api", "shape-types")) return ShapeTypes();

            throw new HttpApiException(404, "not found: GET " + string.Join("/", s));
        }

        private bool Match(string[] s, params string[] parts)
        {
            if (s.Length != parts.Length) return false;
            for (int i = 0; i < parts.Length; i++) if (s[i] != parts[i]) return false;
            return true;
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

        private byte[] ActiveSlide(Presentation pres)
        {
            int idx = 1;
            try
            {
                var win = _app.ActiveWindow;
                if (win != null && win.View != null)
                {
                    object slideObj = win.View.Slide;
                    if (slideObj is Slide slide) idx = slide.SlideIndex;
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

        private byte[] Thumbnail(Presentation pres, string slideIndex, HttpListenerRequest request)
        {
            var slide = Inspector.SlideAt(pres, slideIndex);
            string dir = Path.Combine(Path.GetTempPath(), "betterwpp_thumb");
            Directory.CreateDirectory(dir);
            string file = Path.Combine(dir, "slide_" + slide.SlideIndex + ".png");
            if (File.Exists(file)) File.Delete(file);

            int width, height;
            int.TryParse(request.QueryString["width"], out width);
            int.TryParse(request.QueryString["height"], out height);

            if (width > 0 && height > 0) slide.Export(file, "PNG", width, height);
            else slide.Export(file, "PNG");

            if (!File.Exists(file)) throw new HttpApiException(500, "thumbnail export failed");
            ContentType = "image/png";
            return File.ReadAllBytes(file);
        }

        private byte[] ShapeMedia(Presentation pres, string slideIndex, string shapeId, HttpListenerRequest request)
        {
            var shape = Inspector.FindShape(pres, slideIndex, shapeId);
            string prop = request.QueryString["prop"];

            // 图片：导出为 PNG 二进制
            if (shape.Type == NetOffice.OfficeApi.Enums.MsoShapeType.msoPicture ||
                shape.Type == NetOffice.OfficeApi.Enums.MsoShapeType.msoLinkedPicture)
            {
                string dir = Path.Combine(Path.GetTempPath(), "betterwpp_thumb");
                Directory.CreateDirectory(dir);
                string file = Path.Combine(dir, "shape_" + shape.Id + ".png");
                if (File.Exists(file)) File.Delete(file);
                shape.Export(file, NetOffice.PowerPointApi.Enums.PpShapeFormat.ppShapeFormatPNG);
                if (File.Exists(file))
                {
                    ContentType = "image/png";
                    return File.ReadAllBytes(file);
                }
            }

            // 媒体（音频/视频）：COM 不直接导出二进制，返回描述 JSON，标 unsupported
            var f = new Dictionary<string, string>
            {
                { "id", shape.Id.ToString() },
                { "name", Json.Str(shape.Name) },
                { "type", Json.Str(shape.Type.ToString()) },
                { "binary", Json.Str("unsupported") },
                { "note", Json.Str("media binary export needs PPTX package parsing") }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] EnumList(string label, Type enumType)
        {
            var f = new Dictionary<string, string>
            {
                { "kind", Json.Str(label) },
                { "values", Inspector.EnumNames(enumType) }
            };
            return Utf8(Json.Obj(f));
        }

        private byte[] ShapeTypes()
        {
            return EnumList("MsoShapeType", typeof(NetOffice.OfficeApi.Enums.MsoShapeType));
        }

        private byte[] NotImpl(string reason)
        {
            ContentType = "application/json; charset=utf-8";
            return Utf8(Json.Obj(new Dictionary<string, string>
            {
                { "ok", Json.Bool(false) },
                { "error", Json.Str("not-implemented: " + reason) }
            }));
        }

        // ---------------- POST ----------------
        private byte[] Post(string[] s, HttpListenerRequest request)
        {
            var writer = new Writer(_app, this);
            return writer.Handle(s, request);
        }

        internal static byte[] Utf8(string json) { return Encoding.UTF8.GetBytes(json); }
    }
}
