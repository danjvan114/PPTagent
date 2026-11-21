using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Enums;

namespace BetterWPP.Server
{
    /// <summary>
    /// 写操作处理器（POST）。参数走 JSON body；图片支持本地路径，二进制上传后续接。
    /// 所有改动直接作用于本实例 ActivePresentation。
    /// </summary>
    internal sealed class Writer
    {
        private readonly Application _app;
        private readonly Router _router;

        public Writer(Application app, Router router)
        {
            _app = app;
            _router = router;
        }

        public byte[] Handle(string[] s, HttpListenerRequest request)
        {
            string body = ReadBody(request);
            var json = JsonParse.Parse(body);
            var pres = Inspector.Current(_app);

            // 19: 页面顺序 / 删除 / 添加
            //   POST /api/slides/move      {from, to}
            //   POST /api/slides/delete    {index}
            //   POST /api/slides/add        {index?, layout?}
            if (s.Length == 3 && s[0] == "api" && s[1] == "slides")
            {
                switch (s[2])
                {
                    case "move": return SlideMove(pres, json);
                    case "delete": return SlideDelete(pres, json);
                    case "add": return SlideAdd(pres, json);
                }
            }

            // 14: 编辑元素指定属性（可带图片路径替换图片）
            //   POST /api/slides/{i}/shapes/{id}  {left,top,width,height,text,name,picturePath?}
            if (s.Length == 5 && s[0] == "api" && s[1] == "slides" && s[3] == "shapes")
                return ShapeUpdate(pres, s[2], s[4], json);

            // 18: 切换特效 编辑/删除/添加
            //   POST /api/slides/{i}/transition  {entryEffect?, speed?, advanceOnClick?, advanceTime?, delete?}
            if (s.Length == 4 && s[0] == "api" && s[1] == "slides" && s[3] == "transition")
                return TransitionUpdate(pres, s[2], json);

            // 15/16/17/21: 动画
            //   POST /api/slides/{i}/animations/add      {shapeId, effectType?, trigger?}
            //   POST /api/slides/{i}/animations/update   {shapeId, index?, duration?, trigger?}
            //   POST /api/slides/{i}/animations/delete   {shapeId, index?}
            if (s.Length == 5 && s[0] == "api" && s[1] == "slides" && s[3] == "animations")
                return AnimationUpdate(pres, s[2], s[4], json);

            throw new HttpApiException(404, "not found: POST " + string.Join("/", s));
        }

        // ---------- 页面 ----------
        private byte[] SlideMove(Presentation pres, Dictionary<string, object> j)
        {
            int from = ReqInt(j, "from");
            int to = ReqInt(j, "to");
            var slide = pres.Slides[from];
            slide.MoveTo(to);
            return Ok(new Dictionary<string, string>
            {
                { "moved", Json.Bool(true) },
                { "from", from.ToString() },
                { "to", to.ToString() }
            });
        }

        private byte[] SlideDelete(Presentation pres, Dictionary<string, object> j)
        {
            int index = ReqInt(j, "index");
            pres.Slides[index].Delete();
            return Ok(new Dictionary<string, string>
            {
                { "deleted", Json.Bool(true) },
                { "index", index.ToString() }
            });
        }

        private byte[] SlideAdd(Presentation pres, Dictionary<string, object> j)
        {
            int index = j.ContainsKey("index") ? Convert.ToInt32(j["index"], CultureInfo.InvariantCulture) : pres.Slides.Count + 1;
            PpSlideLayout layout = PpSlideLayout.ppLayoutBlank;
            if (j.ContainsKey("layout"))
            {
                try { layout = (PpSlideLayout)Enum.Parse(typeof(PpSlideLayout), j["layout"].ToString(), true); }
                catch { }
            }
            var slide = pres.Slides.Add(index, layout);
            return Ok(new Dictionary<string, string>
            {
                { "added", Json.Bool(true) },
                { "index", slide.SlideIndex.ToString() }
            });
        }

        // ---------- 元素 ----------
        private byte[] ShapeUpdate(Presentation pres, string slideIndex, string shapeId, Dictionary<string, object> j)
        {
            var shape = Inspector.FindShape(pres, slideIndex, shapeId);

            if (j.ContainsKey("left")) shape.Left = Convert.ToSingle(j["left"], CultureInfo.InvariantCulture);
            if (j.ContainsKey("top")) shape.Top = Convert.ToSingle(j["top"], CultureInfo.InvariantCulture);
            if (j.ContainsKey("width")) shape.Width = Convert.ToSingle(j["width"], CultureInfo.InvariantCulture);
            if (j.ContainsKey("height")) shape.Height = Convert.ToSingle(j["height"], CultureInfo.InvariantCulture);
            if (j.ContainsKey("name")) shape.Name = j["name"].ToString();
            if (j.ContainsKey("text"))
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                    shape.TextFrame.TextRange.Text = j["text"].ToString();
            }

            // 图片路径替换：在同位置插入新图片，删除旧图
            if (j.ContainsKey("picturePath"))
            {
                string path = j["picturePath"].ToString();
                if (!File.Exists(path)) throw new HttpApiException(400, "picture not found: " + path);
                float left = shape.Left, top = shape.Top, width = shape.Width, height = shape.Height;
                string name = shape.Name;
                var slide = Inspector.SlideAt(pres, slideIndex);
                slide.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
                shape.Delete();
                return Ok(new Dictionary<string, string>
                {
                    { "pictureReplaced", Json.Bool(true) },
                    { "name", Json.Str(name) }
                });
            }

            return Ok(new Dictionary<string, string>
            {
                { "updated", Json.Bool(true) },
                { "id", shape.Id.ToString() }
            });
        }

        // ---------- 切换特效 ----------
        private byte[] TransitionUpdate(Presentation pres, string slideIndex, Dictionary<string, object> j)
        {
            var slide = Inspector.SlideAt(pres, slideIndex);
            var t = slide.SlideShowTransition;

            if (j.ContainsKey("delete"))
            {
                t.EntryEffect = PpEntryEffect.ppEffectCut;
                return Ok(new Dictionary<string, string> { { "transitionReset", Json.Bool(true) } });
            }
            if (j.ContainsKey("entryEffect"))
            {
                try { t.EntryEffect = (PpEntryEffect)Enum.Parse(typeof(PpEntryEffect), j["entryEffect"].ToString(), true); }
                catch (Exception ex) { throw new HttpApiException(400, "bad entryEffect: " + ex.Message); }
            }
            if (j.ContainsKey("speed"))
            {
                try { t.Speed = (PpTransitionSpeed)Enum.Parse(typeof(PpTransitionSpeed), j["speed"].ToString(), true); }
                catch { }
            }
            if (j.ContainsKey("advanceOnClick")) t.AdvanceOnClick = BoolTri(j["advanceOnClick"]);
            if (j.ContainsKey("advanceTime")) t.AdvanceTime = Convert.ToSingle(j["advanceTime"], CultureInfo.InvariantCulture);

            return Ok(new Dictionary<string, string>
            {
                { "transitionUpdated", Json.Bool(true) },
                { "slideIndex", slide.SlideIndex.ToString() }
            });
        }

        // ---------- 动画 ----------
        private byte[] AnimationUpdate(Presentation pres, string slideIndex, string action, Dictionary<string, object> j)
        {
            var slide = Inspector.SlideAt(pres, slideIndex);
            var main = slide.TimeLine.MainSequence;

            switch (action)
            {
                case "add":
                {
                    int shapeId = ReqInt(j, "shapeId");
                    var shape = FindShapeOnSlide(slide, shapeId);
                    // AddEffect(shape, effectId, level, trigger, index)
                    MsoAnimEffect effect = MsoAnimEffect.msoAnimAppear;
                    if (j.ContainsKey("effectType"))
                    {
                        try { effect = (MsoAnimEffect)Enum.Parse(typeof(MsoAnimEffect), j["effectType"].ToString(), true); }
                        catch { }
                    }
                    MsoAnimTriggerType trigger = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                    if (j.ContainsKey("trigger"))
                    {
                        try { trigger = (MsoAnimTriggerType)Enum.Parse(typeof(MsoAnimTriggerType), j["trigger"].ToString(), true); }
                        catch { }
                    }
                    var ef = main.AddEffect(shape, effect, MsoAnimateByLevel.msoAnimateTextByAllLevels, trigger, main.Count + 1);
                    return Ok(new Dictionary<string, string>
                    {
                        { "animationAdded", Json.Bool(true) },
                        { "index", ef.Index.ToString() }
                    });
                }
                case "update":
                {
                    int index = ReqInt(j, "index");
                    var ef = main[index];
                    if (j.ContainsKey("duration")) ef.Timing.Duration = Convert.ToSingle(j["duration"], CultureInfo.InvariantCulture);
                    if (j.ContainsKey("triggerDelayTime")) ef.Timing.TriggerDelayTime = Convert.ToSingle(j["triggerDelayTime"], CultureInfo.InvariantCulture);
                    return Ok(new Dictionary<string, string> { { "animationUpdated", Json.Bool(true) }, { "index", index.ToString() } });
                }
                case "delete":
                {
                    int index = ReqInt(j, "index");
                    main[index].Delete();
                    return Ok(new Dictionary<string, string> { { "animationDeleted", Json.Bool(true) }, { "index", index.ToString() } });
                }
            }
            throw new HttpApiException(400, "unknown animation action: " + action);
        }

        private static Shape FindShapeOnSlide(Slide slide, int shapeId)
        {
            foreach (Shape s in slide.Shapes) if (s.Id == shapeId) return s;
            throw new HttpApiException(404, "shape not found: " + shapeId);
        }

        // ---------- 辅助 ----------
        private static int ReqInt(Dictionary<string, object> j, string key)
        {
            if (!j.ContainsKey(key)) throw new HttpApiException(400, "missing param: " + key);
            return Convert.ToInt32(j[key], CultureInfo.InvariantCulture);
        }

        private static MsoTriState BoolTri(object o)
        {
            bool b = o is bool ? (bool)o : Convert.ToBoolean(o, CultureInfo.InvariantCulture);
            return b ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }

        private static byte[] Ok(Dictionary<string, string> fields)
        {
            fields["ok"] = Json.Bool(true);
            return Router.Utf8(Json.Obj(fields));
        }

        private static string ReadBody(HttpListenerRequest request)
        {
            if (!request.HasEntityBody) return "";
            using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                return reader.ReadToEnd();
        }
    }
}
