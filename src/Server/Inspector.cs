using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;

namespace BetterWPP.Server
{
    /// <summary>
    /// 只读采集：把 NetOffice 包装对象转成可序列化的字典/原始类型。
    /// 取不到的字段不编造，标 unsupported / null。
    /// </summary>
    internal static class Inspector
    {
        public static Presentation Current(Application app)
        {
            var pres = app.ActivePresentation;
            if (pres == null) throw new HttpApiException(409, "no active presentation");
            return pres;
        }

        public static Slide SlideAt(Presentation pres, string index)
        {
            int i;
            if (!int.TryParse(index, out i) || i < 1 || i > pres.Slides.Count)
                throw new HttpApiException(400, "invalid slide index: " + index);
            return pres.Slides[i];
        }

        public static Shape FindShape(Presentation pres, string slideIndex, string shapeId)
        {
            var slide = SlideAt(pres, slideIndex);
            int id;
            if (!int.TryParse(shapeId, out id)) throw new HttpApiException(400, "invalid shape id");
            foreach (Shape s in slide.Shapes)
            {
                if (s.Id == id) return s;
            }
            throw new HttpApiException(404, "shape not found: " + shapeId);
        }

        // ---- 20: 画布大小与方向 ----
        public static Dictionary<string, string> Setup(Presentation pres)
        {
            var ps = pres.PageSetup;
            return new Dictionary<string, string>
            {
                { "width", Json.Num(ps.SlideWidth) },
                { "height", Json.Num(ps.SlideHeight) },
                { "orientation", Json.Str(ps.SlideOrientation.ToString()) },
                { "slideSize", Json.Str(ps.SlideSize.ToString()) },
                { "presentation", Json.Str(pres.Name) },
                { "fullName", Json.Str(pres.FullName) }
            };
        }

        // ---- 1: 页数 + 每页序号/标题 ----
        public static Dictionary<string, string> Slides(Presentation pres)
        {
            var arr = new List<string>();
            foreach (Slide s in pres.Slides)
            {
                string title = "";
                try { if (s.Shapes.HasTitle) title = s.Shapes.Title.TextFrame.TextRange.Text; } catch { }
                var f = new Dictionary<string, string>
                {
                    { "index", s.SlideIndex.ToString() },
                    { "slideId", s.SlideID.ToString() },
                    { "title", Json.Str(title) }
                };
                arr.Add(Json.Obj(f));
            }
            return new Dictionary<string, string>
            {
                { "count", pres.Slides.Count.ToString() },
                { "slides", "[" + string.Join(",", arr) + "]" }
            };
        }

        // ---- 4: 页面元素清单 ----
        public static Dictionary<string, string> Shapes(Presentation pres, string slideIndex)
        {
            var slide = SlideAt(pres, slideIndex);
            var arr = new List<string>();
            foreach (Shape s in slide.Shapes) arr.Add(ShapeJson(s, false));
            return new Dictionary<string, string>
            {
                { "slideIndex", slide.SlideIndex.ToString() },
                { "shapes", "[" + string.Join(",", arr) + "]" }
            };
        }

        // ---- 5: 单个元素属性 ----
        public static Dictionary<string, string> ShapeDetail(Presentation pres, string slideIndex, string shapeId)
        {
            var s = FindShape(pres, slideIndex, shapeId);
            var meta = new Dictionary<string, string>
            {
                { "slideIndex", SlideAt(pres, slideIndex).SlideIndex.ToString() }
            };
            meta["shape"] = ShapeJson(s, true);
            return meta;
        }

        private static string ShapeJson(Shape s, bool detail)
        {
            var f = new Dictionary<string, string>
            {
                { "id", s.Id.ToString() },
                { "name", Json.Str(s.Name) },
                { "type", Json.Str(s.Type.ToString()) },
                { "left", Json.Num(s.Left) },
                { "top", Json.Num(s.Top) },
                { "width", Json.Num(s.Width) },
                { "height", Json.Num(s.Height) },
                { "rotation", Json.Num(s.Rotation) },
                { "visible", Json.Bool(s.Visible == MsoTriState.msoTrue) },
                { "zOrder", s.ZOrderPosition.ToString() }
            };

            string text = "";
            try { if (s.HasTextFrame == MsoTriState.msoTrue) text = s.TextFrame.TextRange.Text; } catch { }
            f["text"] = Json.Str(text);

            // 图片/媒体标识
            bool isPicture = s.Type == MsoShapeType.msoPicture || s.Type == MsoShapeType.msoLinkedPicture;
            f["isPicture"] = Json.Bool(isPicture);
            if (isPicture)
            {
                // picId 复用内部 Shape.Id；媒体关系 ID 需解 PPTX 包，标 unsupported
                f["picId"] = s.Id.ToString();
                f["mediaRelationId"] = Json.Str("unsupported");
            }
            bool isMedia = s.Type == MsoShapeType.msoMedia || s.Type == MsoShapeType.msoWebVideo;
            f["isMedia"] = Json.Bool(isMedia);
            if (isMedia)
            {
                try { f["mediaType"] = Json.Str(s.MediaType.ToString()); } catch { f["mediaType"] = Json.Str("unknown"); }
            }

            if (detail)
            {
                try
                {
                    var pf = s.PictureFormat;
                    f["picture"] = Json.Obj(new Dictionary<string, string>
                    {
                        { "brightness", Json.Num(pf.Brightness) },
                        { "contrast", Json.Num(pf.Contrast) },
                        { "cropLeft", Json.Num(pf.CropLeft) },
                        { "cropRight", Json.Num(pf.CropRight) },
                        { "cropTop", Json.Num(pf.CropTop) },
                        { "cropBottom", Json.Num(pf.CropBottom) }
                    });
                }
                catch { f["picture"] = "null"; }

                try
                {
                    if (s.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var font = s.TextFrame.TextRange.Font;
                        f["font"] = Json.Obj(new Dictionary<string, string>
                        {
                            { "name", Json.Str(font.Name) },
                            { "size", Json.Num(font.Size) },
                            { "bold", Json.Bool(font.Bold == MsoTriState.msoTrue) },
                            { "italic", Json.Bool(font.Italic == MsoTriState.msoTrue) }
                        });
                    }
                }
                catch { f["font"] = "null"; }
            }
            return Json.Obj(f);
        }

        // ---- 7: 页面切换特效及选项 ----
        public static Dictionary<string, string> Transition(Presentation pres, string slideIndex)
        {
            var slide = SlideAt(pres, slideIndex);
            var t = slide.SlideShowTransition;
            var f = new Dictionary<string, string>
            {
                { "slideIndex", slide.SlideIndex.ToString() },
                { "entryEffect", Json.Str(t.EntryEffect.ToString()) },
                { "speed", Json.Str(t.Speed.ToString()) },
                { "duration", Json.Num(t.Duration) },
                { "advanceOnClick", Json.Bool(t.AdvanceOnClick == MsoTriState.msoTrue) },
                { "advanceOnTime", Json.Bool(t.AdvanceOnTime == MsoTriState.msoTrue) },
                { "advanceTime", Json.Num(t.AdvanceTime) },
                { "hidden", Json.Bool(t.Hidden == MsoTriState.msoTrue) },
                { "loopSoundUntilNext", Json.Bool(t.LoopSoundUntilNext == MsoTriState.msoTrue) }
            };
            return f;
        }

        // ---- 8/9: 动画时序表（主序列 + 交互序列），含目标对象/类型/时长/触发 ----
        public static Dictionary<string, string> Timeline(Presentation pres, string slideIndex)
        {
            var slide = SlideAt(pres, slideIndex);
            var tl = slide.TimeLine;
            var sequences = new List<string>();

            sequences.Add(SequenceJson(tl.MainSequence, "main"));
            try
            {
                foreach (Sequence seq in tl.InteractiveSequences) sequences.Add(SequenceJson(seq, "interactive"));
            }
            catch { }

            return new Dictionary<string, string>
            {
                { "slideIndex", slide.SlideIndex.ToString() },
                { "sequences", "[" + string.Join(",", sequences) + "]" }
            };
        }

        private static string SequenceJson(Sequence seq, string kind)
        {
            var effects = new List<string>();
            foreach (Effect e in seq) effects.Add(EffectJson(e));
            return Json.Obj(new Dictionary<string, string>
            {
                { "kind", Json.Str(kind) },
                { "count", seq.Count.ToString() },
                { "effects", "[" + string.Join(",", effects) + "]" }
            });
        }

        private static string EffectJson(Effect e)
        {
            var f = new Dictionary<string, string>
            {
                { "index", e.Index.ToString() },
                { "displayName", Json.Str(e.DisplayName) },
                { "effectType", Json.Str(e.EffectType.ToString()) },
                { "exit", Json.Bool(e.Exit) }
            };

            string target = "null";
            try { if (e.Shape != null) target = Json.Str(e.Shape.Name); } catch { }
            f["targetName"] = target;
            string targetId = "null";
            try { if (e.Shape != null) targetId = e.Shape.Id.ToString(); } catch { }
            f["targetId"] = targetId;

            try
            {
                var tm = e.Timing;
                f["timing"] = Json.Obj(new Dictionary<string, string>
                {
                    { "duration", Json.Num(tm.Duration) },
                    { "triggerType", Json.Str(tm.TriggerType.ToString()) },
                    { "triggerDelayTime", Json.Num(tm.TriggerDelayTime) },
                    { "repeatCount", Json.Num(tm.RepeatCount) },
                    { "repeatDuration", Json.Num(tm.RepeatDuration) },
                    { "speed", Json.Num(tm.Speed) },
                    { "autoReverse", Json.Bool(tm.AutoReverse == MsoTriState.msoTrue) },
                    { "rewindAtEnd", Json.Bool(tm.RewindAtEnd == MsoTriState.msoTrue) }
                });
            }
            catch { f["timing"] = "null"; }

            try
            {
                var ep = e.EffectParameters;
                f["parameters"] = Json.Obj(new Dictionary<string, string>
                {
                    { "direction", Json.Str(ep.Direction.ToString()) },
                    { "amount", Json.Num(ep.Amount) },
                    { "size", Json.Num(ep.Size) }
                });
            }
            catch { f["parameters"] = "null"; }

            return Json.Obj(f);
        }

        // ---- 10/11: 全量动画/切换特效值（COM 枚举反射，供外部做映射表）----
        public static string EnumNames(Type enumType)
        {
            var names = new List<string>();
            foreach (var n in Enum.GetNames(enumType)) names.Add(Json.Str(n));
            return "[" + string.Join(",", names) + "]";
        }
    }
}
