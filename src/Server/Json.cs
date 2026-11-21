using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace BetterWPP.Server
{
    /// <summary>极简 JSON 写入器（不依赖第三方库）。支持对象/数组/字符串/数字/布尔/null。</summary>
    internal static class Json
    {
        public static string Str(string s)
        {
            if (s == null) return "null";
            var sb = new StringBuilder(s.Length + 8);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u" + ((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        public static string Num(double n)
        {
            return n.ToString("0.######", CultureInfo.InvariantCulture);
        }

        public static string Bool(bool b) { return b ? "true" : "false"; }

        public static string Obj(Dictionary<string, string> fields)
        {
            var sb = new StringBuilder("{");
            bool first = true;
            foreach (var kv in fields)
            {
                if (!first) sb.Append(',');
                first = false;
                sb.Append(Str(kv.Key)).Append(':').Append(kv.Value ?? "null");
            }
            sb.Append('}');
            return sb.ToString();
        }
    }
}
