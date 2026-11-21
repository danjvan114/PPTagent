using System;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.PowerPointApi.Tools;

namespace BetterWPP
{
    // 上架前用 guidgen / VS 工具生成新 GUID，并同步 register.cmd 中的 CLSID
    [ComVisible(true)]
    [Guid("11111111-1111-1111-1111-111111111111")]
    [ProgId("BetterWPP.Connect")]
    [COMAddin("BetterWPP", "BetterWPP PowerPoint COM add-in", LoadBehavior.LoadAtStartup)]
    [CustomUI("", "BetterWPP.src.RibbonUI.xml")]
    public class Connect : COMAddin
    {
        public Connect()
        {
            OnConnection += Connect_OnConnection;
            OnDisconnection += Connect_OnDisconnection;
        }

        private void Connect_OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            // this.Application 为 NetOffice 包装的 PowerPoint.Application
        }

        private void Connect_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
        }

        // Ribbon 按钮回调；签名需为 void 且接收 IRibbonControl 参数
        public void OnHelloButton(object control)
        {
            System.Windows.Forms.MessageBox.Show("BetterWPP 加载成功！", "BetterWPP");
        }

        // 兜底：若 CustomUI 特性未取到内嵌资源，重写此方法返回 Ribbon XML
        protected override string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("RibbonUI.xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new System.IO.StreamReader(asm.GetManifestResourceStream(name));
                    return reader.ReadToEnd();
                }
            }
            return base.GetCustomUI(ribbonID);
        }
    }
}
