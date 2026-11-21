using System;
using System.Runtime.InteropServices;
using Extensibility;
using NetOffice.PowerPointApi;

namespace BetterWPP
{
    // 上架前用 guidgen / VS 工具生成新 GUID，并同步 register.cmd 中的 CLSID
    [ComVisible(true)]
    [Guid("11111111-1111-1111-1111-111111111111")]
    [ProgId("BetterWPP.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        private Application _app;

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            _app = Factory.GetApplication<Application>(application);
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app?.Dispose();
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }
    }
}
