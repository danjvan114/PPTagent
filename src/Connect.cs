using System;
using System.Runtime.InteropServices;
using NetOffice.PowerPointApi.Tools;

namespace BetterWPP
{
    // 上架前用 guidgen / VS 工具生成新 GUID，并同步 register.cmd 中的 CLSID
    [ComVisible(true)]
    [Guid("11111111-1111-1111-1111-111111111111")]
    [ProgId("BetterWPP.Connect")]
    [COMAddin("BetterWPP", "BetterWPP PowerPoint COM add-in", LoadBehavior.LoadAtStartup)]
    public class Connect : COMAddin
    {
        // 基类已实现 IDTExtensibility2；订阅事件在构造函数里挂接
        public Connect()
        {
            OnConnection += Connect_OnConnection;
            OnDisconnection += Connect_OnDisconnection;
        }

        private void Connect_OnConnection(object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            // this.Application 为 NetOffice 包装的 PowerPoint.Application，可在此挂事件、取 Ribbon
        }

        private void Connect_OnDisconnection(Extensibility.ext_DisconnectMode removeMode, ref Array custom)
        {
        }
    }
}
