using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.Tools;
using NetOffice.PowerPointApi.Tools;
using BetterWPP.Server;

namespace BetterWPP
{
    // 上架前用 guidgen / VS 工具生成新 GUID，并同步 register.cmd 中的 CLSID
    [ComVisible(true)]
    [Guid("11111111-1111-1111-1111-111111111111")]
    [ProgId("BetterWPP.Connect")]
    [COMAddin("BetterWPP", "BetterWPP PowerPoint COM add-in", LoadBehavior.LoadAtStartup)]
    [CustomUI("RibbonUI.xml", true)]
    public class Connect : COMAddin
    {
        private HttpServer _server;

        public Connect()
        {
            OnConnection += Connect_OnConnection;
            OnDisconnection += Connect_OnDisconnection;
        }

        private void Connect_OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            // this.Application 为 NetOffice 包装的 PowerPoint.Application
            if (this.Application != null)
            {
                _server = new HttpServer(this.Application);
            }
        }

        private void Connect_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (_server != null) { _server.Dispose(); _server = null; }
        }

        // ---- Ribbon 回调：启动服务 ----
        public void OnStartService(object control)
        {
            if (_server == null) return;
            if (_server.IsRunning) { MessageBox.Show("服务已在运行", "BetterWPP"); return; }
            bool ok = _server.Start();
            if (ok) MessageBox.Show("服务已启动\n" + string.Join("\n", _server.GetUrls()), "BetterWPP");
            else MessageBox.Show("启动失败：" + _server.LastError, "BetterWPP", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ---- Ribbon 回调：停止服务 ----
        public void OnStopService(object control)
        {
            if (_server == null) return;
            _server.Stop();
            MessageBox.Show("服务已停止", "BetterWPP");
        }

        // ---- Ribbon 回调：连接码（显示地址）----
        public void OnShowCode(object control)
        {
            if (_server == null) return;
            if (!_server.IsRunning) { MessageBox.Show("服务未启动", "BetterWPP"); return; }
            MessageBox.Show(string.Join("\n", _server.GetUrls()), "连接码");
        }

        // ---- Ribbon 回调：网络设置（本地 / 局域网）----
        public void OnNetworkSetting(object control)
        {
            if (_server == null) return;
            var mode = _server.Mode;
            var dlg = new Form
            {
                Text = "网络设置",
                Width = 320,
                Height = 180,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen
            };
            var rbLocal = new RadioButton { Text = "仅本机（127.0.0.1）", Left = 20, Top = 20, Width = 260, Checked = mode == BindMode.Loopback };
            var rbLan = new RadioButton { Text = "局域网（0.0.0.0，同网段可连）", Left = 20, Top = 50, Width = 260, Checked = mode == BindMode.Lan };
            var btnOk = new Button { Text = "确定", Left = 120, Top = 95, Width = 80, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "取消", Left = 210, Top = 95, Width = 80, DialogResult = DialogResult.Cancel };
            dlg.Controls.AddRange(new Control[] { rbLocal, rbLan, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _server.SetMode(rbLan.Checked ? BindMode.Lan : BindMode.Loopback);
                if (_server.IsRunning) MessageBox.Show("已切换，重启服务后生效", "BetterWPP");
            }
        }

        // 兜底：若 CustomUI 特性未取到内嵌资源，重写此方法返回 Ribbon XML
        public override string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("RibbonUI.xml", StringComparison.OrdinalIgnoreCase))
                {
                    var stream = asm.GetManifestResourceStream(name);
                    try
                    {
                        var reader = new System.IO.StreamReader(stream);
                        return reader.ReadToEnd();
                    }
                    finally { stream.Dispose(); }
                }
            }
            return base.GetCustomUI(ribbonID);
        }
    }
}
