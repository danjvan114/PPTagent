# BetterWPP

PowerPoint 桌面版 **COM 加载项**（.NET Framework 4.8，C#）。非 VSTO、非 Web Add-in。

## 技术栈
- .NET Framework 4.8（COM/VSTO 平台不支持 .NET Core / .NET 5+）
- NetOfficeFw.PowerPoint（NuGet 互操作包装，编译不依赖本机安装的 PIA）
- `IDTExtensibility2` 入口，注册表方式挂载到 PowerPoint
- GitHub Actions：`windows-latest` + MSBuild，公开仓库免费

## 目录
```
agent/
├─ BetterWPP.sln
├─ .github/workflows/build.yml   # CI：编译并上传 artifact
└─ src/
   ├─ BetterWPP.csproj
   ├─ Connect.cs                 # 加载项入口
   └─ register.cmd               # 本机注册脚本
```

## 本地开发
前置：Windows、PowerPoint 桌面版、VS Build Tools 2022（.NET 桌面开发工作负载）。

```powershell
msbuild BetterWPP.sln -p:Configuration=Release -p:Platform=x64
```

编译后以管理员身份运行 `src\register.cmd`，重启 PowerPoint。

## CI 说明
- 推送到 `main/master`、提 PR、或 Actions 页手动触发都会跑。
- runner 只负责编译出 DLL；**不安装 PowerPoint**，无法在 CI 内启动 PPT 验证加载效果，调试必须回本机。
- artifact 含加载项 DLL、NetOffice 依赖、pdb、注册脚本。

## 上架前必改
1. `Connect.cs` 中的占位 GUID 换成新生成的，同步 `register.cmd` 的 `CLSID`。
2. 需要强名称签名时，把 `.pfx` 放进仓库 Secrets，开启 csproj 的 `SignAssembly`。
3. 目标机需安装 .NET Framework 4.8 和 PowerPoint 桌面版。
