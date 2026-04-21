# 通用 MSIX/UWP 打包工具 v1.0

将任意便携版 Windows 程序一键打包为 MSIX（UWP）格式。

---

## 📁 目录结构

```
msix_toolkit\
├── make_msix.ps1       ← 主脚本（一键运行）
├── config.json         ← 配置文件（每次只改这里）
├── tools\              ← makeappx.exe / signtool.exe（已内置）
│   ├── makeappx.exe
│   ├── signtool.exe
│   └── *.dll
└── README.md           ← 本文件
```

---

## 🚀 快速开始

### 第一步：编辑 config.json

```json
{
  "app": {
    "name":        "MyApp",           // 包名（无空格）
    "displayName": "My Application",  // 显示名称
    "executable":  "MyApp\\myapp.exe",// 相对于包根目录的 exe 路径
    "version":     "1.0.0.0"          // 四段版本号
  },
  "icon": {
    "sourceExe": "C:\\Program Files\\MyApp\\myapp.exe",  // 从 exe 提取图标
    "sourcePng": "",                                      // 或指定 PNG 文件
    "bgColor":   "auto"                                   // auto=自动提取
  }
}
```

### 第二步：运行脚本

```powershell
# 右键 → 用 PowerShell 运行，或：
powershell -ExecutionPolicy Bypass -File make_msix.ps1
```

### 第三步：完成！

脚本自动完成：
1. ✅ 生成自签名证书（有效期 10 年）
2. ✅ 提取 exe 图标颜色，生成 UWP 图标（4 种尺寸）
3. ✅ 生成 AppxManifest.xml（含文件类型关联）
4. ✅ 打包 → 签名 → 安装

---

## 📦 在其他电脑安装

每次打包会生成两个文件：
- `AppName.msix` — 安装包
- `AppName.cer` — 证书（公钥）

**在目标电脑上：**
1. 双击 `AppName.cer` → 安装到「受信任的根证书颁发机构」
2. 双击 `AppName.msix` → 安装应用

> 💡 如果目标电脑已安装过同一证书，跳过步骤 1

---

## ⚙️ config.json 完整说明

| 字段 | 说明 | 示例 |
|------|------|------|
| `app.name` | 包名（无空格、无特殊字符） | `"SublimeText"` |
| `app.displayName` | 显示名称 | `"Sublime Text"` |
| `app.description` | 描述 | `"Text editor"` |
| `app.publisher` | 证书主题（需与 cert.subject 一致） | `"CN=SelfSigned"` |
| `app.publisherDisplay` | 发布者显示名 | `"My Company"` |
| `app.version` | 四段版本号 | `"4200.0.0.0"` |
| `app.executable` | exe 相对路径 | `"MyApp\\myapp.exe"` |
| `app.arch` | 架构 | `"x64"` 或 `"x86"` |
| `icon.sourceExe` | 从此 exe 提取图标 | `"C:\\...\\myapp.exe"` |
| `icon.sourcePng` | 或使用此 PNG（优先级高于 exe） | `"C:\\...\\icon.png"` |
| `icon.bgColor` | 图标背景色 | `"auto"` 或 `"#1E1E1E"` |
| `icon.logoColor` | 图标主色（仅参考，不影响生成） | `"auto"` |
| `cert.subject` | 证书主题 | `"CN=SelfSigned"` |
| `cert.password` | PFX 密码 | `"YourPassword123"` |
| `cert.validYears` | 有效年数 | `10` |
| `tile.backgroundColor` | 开始菜单磁贴背景色 | `"#1E1E1E"` |
| `fileTypes.enabled` | 是否注册文件类型关联 | `true` |
| `fileTypes.groups` | 文件类型分组 | 见 config.json |
| `output.dir` | 输出目录（空=脚本目录） | `"C:\\output"` |
| `output.autoInstall` | 打包后自动安装 | `true` |
| `output.keepPackageDir` | 保留临时打包目录 | `false` |

---

## 🔧 常见问题

**Q: 提示"无法运行 makeappx.exe"**
> 工具目录已内置，无需额外安装。如果仍报错，将 `tools\` 目录复制到 `C:\DAIKIN\` 等受信任路径。

**Q: 某些文件类型无法注册（如 .bat .cmd）**
> Windows 保护这些扩展名，脚本会自动跳过并继续打包。

**Q: 安装时提示"证书不受信任"**
> 先双击 `.cer` 文件安装证书，选择「受信任的根证书颁发机构」。

**Q: 如何打包 32 位程序？**
> 将 `app.arch` 改为 `"x86"`。

**Q: 应用文件放在哪里？**
> 脚本自动从 `C:\Program Files\<appSubDir>` 复制。
> 如果不在 Program Files，将应用文件夹放到 `msix_toolkit\app\` 目录下。

---

## 📋 已验证的应用

| 应用 | 版本 | 状态 |
|------|------|------|
| Sublime Text | 4200 | ✅ |

---

*工具版本: v1.0 | 2026-04-21*
