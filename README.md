# WordVCS — Word 论文版本控制系统 VSTO 插件

WordVCS 是一款为学术论文写作而设计的 **Microsoft Word VSTO Add-in 插件**。它将 Git 版本控制与 Word 批注系统深度融合，实现每次修改可追溯、每条批注有回应、每个版本有上下文。

---

## 核心功能

| 功能 | 说明 |
|------|------|
| **Git 版本控制** | 在 Word 侧边栏直接管理提交/分支/标签，完整 Git 操作 |
| **批注 ↔ 提交映射** | 提交时勾选解决的批注，自动在 SQLite 中建立关联 |
| **文本差异对比** | 红/绿/黄高亮显示版本间的文字差异 |
| **导师反馈接收** | 一键导入导师批注文件，自动创建 feedback 分支 |
| **纯本地运行** | 零云端依赖，数据存于本地 `.wordvcs/` 目录 |

---

## 快速开始

### 1. 环境要求

- Windows 10/11
- Visual Studio 2022（Community 免费版本即可）
  - 安装时勾选 **"Office/SharePoint 开发"** 负载
- Microsoft Word 2016/2019/2021 或 Microsoft 365 桌面版
- .NET Framework 4.8 SDK

### 2. 编译运行

```powershell
# 克隆或解压项目后
cd WordVCS

# 还原依赖
dotnet restore WordVCS.sln

# 在 Visual Studio 中打开
start WordVCS.sln

# 按 F5 启动调试
```

### 3. 使用

1. 在 Word 中打开论文 `.docx` 文件
2. 点击顶部 Ribbon 的 **「论文版本」** 标签页
3. 点击 **「打开侧边栏」** 启动版本控制面板
4. 在侧边栏中进行提交、管理批注、查看差异

---

## 项目结构

```
WordVCS/
├── WordVCS.sln                     # VS 解决方案
├── src/
│   ├── WordVCS.Core/               # 核心业务逻辑（无 UI 依赖）
│   │   ├── Models/Models.cs        # 数据模型
│   │   └── Services/               # Git/Diff/Comment/Mapping 服务
│   ├── WordVCS.UI/                 # WPF 侧边栏界面
│   │   ├── Controls/               # 时间线/批注面板/差异查看器/弹窗
│   │   ├── ViewModels/             # MVVM 数据绑定
│   │   ├── Converters/             # 值转换器
│   │   └── Themes/Styles.xaml      # UI 主题
│   └── WordVCS.AddIn/              # VSTO 插件入口
│       ├── ThisAddIn.cs            # 生命周期管理
│       ├── TaskPaneManager.cs      # UI ↔ Core 桥梁
│       ├── WordCommentService.cs   # Word COM 批注提取
│       └── Ribbon.xml/cs           # Ribbon 菜单
├── scripts/
│   ├── build.ps1                   # 构建脚本
│   └── setup-dev.ps1               # 环境检查
└── installer/
    └── WordVCS.wxs                 # WiX 安装器 (计划中)
```

---

## 技术栈

| 组件 | 技术 |
|------|------|
| 宿主 | VSTO (Visual Studio Tools for Office) |
| 语言 | C# 10.0 / .NET Framework 4.8 |
| UI | WPF (Windows Presentation Foundation) |
| Git | LibGit2Sharp |
| 文档解析 | Open XML SDK |
| 差异算法 | DiffPlex |
| 本地存储 | SQLite (System.Data.SQLite) |

---

## 许可证

Apache-2.0 License
