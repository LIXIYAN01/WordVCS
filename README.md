# WordVCS — Word 论文版本控制插件

WordVCS 是 **Word COM Add-in 插件**，将 Git 版本控制深度融入 Word 的批注工作流。

打开 Word 即自动加载，在「开始」旁显示 **「论文版本」** 标签页，侧边栏面板嵌入 Word 窗口内。

---

## 核心功能

| 功能 | 说明 |
|------|------|
| **Ribbon 集成** | Word 顶部自定义标签页「论文版本」，包含提交/分支/标签/反馈按钮 |
| **Git 版本控制** | 基于 libgit2sharp 的完整 Git 操作 (init/commit/branch/tag/restore) |
| **批注 ↔ 提交映射** | 提交时勾选已解决的批注，SQLite 本地数据库永久关联 |
| **文本差异对比** | 红/绿/黄着色显示 .docx 版本间文字变化 |
| **导师反馈接收** | 一键导入导师批注文件，自动创建 feedback 分支和标签 |
| **纯本地运行** | 零云端依赖，数据存储于 `.git/` 和 `.wordvcs/` 目录 |

---

## 安装

### 1. 编译

```powershell
cd WordVCS
dotnet restore WordVCS.sln
# 或使用 MSBuild
msbuild WordVCS.sln /p:Configuration=Release
```

### 2. 注册插件

以**管理员身份**运行：
```bat
cd src\WordVCS.AddIn\bin\Release\net48
register.bat
```

### 3. 使用

1. 打开 Word，查看「开始」旁边的 **「论文版本」** 标签页
2. 点击 **「显示面板」** 打开侧边栏
3. 在侧边栏中提交新版本、关联批注、管理分支、导入导师反馈

卸载：运行 `unregister.bat`

---

## 项目结构

```
WordVCS/
├── WordVCS.sln
├── src/
│   ├── WordVCS.Core/          # 核心业务服务 (Git/Diff/Comment/SQLite 映射)
│   ├── WordVCS.UI/            # WPF 侧边栏界面
│   └── WordVCS.AddIn/         # COM Add-in (Ribbon + 任务窗格 + 注册脚本)
└── scripts/
```

## 技术栈

| 层级 | 技术 |
|------|------|
| 宿主 | COM Add-in (IDTExtensibility2 + IRibbonExtensibility) |
| UI | WPF (Windows Presentation Foundation) |
| Git | LibGit2Sharp 0.30 |
| 文档解析 | Open XML SDK 3.1 |
| 差异算法 | DiffPlex |
| 本地存储 | SQLite |
| 目标框架 | .NET Framework 4.8 |

## 许可证

Apache-2.0
