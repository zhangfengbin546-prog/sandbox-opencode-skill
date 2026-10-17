# sandbox — 企业沙盘模拟大赛战略方案技能

一个给 **OpenCode** 使用的 Skill。加载后会依次弹出三个单选问题（参赛角色、打法、首年补单选择），随后在本地一次性生成约 6000–9000 字的完整战略方案，并把方案表复制到当前用户的桌面、附带超链接。

## 仓库结构

```
skills/sandbox/
├── SKILL.md                 技能主文件（frontmatter 的 name 必须是 sandbox）
├── INSTALL.md               详细安装说明
├── requirements.txt         Python 依赖
├── opencode.example.json    OpenCode 配置参考
├── agents/openai.yaml
├── assets/                  方案表模板
├── references/              策略要点
└── scripts/                 本地计算脚本
```

## 环境要求

- OpenCode 2.x
- Python 3.10 或更高版本
- 推荐模型：`opencode-go/deepseek-v4.1-flash`
- 读取旧版 `.doc` 规则文件时需要 Microsoft Word（`.docx` 不需要）
- 不需要 Microsoft Excel

## 安装

**方式一 · 用 skills CLI（推荐，一条命令）**

```bash
npx skills add zhangfengbin546-prog/sandbox-opencode-skill --skill sandbox -g --all
```

**方式二 · 手动复制**

把 `skills/sandbox` 整个文件夹复制到 OpenCode 的技能目录：

```text
~/.config/opencode/skills/sandbox
```

Windows 上通常是 `C:\Users\你的用户名\.config\opencode\skills\sandbox`。

再安装 Python 依赖：

```powershell
py -m pip install -r "$env:USERPROFILE\.config\opencode\skills\sandbox\requirements.txt"
```

重启 OpenCode 或新建会话，输入 `/sandbox` 调用。

## 更换方案表

删掉 `assets/` 里原来的表格，放进自己的表格即可，**不需要**修改 `SKILL.md` 或脚本。支持 `.xlsx`、`.xlsm`、`.xls`、`.xlsb`、`.ods`、`.csv`。

## 更新已安装的版本

```bash
npx skills update
```

更多细节见 `skills/sandbox/INSTALL.md`。
