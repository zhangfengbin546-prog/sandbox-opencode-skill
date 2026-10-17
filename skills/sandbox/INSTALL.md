# sandbox Skill 安装说明（OpenCode）

## 环境要求

- OpenCode 2.x
- Python 3.10 或更高版本
- 推荐模型：`opencode-go/deepseek-v4.1-flash`
- 读取旧版 `.doc` 规则文件时，Windows 还需要 Microsoft Word；`.docx` 不需要 Word
- 不需要 Microsoft Excel，因为本 Skill 不打开或修改工作簿

## 全局安装

解压后确认 `sandbox` 文件夹中直接包含 `SKILL.md`，再把整个文件夹复制到：

```text
~/.config/opencode/skills/sandbox
```

Windows 对应路径通常是：

```text
C:\Users\你的用户名\.config\opencode\skills\sandbox
```

也可以安装到单个项目：

```text
<项目目录>/.opencode/skills/sandbox
```

安装 Python 依赖：

```powershell
py -m pip install -r "$env:USERPROFILE\.config\opencode\skills\sandbox\requirements.txt"
```

macOS 或 Linux：

```bash
python3 -m pip install -r ~/.config/opencode/skills/sandbox/requirements.txt
```

## OpenCode 配置

不要直接覆盖已有 `opencode.json`。参考 `opencode.example.json`，把 `model`、`agent.build.reasoningEffort` 和所需权限合并进全局或项目配置。模型也可以在 OpenCode 中运行 `/models` 后选择 `opencode-go/deepseek-v4.1-flash`，并把推理强度切换为 `low`。OpenCode 默认可用 `Ctrl+T` 轮换模型变体；如果你的键位配置不同，请查看 `variant_cycle`。

重新启动 OpenCode或新建会话后，输入 `/sandbox` 调用。运行时会使用内置 `question` 工具依次弹出角色、打法和首年补单三个单选问题。

## 更换方案表

删除 `assets/` 中原来的表格，放入自己的表格即可，不需要修改 `SKILL.md` 或脚本。支持 `.xlsx`、`.xlsm`、`.xls`、`.xlsb`、`.ods` 和 `.csv`。

如果调用时另行上传了表格，上传表格优先于 `assets/`。方案末尾会先显示“根据规则制定的方案表已经生成，并放置到桌面，请您查收”，下一行显示该桌面文件的超链接，不显示裸路径或内部复制信息。

方案表会直接复制到操作系统中当前用户配置的真实桌面目录。Windows 会读取系统桌面位置，因此也支持桌面被重定向到其他盘符或 OneDrive。规则抽取文件和战略草稿使用系统临时目录，不会出现在桌面。若桌面已有同名文件，新文件会自动追加数字序号，不覆盖旧文件。

## 速度说明

完成三个弹窗问题后，Skill 只运行一次本地总流程，工具结果会直接带回最终方案。方案正文由本地脚本快速组织为约 6000 至 9000 个中文字符，DeepSeek 不再做第二轮推理或扩写。推荐使用 `low` 推理强度；正常本地文件条件下以 3 分钟内完成为设计目标，但模型服务排队、网络延迟和用户回答弹窗所用时间不受 Skill 控制。
