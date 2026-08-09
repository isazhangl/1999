# Lovart 接入说明（1999 项目）

记录 Lovart AI 在本项目中**实际可用**的调用方式。

## 结论：走通的方案 = 官方 Skill（命令行）

通过官方 skill 调用,**不直接 curl API**(SKILL.md 明确禁止直接构造 API 请求)。

### 安装过程 & 一个坑

```bash
npx -y skills add lovartai/lovart-skill --skill lovart-api --agent claude-code
```

⚠️ **注意**:该命令只复制了文档(`SKILL.md` 等),**漏了 `scripts/` 目录**(核心脚本 `agent_skill.py`)。
已从源仓库 `lovartai/lovart-skill` 的 `skills/lovart-skill/scripts/` 手动补齐,现位于:

```
.claude/skills/lovart-api/
├── SKILL.md
├── README*.md
└── scripts/agent_skill.py   # ← 手动补齐,核心脚本(纯 Python 标准库,无第三方依赖)
```

> 若日后 `skills` 工具升级修复了此问题,可重装并对比 `scripts/` 是否一致。

## 密钥配置（不写死在代码里）

脚本从**环境变量**读取凭证(见 `agent_skill.py` 中 `os.environ.get("LOVART_ACCESS_KEY")`):

- `LOVART_ACCESS_KEY`
- `LOVART_SECRET_KEY`

密钥存放在项目根目录 **`.env`**(已被 `.gitignore` 忽略,**不进 git**):

```dotenv
LOVART_ACCESS_KEY=你的_ak
LOVART_SECRET_KEY=你的_sk
```

**重要**:Python 脚本不会自动读取 `.env`,需要先把变量导入当前 shell 再调用命令:

```bash
set -a; source .env; set +a          # 把 .env 导入环境
python3 .claude/skills/lovart-api/scripts/agent_skill.py projects --json
```

（也可以在调用时临时传入:`--ak ... --sk ...`,但不推荐,会进命令历史。）

## 常用命令

脚本路径统一为 `python3 .claude/skills/lovart-api/scripts/agent_skill.py <命令>`。

| 用途 | 命令 |
|---|---|
| 查看本地状态/当前项目 | `config --json` |
| 列出项目 | `projects --json` |
| 新建并激活项目 | `create-project` 或 `project-add --project-id ID --name "名称"` |
| 切换项目 | `project-switch --project-id ID` |
| 列出对话线程 | `threads --json` |
| **生成图片（主力，阻塞到完成）** | `chat --prompt "描述" --json --download` |
| 流式多图 | `watch --prompt "生成4张变体" --json` |
| 继续同一对话 | `chat --thread-id ID --prompt "..." --json --download` |
| 上传本地图做参考 | `upload --file /path/img.png` |
| 高成本操作(视频等)确认 | `confirm --thread-id ID --json --download` |
| 放大/超分 | `chat --prompt "upscale" --include-tools upscale_image --attachments URL --json --download` |
| 深度规划模式(品牌系统等) | `chat --prompt "..." --mode thinking --json --download` |
| 快速/无限模式切换 | `set-mode --fast` / `set-mode --unlimited` |

### 首次生成前必做（SKILL.md 规则）
1. `config --json` —— 检查有没有 `active_project`,没有就 `create-project` 或 `project-add`。
2. `threads --json` —— 有相关近期线程就用 `--thread-id` 复用,新话题才新建。

## 认证 & 网络

- 认证方式:**AK/SK HMAC-SHA256 签名**(脚本内部完成)。
- 本环境走代理:若遇到 TLS 校验失败,脚本支持 `LOVART_INSECURE_SSL=1` opt-out(仅在必要时用),且内置 3 次重试 + SSL fallback。优先不开,先试正常调用。

## 验证状态

- ✅ 脚本已就位,CLI 全部子命令可列出;`config`/`projects` 无 key 时优雅返回本地状态。
- ⏳ **端到端生成测试待完成** —— 需满足:① `.env` 填入**新**密钥(见下）;② **重开 session**(skill 不热加载)。之后跑一次最小生成(如 `chat --prompt "a simple round brilliant diamond on white background" --json --download`)验证整条链路。

## ⚠️ 安全提醒

之前在聊天中粘贴过的一对 AK/SK 已视为**泄露**,请到 Lovart 后台**作废并重新生成**,把**新密钥**填进 `.env`。已扫描确认:泄露的密钥字符串**未进入本仓库的工作区或 git 历史**。
