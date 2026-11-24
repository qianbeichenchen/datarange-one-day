# Demo 开发与 GitHub 同步约定

## 持久用户偏好

- 每个 Demo 使用独立的 Git 仓库和独立的 GitHub 仓库。
- GitHub 所有者为 `qianbeichenchen`：https://github.com/qianbeichenchen 。不要将本目录整体初始化为一个仓库。
- 新建 Demo 时，在其目录创建 Git 仓库，默认分支为 `main`；仓库名默认使用 Demo 目录名。
- 默认创建 **private** 仓库；用户明确要求公开时才改用 public。
- 用户已授权：新 Demo 完成并验证后创建对应 GitHub 仓库、提交并推送；此后每轮修改完成并验证后自动提交并推送，无需重复询问。不是每次文件保存都提交。
- 创建新的 Demo 时，在其根目录写入独立的 `AGENTS.md`，保留本文件的 GitHub 同步约定，以便直接从该 Demo 目录启动的新会话仍能读取规则。
- 此约定适用于后续新建和修改的 Demo；不要仅因读取此文件就批量上传历史 Demo。

## 提交与推送流程

1. 先检查当前 Demo 的 Git 状态、远程地址、当前分支和已有说明。已有远程地址不得直接覆盖；确认目标属于用户期望的项目。
2. 首次同步前创建或补充 `.gitignore`，排除 `node_modules/`、缓存、临时文件、`.env`、私钥及其他凭据。检查许可证密钥、内部客户资料、专有库和不允许分发的文件，避免提交这些内容。
3. 进行与本次改动相称的验证。验证失败时先修复；如因环境限制无法验证，应在结果中明确说明，不能声称验证通过。
4. 仅暂存本轮任务相关的文件，保留用户其他修改；写明改动内容后 commit。没有新改动时不创建空提交。
5. 使用 GitHub CLI 或其他已授权的工具创建仓库并绑定 `origin`，推送当前分支。若同名远程仓库已存在，先核对归属和内容，不覆盖远程历史。
6. 禁止默认 force push、重写远程历史或为处理冲突而丢弃改动。遇到认证失败、权限问题或冲突时保留本地成果，说明阻塞原因。
7. 最终回复给出 Demo 路径、GitHub 仓库链接及验证和推送结果；未成功推送时明确标注。

## 初始环境检查

- 2026-11-24：已安装 Git，Git 提交身份已配置；便携版 GitHub CLI 位于 `C:\Users\admin\Documents\ChatGPT\SpreadJS Demo\.tools\github-cli\bin\gh.exe`，不依赖全局 PATH。
- 已完成 GitHub 授权，通过 `gh auth status` 和 `gh api user` 核实当前账户为 `qianbeichenchen`，使用 HTTPS，凭据保存在系统 keyring，授权包含 `repo`。尚未以实际建库或 push 操作测试写入。
- 本环境的受限沙箱会阻止 GitHub 网络连接；出现 socket access forbidden 等错误时，按工具权限流程申请网络执行权限后重试，不要据此误判凭据失效，也不要重新登录或输出 token。
- `.tools/` 为本地工具目录，不应提交到任何 Demo 仓库。
- 配置身份不等于 GitHub 已登录。首次使用前必须核实认证账户与目标所有者。
