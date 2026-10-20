# 国六环保样品参数表核对（WorkBuddy Skill）

把「翻两份视同报告、逐字段填重型柴油车样品参数表」这件事，改成**先出核对表、人工确认、再写进新副本、最后自动核查**。

面向专用车环保申报场景：每台改装车都要填一份《重型柴油车样品参数表（国六环保）》，数据来自一份排放（PEMS）视同报告和一份噪声视同报告。两份报告字段重叠但值不一样，还有些字段报告里根本没有，人工填最容易错。

## 它解决什么

- **哪些字段该信哪份报告**：噪声报告和排放报告各有独占栏目（喷油泵、EGR、OECU 只在排放报告；排气消声器、排气管、膨胀室只在噪声报告），抄错一份就是空值。
- **同名字段值不同怎么办**：比如最大设计车速两份报告不同——规则里没写推导依据的一律不猜，标成冲突交给人。
- **哪些必须人工填**：整车参数、底盘参数一律留空列成清单，避免把上一台车的数据留在模板里。
- **填错了怎么发现**：写表后跑核查，回读比对 + 9 类一致性校验（质量、轮胎数、速比个数、功率、后处理、分类、车型、格式、空值）。
- **每个值到底抄的哪份文件**：输出一份《数据来源核对表》Word，逐字段标明来源文件（A 排放 / B 噪声）、报告栏目路径（如「发动机 › 最大净功率/转速」）、以及参数表里实际写入值是否与之一致。可直接打印签字。

## 目录结构

```
guo6-param-sheet-check/
├── SKILL.md                      # 主流程：七步
├── references/
│   ├── field-rules.tsv           # 67 个字段的取数规则（唯一规则源）
│   └── report-layout.md          # 报告与参数表的栏目位置、常见坑
└── scripts/
    ├── docx_table.py             # docx 表格读写（仅标准库）
    ├── docx_write.py             # 从零生成 docx（标题/段落/表格）
    ├── build_audit.py            # 生成核对表 audit.md / audit.json（含溯源字段）
    ├── fill_sheet.py             # 写进参数表副本（默认只预览）
    ├── verify_sheet.py           # 核查成品，出 verify.md
    └── make_source_doc.py        # 生成《数据来源核对表》Word
```

## 安装

复制到 WorkBuddy 的用户级技能目录：

```bash
git clone https://github.com/guoxiaobing020408/Skill-HuanBao.git
cp -r Skill-HuanBao/guo6-param-sheet-check ~/.workbuddy/skills/
```

Windows 上就是复制到 `C:\Users\<用户名>\.workbuddy\skills\guo6-param-sheet-check\`。

装完在对话里说「填国六环保参数表」「核一下这台车的样品参数表」即可触发。

## 更新仓库

改完技能后，双击本目录下的 `push.bat` 即可：它先把 `~\.workbuddy\skills\guo6-param-sheet-check\` 的文件同步到这里，再调用 `push_api.py` 通过 GitHub API 整体上传。

首次使用需要一次 Token：

1. 打开 https://github.com/settings/tokens → **Generate new token (classic)**
2. Note 随便填，Expiration 选 No expiration，勾选 **repo**
3. 生成后复制 `ghp_...`，存成文件 `C:\Users\<用户名>\.workbuddy\github_token.txt`（纯文本一行）

之后双击 `push.bat` 就静默上传，不用再输任何东西。

> 为什么不用 `git push`：部分网络环境无法访问 `github.com` 的 443 端口（`api.github.com` 却正常），git 协议会卡在 CONNECT。`push_api.py` 只走 `api.github.com`，不受影响，也不依赖本机 git。

## 使用

```bash
cd ~/.workbuddy/skills/guo6-param-sheet-check

# 1 出核对表
python scripts/build_audit.py \
  --sheet "<参数表模板>" --pm "<排放报告>" --zs "<噪声报告>" \
  --rules references/field-rules.tsv --out-dir "out/<车号>"

# 2 看核对表，确认后写进副本（不加 --apply 只预览）
python scripts/fill_sheet.py --sheet "<参数表模板>" \
  --audit "out/<车号>/audit.json" --out "<车号>_样品参数表_<日期>.docx" --apply

# 3 核查
python scripts/verify_sheet.py --sheet "<刚生成的副本>" --audit "out/<车号>/audit.json"

# 4 出数据来源核对表（Word，可直接打印签字）
python scripts/make_source_doc.py --audit "out/<车号>/audit.json" \
  --sheet "<刚生成的副本>" --vehicle "<车型>" --out "out/<车号>/数据来源核对表.docx"
```

补充参数：`--set "字段=值"` 当场补人工字段；`--clear-manual` 清掉模板里上一台车的残留值；`--force` 覆盖已存在的副本。原件始终只读。

## 规则怎么改

`references/field-rules.tsv` 七列：参数表分段、参数表字段、取数方式、报告分段、报告字段、取值处理、说明。

取数方式取 `自动` / `排放` / `噪声` / `人工` / `不填`；取值处理取 `value` / `pair`（型号/厂家拼接）/ `fuel_maker` / `fuel_spec`。

规则表里不含任何具体车型和具体数值，换车只换输入文件。

## 环境

Python 3.10+，只用标准库，无第三方依赖。输入文件为 `.docx`，或扩展名是 `.doc` 但实际是 docx 包；真正的 Word 97 二进制 `.doc` 读不出表格，需先另存为 `.docx`。

## 已知限制

- 只适配「重型柴油车样品参数表（国六环保）」这一版式，参数表改版需同步改规则表字段名。
- 报告版式差异过大（栏目改名、参数不在「检验车辆基本参数」表内）时，需改规则表的报告字段名。
- 城市/非城市分类冲突、两份报告发动机型号不同，属业务判断，脚本只提示不解决。
- 未覆盖：公告参数自动抓取、多车型批量处理。

## 验证情况

- 用例一：WRD5120TXC，42 个自动字段与已人工填好的标准答案逐字一致。
- 用例二：WRD5041THX，写入 43 个单元格后回读 0 差异；核查 0 错误、警告全部指向模板残留的人工字段。
- 用例三：用构造的两份报告（含两份取值不同的 DOC、变速箱型号）跑完整链路，来源核对表 6 张表结构正确，单来源字段（喷油泵→A、排气管→B、速比→B）与双来源一致字段标注均正确，冲突字段按 A/B 两列并排。

## 许可

未声明许可证。使用前请与作者确认。
