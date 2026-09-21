# PDF 原格式翻译（pdft）

把外文 PDF（论文 / 讲义 / 教材，含大量公式）翻译成中文，**保持原版式**，输出 Word + PDF + HTML。
公式不重新排版——直接按原页面高清裁图嵌入，因此公式永远和原文像素一致、零出错。

![示例](docs/sample.png)

## 特点

- **公式保真**：不重建公式，按字形裁图，`cases` 大括号、分数、根号、向量箭头都完整
- **版式还原**：色块（tcolorbox）、标题、项目符号列表、页眉页脚、页码域
- **免密钥大模型**：默认走 WorkBuddy 云服务内置通道，用户不用填 API Key
- **免费额度内很快**：默认模型 `deepseek-v4-flash`，实测 9 批 × 16 条并发只要 **5 秒**
  （对比 `default` 同样负载 38–129 秒）
- **译文缓存**：按「原文 + 公式图指纹」缓存，同一份 PDF 重跑秒回
- **三格式输出**：DOCX（Word 原生列表 + 域页码）、PDF（Chrome 打印，文字可搜索）、HTML

## 快速开始

```bash
pip install -r requirements.txt

# 本地起服务，浏览器打开 http://127.0.0.1:8765
python app.py

# 或者命令行直接翻译
python app.py input.pdf --engine builtin --model deepseek-v4-flash \
    --formats docx,pdf,html -o out/译本
```

命令行参数：

| 参数 | 说明 |
|---|---|
| `--engine` | `builtin`（免密钥云通道）/ `custom`（自备 OpenAI 兼容接口）/ `none`（只还原版式不翻译） |
| `--model` | 内置通道模型 id，见 `/api/models`；推荐 `deepseek-v4-flash` |
| `--formats` | `docx,pdf,html` 任意组合 |
| `--mode` | `faithful`（公式成图，默认）/ `editable`（尝试重建可编辑公式） |
| `--src/--tgt` | 源语言 / 目标语言，默认 `法语` → `简体中文` |

## 部署

```bash
export WB_ENDPOINT=https://your-app.example.com      # 可选
export WB_PUBLISHABLE_KEY=wbpk_xxx                   # 建议用环境变量，别写进代码
python app.py                                        # 监听 8000 端口（容器里）
```

> ⚠️ 如果要**公开**这个仓库，请先把 `app.py` 里 `BUILTIN['publishable_key']` 的默认值清空，
> 改用环境变量注入——否则任何人都能用你的账号额度调用云通道。

## 实现要点

翻译质量的关键全在「怎么把 PDF 拆成可译文本 + 公式图」，这部分踩坑最多，记在
`pdftrans/extract.py` 与 `~/.workbuddy/skills/pdf-format-translation/SKILL.md`：

1. **视觉基线带合并**：PDF 会把一个分数拆成「分子行 / 分母行 / 根号行」多个伪行，
   先并回真实行，否则公式会被切碎
2. **续行并入**：`cases` 的行、分数的分子分母并回所属行（纵向重叠即可判定）
3. **成片挖块**：连续纯数学行整体挖成一块居中大图
4. **行内数学区整块裁切**：一行内连续的公式区域只出一张图（早期每行十几个小图，
   模型会把占位符放错位置，译文出现乱序残片）
5. **邻行字符遮挡**：PDF 行的包围盒纵向互相咬合，矩形裁图必然把邻行字符切进公式图。
   裁图后按字形包围盒把**不属于本图**的字符涂白，同时把公式自身区域从掩码里挖掉
   （实测 320 张裁图中 249 张受影响，现已全部清理）
6. **渲染层超高图提升**：高度 > 2×字号的公式图提升为独立居中大图，避免顶到上一行；
   纯图片的列表项不提升，避免出现空的列表符号

## 目录结构

```
app.py                  FastAPI 服务 + CLI
static/index.html       前端（上传、模型选择、进度、下载）
pdftrans/extract.py     PDF → 结构化 IR（版式、公式裁图、遮挡）
pdftrans/translate.py   大模型翻译（内置通道 / 自定义接口）、占位符自愈、缓存
pdftrans/build.py       渲染 DOCX / HTML / PDF
requirements.txt
```

## 已验证

- 12 页法语数学讲义（含 56 个定理环境、cases、分数、向量、欧拉公式、单位根）
- 输出：乱码字符 0、空列表项 0、标记残留 0，公式图无邻行残片
- 速度：缓存命中 16 秒；冷启动（150 段）约 20–60 秒（取决于云通道排队情况）

## License

MIT
