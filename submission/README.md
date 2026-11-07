# 2026 开源鸿蒙大学生创新大赛 · 赛道一 提交材料

**作品：HapTest: 面向 OpenHarmony 应用的多模态智能测试系统**

> OpenHarmony 应用的智能测试与性能观测基础设施

| 项目 | 内容 |
| --- | --- |
| 学校 | 北京航空航天大学 |
| 团队 | SMAT 团队 |
| 队长 | 刘发荣 |
| 作品仓库 | GitHub: https://github.com/growvv/HapTest-2026；GitCode: https://gitcode.com/rogn/HapTest-2026 |
| 开源许可 | Apache License 2.0 |

## 一、材料索引

| 文件 | 说明 |
| --- | --- |
| `01-参赛作品简介.docx` / `.pdf` | 一页作品定位、问题、方案、创新点、完成度与来源标注 |
| `02-项目文档.docx` / `.pdf` | 按大赛模板撰写的完整设计文档（简介、设计描述、系统结构、模块、流程、接口、UI、分工、困难、参考） |
| `03-答辩PPT.pptx` / `.pdf` | 16 页答辩 PPT（按大赛模板主题制作） |
| `04-演示视频.mp4` | 2 分 05 秒、1920×1080、带中文旁白的演示视频 |
| `05-辅助材料/` | FSE'25 / ASE'25 论文首页、专利审查通知书、测试与构建记录、来源与版权说明 |
| `演示视频脚本.md` | 视频分镜、旁白与时间轴，便于复现或重新录制 |
| `assets/` | 架构图、流程图、UI 示意图等源图 |
| `templates/` | 大赛提供的空白模板（docx / pptx），便于复核版式 |
| `build-scripts/` | 文档、PPT、图表与视频的生成脚本（开发工具，不是提交必需） |

## 二、如何查看

- 文档：直接打开 `.docx` / `.pdf`；PDF 为版式预览，Word 中可继续编辑。
- PPT：打开 `.pptx`；`.pdf` 为逐页预览。PPT 内含演讲者备注，可作为答辩讲稿。
- 视频：直接播放 `04-演示视频.mp4`；建议 1080p 全屏观看。
- 辅助材料：`05-辅助材料/README.md` 有逐项说明。

## 三、如何复现材料（可选）

生成脚本位于 `build-scripts/`，依赖：

- Python 3 + `python-docx`、`python-pptx`、`Pillow`
- Graphviz（`dot`）
- LibreOffice（docx/pptx → PDF）
- Node.js + Playwright（UI 示意与视频帧）
- ffmpeg、Qwen TTS（dashscope，中文旁白）、Playwright（UI 示意与视频帧）

常用命令：

```sh
python3 build-scripts/diagrams.py          # 生成架构/流程图
python3 build-scripts/build_docx.py        # 生成 02-项目文档.docx
python3 build-scripts/build_pptx.py        # 生成 03-答辩PPT.pptx
python3 build-scripts/build_intro_docx.py  # 生成 01-参赛作品简介.docx
python3 build-scripts/tts_qwen.py           # 生成 Qwen 旁白与字幕时间轴（需 DASHSCOPE_API_KEY）
python3 build-scripts/build_video_qwen.py   # 合成带字幕的演示视频
```

PDF 由 LibreOffice headless 导出；视频由 `video_script.json`、Qwen TTS（`tts_qwen.py`）与 Playwright/ffmpeg 生成，字幕按句对齐后烧录。脚本中的字体、路径按 macOS 环境配置，其他平台需替换为可用中文字体。

## 四、来源与边界

- **FSE'25 论文**：HapTest 动态分析框架，第一作者刘发荣；20 款商业应用、26 个未报告问题等数据来自论文表 4，已在文档与 PPT 中标注来源。
- **ASE'25 论文**：UI 重叠实证研究与 HCO-Eye，第一作者刘发荣；100 款应用、3326 万个重叠实例、HCOBench 93.12% F1 等数据来自论文，已在文档与 PPT 中标注来源。
- **发明专利**：申请号 202510375959.1，申请人北京航空航天大学；审查通知书为官方公开文件。
- **代码底座**：仓库起源于 SMAT-Lab/HapTest 与华为开源 HapTest 工作，保留 Apache-2.0 LICENSE 与上游版权头；2026 年新增 `HarmonyDevice`、`haptest harmony` CLI、JSON-lines `serve`、安全不变量、真机 E2E 与性能观测设计。
- **不造假**：当前仓库构建通过、60 个单元测试全绿；无设备时 E2E 明确 `SKIPPED` 并退出 0；性能观测只使用有数据来源的指标；VLM 多模态诊断为 ASE'25 研究延伸，未并入当前仓库主链路。

## 五、团队

北京航空航天大学 SMAT 团队 · 刘发荣。论文、专利与开源项目为研究积累，按来源标注；本仓库代码与测试为可复现的工程交付。
