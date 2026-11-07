#!/usr/bin/env python3
"""Generate the diagrams used by the HapTest submission package.

Requires graphviz (`dot`) on PATH. Outputs PNG at 240 dpi into ../assets/.
"""
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, "..", "assets"))
os.makedirs(OUT, exist_ok=True)

FONT = "PingFang SC"
BLUE = "#4874CB"
BLUE_L = "#EAF0FA"
ORANGE = "#EE822F"
ORANGE_L = "#FDF0E6"
TEAL = "#30C0B4"
TEAL_L = "#E6F7F5"
GREEN = "#75BD42"
GREEN_L = "#EEF7E8"
RED = "#E54C5E"
RED_L = "#FCE9EC"
INK = "#0F1423"
GRAY = "#5B6577"
LINE = "#8A94A6"


def render(name: str, dot: str, dpi: int = 240, engine: str = "dot") -> None:
    dot_path = os.path.join(OUT, name + ".dot")
    png_path = os.path.join(OUT, name + ".png")
    with open(dot_path, "w", encoding="utf-8") as fh:
        fh.write(dot)
    cmd = [engine]
    if engine == "neato":
        cmd.append("-n2")
    cmd += ["-Tpng", f"-Gdpi={dpi}", dot_path, "-o", png_path]
    subprocess.run(cmd, check=True)
    print("wrote", png_path)


# ---------------------------------------------------------------- architecture
architecture = f'''
digraph architecture {{
  graph [bgcolor="white", fontname="{FONT}", fontsize=12, rankdir=TB,
         nodesep=0.22, ranksep=0.30, pad=0.18, compound=true, newrank=true];
  node [shape=box, style="rounded,filled", fontname="{FONT}", fontsize=11,
        color="{BLUE}", fillcolor="{BLUE_L}", penwidth=1.2, margin="0.10,0.06"];
  edge [fontname="{FONT}", fontsize=9, color="{LINE}", arrowsize=0.55, penwidth=1.0];

  subgraph cluster_entry {{
    label="驱动入口"; fontcolor="{BLUE}"; color="{BLUE}"; style="rounded,dashed"; bgcolor="#FAFBFE";
    node [fillcolor="white"];
    cli   [label="harmony CLI\\n命令 · JSON 输出", color="{BLUE}"];
    serve [label="harmony serve\\nJSON-lines 会话", color="{BLUE}"];
    fuzzc [label="Fuzz 入口\\n--policy", color="{BLUE}"];
    web   [label="Web Viewer\\nui-viewer / screen", color="{BLUE}"];
  }}

  subgraph cluster_core {{
    label="核心服务"; fontcolor="{ORANGE}"; color="{ORANGE}"; style="rounded,dashed"; bgcolor="#FFFCF8";
    node [fillcolor="{ORANGE_L}", color="{ORANGE}"];
    hdev   [label="HarmonyDevice\\nobserve · tap · swipe", color="{ORANGE}"];
    runner [label="Fuzz / Runner\\n调度与生命周期", color="{ORANGE}"];
    policy [label="Policy 管线\\n随机 · DFS · 静态 · LLM", color="{ORANGE}"];
    ptg    [label="PTG / Model\\n页面迁移图", color="{ORANGE}"];
  }}

  subgraph cluster_analysis {{
    label="分析与观测"; fontcolor="{TEAL}"; color="{TEAL}"; style="rounded,dashed"; bgcolor="#F7FCFB";
    node [fillcolor="{TEAL_L}", color="{TEAL}"];
    static [label="静态分析\\nArkUI 布局", color="{TEAL}"];
    llm    [label="LLM 引导\\n动作评分", color="{TEAL}"];
    perf   [label="性能观测\\nDevEco trace", color="{TEAL}", style="rounded,filled,dashed"];
    report [label="证据与报告\\n日志 · 截图 · 用例", color="{TEAL}"];
  }}

  subgraph cluster_device {{
    label="设备适配"; fontcolor="{GREEN}"; color="{GREEN}"; style="rounded,dashed"; bgcolor="#FAFCF7";
    node [fillcolor="{GREEN_L}", color="{GREEN}"];
    hdc    [label="HDC\\nshell · file", color="{GREEN}"];
    uitest [label="uitest agent\\nUI 树 · 截图", color="{GREEN}"];
    hypium [label="Hypium RPC\\n测试协议", color="{GREEN}"];
    arkui  [label="ArkUI Inspector\\n控件树", color="{GREEN}"];
  }}

  device [label="OpenHarmony 设备\\nArkTS 应用 · ArkUI · uitest", shape=box3d,
          style="filled", fillcolor="{INK}", fontcolor="white", color="{INK}", fontsize=13];

  cli -> hdev; serve -> hdev;
  fuzzc -> runner; runner -> policy; policy -> ptg; ptg -> hdev;
  policy -> static [style=dashed, label="源码"];
  policy -> llm [style=dashed, label="--llm"];
  hdev -> hdc; hdev -> uitest; hdev -> hypium; hdev -> arkui;
  hdc -> device; uitest -> device; hypium -> device; arkui -> device;
  runner -> report [style=dashed, label="结果"];
  perf -> report [style=dashed];

  {{ rank=same; cli; serve; fuzzc; web; }}
  {{ rank=same; hdev; runner; policy; ptg; }}
  {{ rank=same; static; llm; perf; report; }}
  {{ rank=same; hdc; uitest; hypium; arkui; }}
}}
'''
render("architecture", architecture)


# ---------------------------------------------------------------- PTG loop
ptg_loop = f'''
digraph ptg_loop {{
  graph [bgcolor="white", fontname="{FONT}", fontsize=12, dim="1050,470", pad=0.15];
  node [shape=box, style="rounded,filled", fontname="{FONT}", fontsize=12,
        color="{BLUE}", fillcolor="{BLUE_L}", penwidth=1.3, margin="0.14,0.08", pin=true];
  edge [fontname="{FONT}", fontsize=10, color="{LINE}", arrowsize=0.7, penwidth=1.2];

  obs   [label="1  观察\nobserve / dumpLayout", color="{BLUE}", fillcolor="{BLUE_L}", pos="120,360!"];
  state [label="2  状态识别\n页面 + 控件 + 内容签名", color="{TEAL}", fillcolor="{TEAL_L}", pos="525,360!"];
  cand  [label="3  候选动作\n属性 + 可执行性过滤", color="{ORANGE}", fillcolor="{ORANGE_L}", pos="930,360!"];
  score [label="4  策略评分\n目标 · 新状态 · 历史 · 预算", color="{ORANGE}", fillcolor="{ORANGE_L}", pos="930,110!"];
  act   [label="5  受约束执行\ntap / swipe / type / key", color="{GREEN}", fillcolor="{GREEN_L}", pos="525,110!"];
  ptg   [label="6  更新 PTG\n节点 + 迁移边 + 证据", color="{RED}", fillcolor="{RED_L}", pos="120,110!"];

  obs -> state;
  state -> cand;
  cand -> score;
  score -> act;
  act -> ptg;
  ptg -> obs [label="重新观察", style=dashed];
}}
'''
render("ptg-loop", ptg_loop, engine="neato")


# ---------------------------------------------------------------- device API
device_api = f'''
digraph device_api {{
  graph [bgcolor="white", fontname="{FONT}", fontsize=12, rankdir=TB,
         nodesep=0.28, ranksep=0.30, pad=0.20, newrank=true];
  node [shape=box, style="rounded,filled", fontname="{FONT}", fontsize=11,
        color="{BLUE}", fillcolor="{BLUE_L}", penwidth=1.2, margin="0.11,0.06"];
  edge [fontname="{FONT}", fontsize=9, color="{LINE}", arrowsize=0.55, penwidth=1.0];

  api [label="HarmonyDevice（src/device/harmony/session.ts）\\nconnect · observe · find · tap · swipe · typeText · pressKey · launchApp · stopApp · close",
       fillcolor="{BLUE}", fontcolor="white", color="{BLUE}", fontsize=12];

  subgraph cluster_inner {{
    label="会话与约束"; fontcolor="{ORANGE}"; color="{ORANGE}"; style="rounded,dashed"; bgcolor="#FFFCF8";
    node [fillcolor="{ORANGE_L}", color="{ORANGE}"];
    selectors [label="selectors\\ntext / id / hint / name + index", color="{ORANGE}"];
    lease     [label="lease\\n跨进程文件租约 · DEVICE_BUSY", color="{ORANGE}"];
    errors    [label="errors\\n统一错误码 + nextStep", color="{ORANGE}"];
  }}
  subgraph cluster_driver {{
    label="设备驱动"; fontcolor="{GREEN}"; color="{GREEN}"; style="rounded,dashed"; bgcolor="#FAFCF7";
    node [fillcolor="{GREEN_L}", color="{GREEN}"];
    hdc    [label="Hdc（src/device/hdc.ts）\\nshell · file send/recv · targets", color="{GREEN}"];
    agent  [label="uitest agent\\ndumpLayout · screenshot · input", color="{GREEN}"];
    hypium [label="Hypium RPC\\n测试框架可选通道", color="{GREEN}"];
  }}
  device [label="OpenHarmony 设备", shape=box3d, style="filled", fillcolor="{INK}", fontcolor="white", color="{INK}", fontsize=12];

  invariants [shape=box, style="rounded,filled", fillcolor="{RED_L}", color="{RED}",
    label=<<TABLE BORDER="0" CELLBORDER="0" CELLSPACING="3" CELLPADDING="2">
      <TR><TD COLSPAN="2" ALIGN="CENTER"><B>安全不变量（由 HarmonyDevice 强制）</B></TD></TR>
      <TR><TD ALIGN="LEFT">① 动作前必须有 observation；mutate 后立即失效</TD><TD ALIGN="LEFT">⑤ 动作异常 → ACTION_OUTCOME_UNCERTAIN，replayAction=false</TD></TR>
      <TR><TD ALIGN="LEFT">② 坐标动作绑定 observationId + bundle + 屏幕尺寸 + contentSig</TD><TD ALIGN="LEFT">⑥ 同一 serial 文件租约独占；锁屏不返回树与截图</TD></TR>
      <TR><TD ALIGN="LEFT">③ 任一变化 → STALE_OBSERVATION，不执行动作</TD><TD ALIGN="LEFT">⑦ 认证字段命中敏感词 → SENSITIVE_FIELD，拒绝代输</TD></TR>
      <TR><TD ALIGN="LEFT">④ selector 多候选 → AMBIGUOUS_TARGET + 候选列表，不猜</TD><TD ALIGN="LEFT"></TD></TR>
    </TABLE>>];

  api -> selectors; api -> lease; api -> errors;
  api -> hdc; api -> agent; api -> hypium;
  hdc -> device; agent -> device; hypium -> device;
  device -> invariants [style=invis];
  {{ rank=same; selectors; lease; errors; }}
  {{ rank=same; hdc; agent; hypium; }}
}}
'''
render("device-api", device_api)


# ---------------------------------------------------------------- perf pipeline
perf_pipeline = f'''
digraph perf_pipeline {{
  graph [bgcolor="white", fontname="{FONT}", fontsize=12, rankdir=TB,
         nodesep=0.28, ranksep=0.34, pad=0.18, newrank=true];
  node [shape=box, style="rounded,filled", fontname="{FONT}", fontsize=11,
        color="{TEAL}", fillcolor="{TEAL_L}", penwidth=1.2, margin="0.11,0.06"];
  edge [fontname="{FONT}", fontsize=9, color="{LINE}", arrowsize=0.55, penwidth=1.0];

  probe [label="probeMetrics()\n探测可用帧数据源", color="{BLUE}", fillcolor="{BLUE_L}"];
  hitrace [label="hitrace\n--trace_begin / --trace_finish", color="{ORANGE}", fillcolor="{ORANGE_L}"];
  htrace [label=".htrace\n原始 trace", color="{ORANGE}", fillcolor="{ORANGE_L}"];
  streamer [label="trace_streamer\nDevEco 本地解析器", color="{TEAL}", fillcolor="{TEAL_L}"];
  sqlite [label="SQLite\nframe_slice · frame_total_statistics", color="{TEAL}", fillcolor="{TEAL_L}"];
  parser [label="FrameParser\n逐帧 duration + flag", color="{GREEN}", fillcolor="{GREEN_L}"];
  sample [label="FrameSample\ntotalFrames · jankFrames · p95", color="{GREEN}", fillcolor="{GREEN_L}"];
  report [label="ActionTrace / report\n每次动作的帧与耗时", color="{RED}", fillcolor="{RED_L}"];

  probe -> hitrace [label="A 路线", constraint=false];
  hitrace -> htrace;
  htrace -> streamer [constraint=false];
  streamer -> sqlite;
  sqlite -> parser [constraint=false];
  parser -> sample;
  sample -> report [constraint=false];
  {{ rank=same; probe; hitrace; }}
  {{ rank=same; htrace; streamer; }}
  {{ rank=same; sqlite; parser; }}
  {{ rank=same; sample; report; }}
}}
'''
render("perf-pipeline", perf_pipeline)


# ---------------------------------------------------------------- policy pipeline
policy_pipeline = f'''
digraph policy_pipeline {{
  graph [bgcolor="white", fontname="{FONT}", fontsize=12, rankdir=LR,
         nodesep=0.30, ranksep=0.40, pad=0.18];
  node [shape=box, style="rounded,filled", fontname="{FONT}", fontsize=11,
        color="{BLUE}", fillcolor="{BLUE_L}", penwidth=1.2, margin="0.11,0.06"];
  edge [fontname="{FONT}", fontsize=9, color="{LINE}", arrowsize=0.55, penwidth=1.0];

  cand [label="候选事件\\n控件 + 事件类型", color="{BLUE}", fillcolor="{BLUE_L}"];
  filter [label="过滤\\n不可见 · 不可用 · 敏感", color="{TEAL}", fillcolor="{TEAL_L}"];
  score [label="评分\\n新状态 · 目标 · 历史", color="{ORANGE}", fillcolor="{ORANGE_L}"];
  pick [label="选择\\n策略决定下一步", color="{GREEN}", fillcolor="{GREEN_L}"];
  exec [label="执行\\n共用 Executor", color="{RED}", fillcolor="{RED_L}"];
  cand -> filter -> score -> pick -> exec;
}}
'''
render("policy-pipeline", policy_pipeline)
