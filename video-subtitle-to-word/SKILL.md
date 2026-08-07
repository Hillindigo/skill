---
name: video-subtitle-to-word
description: 将公开抖音视频解析、下载并转写为 Word 文档。用于用户提供抖音链接并要求获取字幕、生成完整逐句字幕版、保留博主原话的分点文章版，或指定其他 Word 排版样式时；优先免费公开解析接口，失败后切换其他免费工具。
---

# 抖音字幕转 Word

## 目标

将公开抖音链接处理为可交付的 `.docx`，默认支持两种版本：

1. **完整逐句字幕版**：保留每条转写文本和起止时间。
2. **博主原话分点文章版**：保留博主原话、顺序和表达，不做观点总结；只做分段、标点、主题标题和高置信度 ASR 错字修正。

## 交互门

在下载或转写前，先询问使用者选择：

- `完整逐句字幕版`
- `博主原话分点文章版`
- `其他样式`：让使用者说明样式、是否保留时间戳、是否允许改写。
- 生成文件保存位置：用户可提供目标路径；未提供时使用默认保存位置 `C:\Users\Ni\Desktop\`。
- 是否保存原视频：保存或不保存。

不得默认把“原话整理”改写成摘要。若选择博主原话版，标题是排版辅助，不得把助手自己的分析写进正文。

## 处理流程

### 1. 解析抖音链接

只处理抖音链接，优先使用免费公开接口，不要求用户提供账号 Cookie 或付费 API Key。

按以下顺序尝试解析接口（主用 `douyin.wtf`）：

```text
GET https://douyin.wtf/api/hybrid/video_data?url=<urlencoded-douyin-url>
```

验证：HTTP 成功、JSON `code == 200`、存在 `data.video` 中的可播放地址（优先 H.264），且返回的媒体 URL 可访问。记录创作者、标题、时长、原链接和解析接口。

主接口失败后按顺序尝试以下备用接口：

1. `http://api.youman.team/douyin?url=<urlencoded-douyin-url>`，设置有限超时；验证 `code == 200`、存在 `data.url`，并将 `duration` 按秒处理。
2. `https://collect.xmwxxc.com/collect/douyin/?url=<urlencoded-douyin-url>`，验证返回的 `url` 是可访问的视频地址。
3. `https://api.bugpk.com/api/douyin?url=<urlencoded-douyin-url>`，设置有限超时。
4. 其他当前可访问的免费公开解析工具；优先选择无需登录、无需上传账号 Cookie 的工具。
5. 用户可访问的免费网页解析器，例如 Easydown 或同类服务；必要时用浏览器完成粘贴和下载。

接口切换时记录失败原因。无论接口返回什么，都必须验证媒体文件存在、类型为视频且可读取；不能把“搜索结果有相关页面”当成视频已获取。若所有接口和网页路径都失败，应明确报告解析未完成，并请求用户提供已下载的视频文件，不得虚构字幕。

### 2. 下载和转音频

将视频保存到当前任务的 `work/`，不要把中间视频、模型缓存写入用户主目录。使用 FFmpeg 提取本地转写音频：

```text
ffmpeg -y -i input.mp4 -vn -ac 1 -ar 16000 -c:a pcm_s16le audio.wav
```

若 FFmpeg、faster-whisper 或 `python-docx` 不存在，先检查工作区依赖运行时；只有在用户确认或已有安装规则允许时安装免费依赖。

### 3. 本地语音转写

优先使用本地 `faster-whisper`，不上传音频到付费云服务：

```python
from faster_whisper import WhisperModel

model = WhisperModel("small", device="cpu", compute_type="int8")
segments, info = model.transcribe(
    audio_path,
    language="zh",
    beam_size=5,
    vad_filter=True,
    condition_on_previous_text=True,
)
```

将每个 segment 保存为 JSON：`start`、`end`、`text`，同时记录识别语言和总时长。模型缓存设置到任务 `work/hf-cache/`，避免 Windows 受限目录权限问题。不得把 ASR 结果描述为“官方字幕”；应称为“语音转写”，除非确实从视频获得了官方字幕轨道。

### 4. 生成 Word

使用本 Skill 的 `scripts/build_word.py`：

- `--mode full`：生成带时间戳的完整字幕表格。
- `--mode original-article --sections sections.json`：按照已确认的主题范围，把原始 segment 文本重排成文章。

生成文件时，先在用户指定的路径下创建一个任务文件夹；用户未提供路径时，在默认位置
`C:\Users\Ni\Desktop\` 下创建。文件夹名称使用：

`抖音_<创作者>_<视频标题>`

例如：`抖音_林olin_犹豫的成本比选错的成本更高`。文件夹名称必须清理 Windows
不允许的字符并限制长度，避免路径无法创建。

将生成的 Word 文件放入该文件夹；如果用户选择保存视频，也将视频放入同一文件夹。
完整字幕版和博主原话文章版可以同时放入该文件夹。中间音频、字幕 JSON 和模型缓存继续放在当前任务的工作目录，不作为最终交付文件散落到目标路径。

完成后必须告诉用户：实际保存的文件夹路径、生成的 Word 文件名，以及是否保存了视频。

文档必须写入来源、转写方法、版本类型和准确性说明。创建或修改 DOCX 后，优先调用 `documents` Skill 的 `render_docx.py` 做 PNG 视觉检查；若环境缺少 LibreOffice/soffice，结构校验仍要完成，并在最终说明未完成视觉 QA。

## 两种版本的内容规则

### 完整逐句字幕版

- 不做摘要、不删除停顿内容，不合并成观点文章。
- 每行包含起止时间和对应转写文本。
- 允许修正明显标点，但不能改变语义。
- 文档标题标注“AI 语音转写版”。

### 博主原话分点文章版

- 先根据原始 segment 建立主题范围；正文只能由原始 segment 文本组成。
- 可增加主题标题，例如“注意力会被手机带走”，但标题不冒充博主原话。
- 只允许：分段、标点、去除明显重复口头填充、修正高置信度识别错字。
- 不允许：加入总结、心理分析、事实核查结论、助手观点或外部资料。
- 在文档开头明确说明“博主原话整理版，未进行观点改写”；若用户要求逐字保留，则禁止删除口头重复。

## 失败边界

- 解析失败：说明已尝试的免费接口和具体错误，继续尝试下一条路径，不虚构已下载。
- 视频可下载但音频转写失败：保留视频和错误日志，说明缺少哪一步，不生成伪字幕。
- ASR 质量不稳定：交付前抽查开头、中段、结尾；把不确定词保留为转写结果并标注可能有误，不擅自猜测。
- 没有可写桌面权限：先生成 `outputs/` 文件，再请求用户允许复制到桌面。

## 可复用脚本

- `scripts/build_word.py`：从 transcript JSON 生成两种 Word 版式；只负责确定性排版，不负责联网解析和语义分段。
