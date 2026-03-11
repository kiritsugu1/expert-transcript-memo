```markdown
# 专家访谈转录纪要助手

这个仓库包含两个 Python 脚本，用于把「专家访谈 / 专家讲解」的录音转写文本（如阿里听悟、飞书转录导出的 txt/docx）自动切分、抽取，并生成结构化纪要，再同时输出一份 Markdown 纪要和一份由该 Markdown 渲染生成的 Word（docx）文档。

## 适用场景

- 只适合 **1 名专家主讲** 的访谈/分享场景：有一个主讲专家，其他人主要是提问或互动。  
- 典型用法：专家访谈、培训录音、咨询项目访谈、内部分享会等。

## 仓库内的两个脚本

- `专家call 纪要 V1 书面prompt 202603039`：在保证信息完整的前提下，做**适度书面化**，会略微修改句子，使纪要更像正式文稿。  
- `专家call 纪要 V1 口语prompt 202603039`：尽量**保持原话风格**，只做去噪、删口头禅，不做明显改写，更接近说话现场。

两份脚本的整体流程和数据结构是一致的，主要差异只在于 prompt 设计和输出文风。

## 核心功能

- 支持导入阿里听悟、飞书等工具导出的 txt 或 docx 转录文件。  
- 自动按时间戳解析原始转录文本，切成带开始时间、结束时间和主题的语义片段。  
- 利用大模型（通过 OpenAI 接口）对每个片段生成结构化纪要，并最终合并为一份 Markdown 纪要，同时基于这份 Markdown 生成对齐格式的 Word（docx）文档。  
- 内置并发处理逻辑，可以并行处理多个片段，加快长录音的处理速度。

## 使用前必填配置

在脚本中找到如下配置区，按自己的环境修改路径和参数（Windows 路径注意使用 `r"..."` 原始字符串或双反斜杠）：

```python
# 配置文件路径 - 支持 .txt / .docx 文件，也支持文件夹批量处理
TRANSCRIPT_PATH = r"E:\\call summarizer\\transcript\\Interview recording 183.txt"  # 可以是文件或文件夹
OUTPUT_DIR = Path(r"E:\\call summarizer\\output")  # 输出文件夹

# 可指定目标说话人（留空则使用 AI 自动识别的专家）
TARGET_SPEAKER = ""  # 例如: "说话人 2" 或留空 ""
```

在 `GeminiHandler` 中配置自己的 API Key、Base URL 和模型名称：

```python
class GeminiHandler:
    def __init__(self):
        self.client = OpenAI(
            api_key="YOUR_API_KEY",
            base_url="YOUR_BASE_URL"
        )

    def request_gpt(self, system_prompt, user_prompt, output_type="text", temperature=0, max_retries=3):
        for attempt in range(max_retries):
            try:
                params = {
                    "model": "gemini-2.5-pro-thinking",  # 如需更换模型，在此修改
                    "temperature": temperature,
                    "messages": msg,
                    "stream": True
                }
                ...
```

- `api_key`：填入你自己的 API Key。  
- `base_url`：填入你自己使用的 OpenAI 兼容服务地址。  
- `model`：默认为 `"gemini-2.5-pro-thinking"`，如需更换模型，在这里替换即可。  
- `temperature`：脚本里默认是 `0`，更偏向稳定、可控输出。这个场景一般不需要模型「自由发挥」，有需要可以自行调整。

## 运行方式与输入输出

1. 按上面的说明配置好 `TRANSCRIPT_PATH`、`OUTPUT_DIR`、`TARGET_SPEAKER` 和 `GeminiHandler` 中的 API 参数。  
2. 直接用 Python 运行其中一个脚本，例如：  

   ```bash
   python Zhuan-Jia-call-Ji-Yao-V1-Shu-Mian-prompt-202603039.py
   ```

3. 运行完成后，在 `OUTPUT_DIR` 指定的文件夹下会生成两类文件：  
   - `xxx_memo.md`：按主题结构化后的 Markdown 纪要  
   - `xxx_memo.docx`：基于 Markdown 自动排版生成的 Word 纪要（支持标题、项目符号等格式）

## 专家说话人选择

- 如果你知道专家在转录系统里的说话人编号（例如「说话人 2」），可以在代码中的 `TARGET_SPEAKER` 参数手动填写。  
- 如果 `TARGET_SPEAKER` 留空，脚本会调用大模型自动识别核心受访专家，并以该说话人视角生成纪要。

## 注意事项

- 当前脚本为个人效率工具 Demo，未做完备的异常处理和参数校验，建议在小范围内部使用后再集成到正式流程。  
- 不适合「多人自由讨论」「没有明显专家主讲人」的会议场景，这类场景建议单独设计分段逻辑和 prompt。  
```
