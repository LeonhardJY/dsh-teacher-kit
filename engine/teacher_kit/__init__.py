"""teacher-kit 文档引擎。

设计约定：
- 所有对外能力都从一个入口走：engine/cli.py，协议是"JSON 进、JSON 出"。
  宿主（Node）只负责调度、渲染和 UI，不碰文档结构；文档结构只在这里改。
- 输出里给人的文字一律中文；错误一律 {ok:false, error:{code, message, hint?}}。
- 不联网、不写用户没指定的目录、默认不覆盖已有文件。
"""

__version__ = "0.18.0"