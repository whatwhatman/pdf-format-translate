# -*- coding: utf-8 -*-
"""pdft · 原格式 PDF 翻译工具（本地 Web / CLI）。

亮点
----
* 公式不交给模型重打：按原页面裁切成高清图（或尝试重建 LaTeX），<｜hy_place▁holder▁no▁813｜>块 renvoi 给译文。
* 版式由几何反推：页边距、字号、颜色、tcolorbox 色块、页眉页脚全部还原。
* 翻译批量并发 + 结果缓存 + 术语去重，几十秒出稿。
"""
from .extract import Doc, Block, Run, extract as extract_document
from . import translate, build

__all__ = ['Doc', 'Block', 'Run', 'extract_document', 'translate', 'build']
