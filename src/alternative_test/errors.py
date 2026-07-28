"""固定异常类与退出码。"""

from __future__ import annotations


class AlternativeTestError(Exception):
    """所有替代测试异常的基类。"""

    exit_code: int = 6

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


class InputUnusableError(AlternativeTestError):
    """输入文件损坏、加密或格式伪装且无法读取。"""

    exit_code = 3


class NeedsAnswerError(AlternativeTestError):
    """需要用户回答歧义问题但尚未提供回答。"""

    exit_code = 2


class ValidationError(AlternativeTestError):
    """验证失败 — 源哈希变化、结构不完整等。"""

    exit_code = 4


class OutputConflictError(AlternativeTestError):
    """输出路径已被占用或等于输入路径。"""

    exit_code = 5


class AmountParseError(AlternativeTestError):
    """金额解析失败，含文件、Sheet、行、列定位信息。"""

    def __init__(self, message: str, source_info: str = "") -> None:
        super().__init__(message)
        self.source_info = source_info
