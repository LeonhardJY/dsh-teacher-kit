"""统一错误形状：宿主侧只认 code。"""


class EngineError(Exception):
    """可预期的业务错误（文件不存在、格式不支持、规则不满足）。"""

    def __init__(self, code, message, hint=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint

    def to_dict(self):
        payload = {"code": self.code, "message": self.message}
        if self.hint:
            payload["hint"] = self.hint
        return payload


def require(condition, code, message, hint=None):
    if not condition:
        raise EngineError(code, message, hint)
