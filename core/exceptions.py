"""项目自定义异常。"""


class LoadForecastingError(Exception):
    """项目基础异常。"""


class UnsupportedFileTypeError(LoadForecastingError):
    """文件类型不支持。"""


class WorkbookLoadError(LoadForecastingError):
    """工作簿读取失败。"""


class TemplateDetectionError(LoadForecastingError):
    """模板识别失败。"""


class ForecastError(LoadForecastingError):
    """预测或反推失败。"""
