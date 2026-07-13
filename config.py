# -*- coding: utf-8 -*-
"""
設定・データクラス定義モジュール (config.py)
===========================================
"""

# ============================================================
# 例外の定義
# ============================================================
class AutomationError(Exception):
    """業務自動化処理中に発生する想定内のエラー"""
    def __init__(self, title, message):
        self.title = title
        self.message = message
        super().__init__(message)

# ============================================================
# ブランド設定クラスおよびインスタンス定義
# ============================================================
class BrandConfig:
    def __init__(self, name, allowed_names, prefixes, excluded_prefixes=None):
        self.name = name
        self.allowed_names = [n.upper() for n in allowed_names]
        self.prefixes = tuple(p.upper() for p in prefixes) if prefixes else ()
        self.excluded_prefixes = tuple(p.upper() for p in excluded_prefixes) if excluded_prefixes else ()

BRAND_FRV = BrandConfig(
    name="FRV",
    allowed_names=["FRV", "FJALLRAVEN", "FJÄLLRÄVEN"],
    prefixes=[],  # すべて（除外プレフィックス以外）
    excluded_prefixes=["TNT", "TNP", "TNF", "H", "W"]
)

BRAND_TEN = BrandConfig(
    name="TEN",
    allowed_names=["TEN", "TENNEN"],
    prefixes=["TNT", "TNP", "TNF"],
    excluded_prefixes=[]
)

BRAND_HWG = BrandConfig(
    name="HWG",
    allowed_names=["HWG", "HANWAG"],
    prefixes=["H"],
    excluded_prefixes=[]
)
