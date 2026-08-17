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

# TEN / HWG を先に定義し、FRV の除外条件をここから自動導出する。
# こうすることで、TEN や HWG のプレフィックスが変わったときに
# FRV の除外設定が自動的に追従し、定義のズレを防ぐ。
BRAND_TEN = BrandConfig(
    name="TEN",
    allowed_names=["TEN", "TENNEN"],
    prefixes=["TN", "TEN"],
    excluded_prefixes=[]
)

BRAND_HWG = BrandConfig(
    name="HWG",
    allowed_names=["HWG", "HANWAG"],
    prefixes=["H"],
    excluded_prefixes=[]
)

# WES ブランドは独立した BrandConfig を持たないが、
# 商品コードが "W" で始まるため FRV から除外する。
_WES_PREFIXES = ("W",)

# FRV の除外プレフィックス = TEN のプレフィックス + HWG のプレフィックス + WES のプレフィックス
# ※ ブランド列がある CSV では is_target_brand() がブランド名を優先するため
#   このプレフィックスはブランド列が存在しない CSV へのフォールバック用。
_FRV_EXCLUDED = list(BRAND_TEN.prefixes) + list(BRAND_HWG.prefixes) + list(_WES_PREFIXES)

BRAND_FRV = BrandConfig(
    name="FRV",
    allowed_names=["FRV", "FJALLRAVEN", "FJÄLLRÄVEN"],
    prefixes=[],  # すべて（除外プレフィックス以外）を FRV とみなす
    excluded_prefixes=_FRV_EXCLUDED  # 現在値: ["TN", "TEN", "H", "W"]
)

# ============================================================
# 店舗名照合用ルール (エイリアス・除外条件)
# ============================================================
STORE_MATCHING_RULES = {
    "ヒュッテ": {"include": ["ヒュッテ", "HUTTE"]},
    "TOKYO": {"include": ["TOKYO"], "exclude": ["NODE"]},
    "NODE": {"include": ["NODE"]},
    "大丸心斎橋": {"include": ["大丸", "心斎橋"]},
    "京王新宿": {"include": ["京王", "新宿"]},
    # BUG-04修正: "新宿" のみの店舗は「京王」を含む店舗名と区別するため exclude を追加
    "新宿": {"include": ["新宿"], "exclude": ["京王"]},
    "玉川高島屋": {"include": ["玉川", "高島屋"]},
    "横浜高": {"include": ["横浜", "高島屋"]},
    "横浜高島屋": {"include": ["横浜", "高島屋"]},
    "名刀": {"include": ["名刀", "名古屋"]},
    "ルクア大阪": {"include": ["大阪", "ルクア"]},
    "ルクア": {"include": ["大阪", "ルクア"]},
    "NARITA": {"include": ["NARITA", "成田"]},
}

