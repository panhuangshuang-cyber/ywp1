#!/usr/bin/env python3
"""
SKU 别名映射
优先级：需核查运费SKU.xlsx「别名」列 > 钉钉「对应关系表」
"""

from __future__ import annotations

import os
from functools import lru_cache

import pandas as pd
import requests
import urllib.parse

# ==================== 钉钉配置（别名补充来源） ====================
DING_WORKBOOK_ID = "vy20BglGWOgznx2Ms0ENN6R5WA7depqY"
DING_SHEET_ALIAS = "对应关系表"
DING_OPERATOR_ID = "o6cNzAqelYyzTaxc2KcIfgiEiE"
DING_APPKEY = "dinglhtnqselknfytfqq"
DING_APPSECRET = "S_qytG7eF3SylcE5nRzYTbSMJnayQbJ07nicye3nRhFzzsfbHWGbcCkZNFl0DYKG"

DEFAULT_EXCEL_PATH = os.path.join(os.path.dirname(__file__), "需核查运费SKU.xlsx")


def get_dingtalk_token() -> str:
    url = f"https://oapi.dingtalk.com/gettoken?appkey={DING_APPKEY}&appsecret={DING_APPSECRET}"
    resp = requests.get(url, timeout=10).json()
    if "access_token" not in resp:
        raise RuntimeError(f"[钉钉] Token 获取失败: {resp}")
    return resp["access_token"]


def fetch_alias_mapping() -> pd.DataFrame:
    """从钉钉「对应关系表」拉取别名映射（原始表）"""
    token = get_dingtalk_token()
    sheet_enc = urllib.parse.quote(DING_SHEET_ALIAS)
    headers = {
        "content-type": "application/json",
        "x-acs-dingtalk-access-token": token,
    }
    base = f"https://api.dingtalk.com/v1.0/doc/workbooks/{DING_WORKBOOK_ID}/sheets/{sheet_enc}"
    op = f"?operatorId={DING_OPERATOR_ID}"

    resp = requests.get(base + op, headers=headers, timeout=15).json()
    total_row = resp.get("rowCount", 0)
    print(f"[钉钉对应关系表] 总行数: {total_row}")
    if total_row < 2:
        return pd.DataFrame()

    h_resp = requests.get(
        base + f"/ranges/A1:M{total_row}{op}",
        headers=headers,
        timeout=30,
    ).json()
    display_values = h_resp.get("displayValues", [])
    if not display_values or not display_values[0]:
        return pd.DataFrame()

    header_raw = [str(item).replace("\n", "").strip() for item in display_values[0]]
    return pd.DataFrame(display_values[1:], columns=header_raw)


def _build_alias_map_from_excel(excel_path: str) -> dict[str, str]:
    df = pd.read_excel(excel_path)
    if 'SKU' not in df.columns or '别名' not in df.columns:
        return {}

    alias_map: dict[str, str] = {}
    for _, row in df.iterrows():
        main_sku = str(row['SKU']).strip()
        alias = str(row.get('别名', '')).strip()
        if not main_sku or main_sku.lower() == 'nan':
            continue
        if alias and alias.lower() != 'nan' and alias != main_sku:
            alias_map[alias] = main_sku
    return alias_map


def _build_alias_map_from_dingtalk(df: pd.DataFrame) -> dict[str, str]:
    if df.empty:
        return {}

    cols = {str(col).strip(): col for col in df.columns}
    main_col = next((cols[name] for name in cols if name in ('主SKU', 'SKU', '主sku', 'Our SKU')), None)
    alias_col = next((cols[name] for name in cols if '别名' in name or name in ('Alias', 'alias')), None)
    if not main_col or not alias_col:
        return {}

    alias_map: dict[str, str] = {}
    for _, row in df.iterrows():
        main_sku = str(row[main_col]).strip()
        alias = str(row[alias_col]).strip()
        if main_sku and alias and alias.lower() != 'nan' and main_sku.lower() != 'nan' and alias != main_sku:
            alias_map[alias] = main_sku
    return alias_map


@lru_cache(maxsize=4)
def get_alias_map(excel_path: str = DEFAULT_EXCEL_PATH, use_dingtalk_fallback: bool = True) -> dict[str, str]:
    """
    返回 {别名: 主SKU}。
    Excel「别名」列优先；钉钉表仅补充 Excel 未覆盖的别名。
    """
    alias_map = _build_alias_map_from_excel(excel_path)
    print(f"[别名] Excel 映射: {len(alias_map)} 条")

    if use_dingtalk_fallback:
        try:
            ding_df = fetch_alias_mapping()
            ding_map = _build_alias_map_from_dingtalk(ding_df)
            added = 0
            for alias, main_sku in ding_map.items():
                if alias not in alias_map:
                    alias_map[alias] = main_sku
                    added += 1
            print(f"[别名] 钉钉补充: {added} 条")
        except Exception as exc:
            print(f"[别名] 钉钉拉取失败，仅使用 Excel: {exc}")

    return alias_map


def normalize_sku(raw_sku: str, alias_map: dict[str, str] | None = None) -> str:
    """将别名标准化为主 SKU；未知 SKU 原样返回"""
    sku = str(raw_sku).strip()
    if not sku or sku.lower() == 'nan':
        return ''
    mapping = alias_map if alias_map is not None else get_alias_map()
    return mapping.get(sku, sku)


def normalize_sku_list(skus: list[str], alias_map: dict[str, str] | None = None) -> list[str]:
    mapping = alias_map if alias_map is not None else get_alias_map()
    normalized = [normalize_sku(sku, mapping) for sku in skus]
    return [sku for sku in normalized if sku]


if __name__ == "__main__":
    mapping = get_alias_map()
    print(f"别名总数: {len(mapping)}")
    for alias, main_sku in list(mapping.items())[:10]:
        print(f"{alias} -> {main_sku}")