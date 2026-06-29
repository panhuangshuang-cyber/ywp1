#!/usr/bin/env python3
"""
运德订单 API 模块（独立可复用）
负责拉取运德出库单数据并完成基础清洗
"""

import ast
import hashlib
import json
import os
from datetime import datetime, timedelta

import pandas as pd
import requests

# ==================== 运德 API 配置（可改） ====================
USER_ACCOUNT = 'cUMz0885'
ACCESS_TOKEN = 'A932EDB48CE0434506A9525F9DD26EFD'
ORDER_API_URL = "http://fg.wedoexpress.com/api.php?mod=apiManage&act=queryOrderInfo"

BOX_SKUS = ('36BOX01', '56BOX02', '55BOX04')


def _clear_proxy_env() -> None:
    for key in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy', 'ALL_PROXY', 'all_proxy'):
        os.environ.pop(key, None)


def yd_create_signature(params: dict, access_token: str) -> str:
    """生成运德 API 签名"""
    params_sorted = dict(sorted(params.items()))
    sign_str = ''.join(str(params_sorted[k]) for k in params_sorted)
    return hashlib.md5((sign_str + access_token).encode('utf-8')).hexdigest().upper()


def parse_goods_info(goods_info) -> list[dict]:
    """解析 goodsInfo，返回商品行列表（每项含 userSku / sku / qty）"""
    if goods_info is None:
        return []
    if isinstance(goods_info, list):
        return [item for item in goods_info if isinstance(item, dict)]
    text = str(goods_info).strip()
    if not text:
        return []
    for loader in (json.loads, ast.literal_eval):
        try:
            parsed = loader(text)
        except (json.JSONDecodeError, SyntaxError, ValueError):
            continue
        if isinstance(parsed, list):
            return [item for item in parsed if isinstance(item, dict)]
        if isinstance(parsed, dict):
            return [parsed]
    return []


def extract_user_skus(goods_info) -> list[str]:
    """从 goodsInfo 提取 userSku 列表（保持原始顺序）"""
    skus = []
    for item in parse_goods_info(goods_info):
        sku = str(item.get('userSku') or item.get('sku') or '').strip()
        if sku:
            skus.append(sku)
    return skus


def build_merge_sku_key(skus: list[str]) -> str:
    """
    将一单内的多个 SKU 合并为一个聚合键。
    纸箱 SKU（36BOX01 / 56BOX02 / 55BOX04）排在最前，其余保持 goodsInfo 原始顺序。
    例：['BO8SC', 'MC-S1006C', '36BOX01'] -> '36BOX01/BO8SC/MC-S1006C'
    """
    unique_skus: list[str] = []
    seen: set[str] = set()
    for sku in skus:
        if sku and sku not in seen:
            seen.add(sku)
            unique_skus.append(sku)
    if not unique_skus:
        return ''
    if len(unique_skus) == 1:
        return unique_skus[0]

    box_part = [sku for sku in BOX_SKUS if sku in seen]
    other_part = [sku for sku in unique_skus if sku not in BOX_SKUS]
    return '/'.join(box_part + other_part)


def format_date_label(start_date: str, end_date: str) -> str:
    """生成 Excel 表头用的日期区间，如 2026.3.31-2026.6.29"""
    start = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    return f"{start.year}.{start.month}.{start.day}-{end.year}.{end.month}.{end.day}"


def get_order_date_range(days: int = 90) -> tuple[str, str, str]:
    """
    返回订单查询区间：当天往前推 days 天 至 当天（含首尾）。
    返回 (start_date, end_date, date_label)
    """
    end = datetime.now().date()
    start = end - timedelta(days=days)
    start_date = start.strftime('%Y-%m-%d')
    end_date = end.strftime('%Y-%m-%d')
    return start_date, end_date, format_date_label(start_date, end_date)


def fetch_wedo_orders(days: int = 90) -> pd.DataFrame:
    """拉取最近 N 天出库单（自动分页），返回原始 DataFrame"""
    _clear_proxy_env()

    start_date, end_date, date_label = get_order_date_range(days)
    print(f"[运德API] 查询区间: {start_date} ~ {end_date} ({date_label})")

    all_orders = []
    page = 1
    while True:
        content = {
            'dateType': 1,
            'startDate': start_date,
            'endDate': end_date,
            'page': page,
            'pageSize': 100,
        }
        params = {
            'userAccount': USER_ACCOUNT,
            'content': json.dumps(content, ensure_ascii=False, separators=(',', ':')),
        }
        params['sign'] = yd_create_signature(params, ACCESS_TOKEN)

        try:
            resp = requests.post(ORDER_API_URL, data=params, timeout=120)
            resp.raise_for_status()
            result = json.loads(resp.text)
            orders = result.get('data') or result.get('list') or result.get('rows') or []
            if not orders:
                break
            all_orders.extend(orders)
            print(f"[运德API] 第{page}页 {len(orders)} 条，累计 {len(all_orders)}")
            page += 1
            if page > 200:
                print(f"[运德API] 达到页数上限(200)，已拉取 {len(all_orders)} 条")
                break
        except Exception as exc:
            print(f"[运德API] 第{page}页失败: {exc}")
            break

    return pd.DataFrame(all_orders) if all_orders else pd.DataFrame()


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def clean_orders(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    total = len(df)
    working = df.copy()

    status_col = _find_col(working, ['status', '状态', 'orderStatus', 'order_status'])
    if status_col:
        status_text = working[status_col].astype(str).str.strip()
        invalid_status = status_text.str.contains('废弃|cancel|abandon', case=False, na=False)
        dropped = int(invalid_status.sum())
        if dropped:
            print(f"[清洗] 删除无效状态: {dropped} 条")
        working = working[~invalid_status]

    fee_col = _find_col(working, ['shipfee', 'freight', '运费', 'fee', 'shippingFee'])
    if fee_col:
        fee_num = pd.to_numeric(working[fee_col], errors='coerce')
        fee_text = working[fee_col].astype(str).str.strip()
        invalid_fee = fee_num.isna() | fee_num.lt(3) | fee_text.isin(['', '-', '—', 'null', 'NaN', 'None'])
        dropped = int(invalid_fee.sum())
        if dropped:
            print(f"[清洗] 删除无效运费(<3/空/-): {dropped} 条")
        working = working[~invalid_fee]

    goods_col = _find_col(working, ['goodsInfo', 'goods_info', '产品信息'])
    if goods_col:
        user_skus_list = working[goods_col].apply(extract_user_skus)
        qty_valid = working[goods_col].apply(
            lambda value: all(
                pd.to_numeric(item.get('qty', 1), errors='coerce') == 1
                for item in parse_goods_info(value)
            )
            if parse_goods_info(value) else False
        )
        dropped = int((~qty_valid).sum())
        if dropped:
            print(f"[清洗] 删除商品数量≠1: {dropped} 条")
        has_skus = user_skus_list.apply(len) > 0
        working = working[qty_valid & has_skus].copy()
        working['user_skus'] = user_skus_list.loc[working.index].tolist()
    else:
        working['user_skus'] = [[] for _ in range(len(working))]

    print(f"[清洗] 完成: {total} → {len(working)} 条 (删除 {total - len(working)} 条)")
    return working.reset_index(drop=True)


def get_clean_orders(days: int = 90) -> pd.DataFrame:
    """拉取并清洗订单，返回附带 user_skus 列的 DataFrame"""
    return clean_orders(fetch_wedo_orders(days=days))


if __name__ == "__main__":
    sample = get_clean_orders(days=3)
    print(sample.head() if not sample.empty else "无数据")
    print(f"列名: {list(sample.columns) if not sample.empty else []}")
    print(f"清洗后条数: {len(sample)}")