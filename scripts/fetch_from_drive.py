import os
import io
import json
from datetime import datetime

import msoffcrypto
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

EXCEL_PASSWORD = '1989'

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

STORE_FOLDERS = {
    '쓰윗성수': {'id': 'seongsu',   'color': '#10B981'},
    '성수1층':  {'id': 'seongsu1f', 'color': '#6C63FF'},
    '쓰윗종각': {'id': 'jonggak',   'color': '#F472B6'},
}
CAT_ORDER = ['위스키 아이스크림', '논알콜 젤라또', '커피', '에이드', '디저트', '위스키']
DOW_NAMES = ['월', '화', '수', '목', '금', '토', '일']


def get_service():
    sa_info = json.loads(os.environ['GOOGLE_SERVICE_ACCOUNT_JSON'])
    creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)


def list_folder(service, folder_id):
    res = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields='files(id, name, mimeType)', pageSize=200,
    ).execute()
    return res.get('files', [])


def download_file(service, file_id):
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf


def decrypt_if_needed(buf):
    """비밀번호가 걸린 엑셀이면 복호화해서 반환, 아니면 원본 반환"""
    buf.seek(0)
    try:
        office_file = msoffcrypto.OfficeFile(buf)
        if office_file.is_encrypted():
            print(f'     🔐 암호화 파일 복호화 중 (비밀번호: {EXCEL_PASSWORD})...')
            office_file.load_key(password=EXCEL_PASSWORD)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            print(f'     ✅ 복호화 완료 ({decrypted.getbuffer().nbytes:,} bytes)')
            return decrypted
        else:
            print(f'     ℹ️  암호화 없음')
    except Exception as e:
        print(f'     ❌ 복호화 실패: {e}')
    buf.seek(0)
    return buf


def read_excel(buf, **kwargs):
    last_err = None
    for engine in ['openpyxl', 'xlrd', 'calamine']:
        try:
            buf.seek(0)
            return pd.read_excel(buf, engine=engine, **kwargs)
        except Exception as e:
            last_err = e
            continue
    raise ValueError(f'엑셀 파일을 읽을 수 없습니다. ({last_err})')


def get_sheet_names(buf):
    """엑셀 파일의 시트 이름 목록 반환"""
    for engine in ['openpyxl', 'xlrd', 'calamine']:
        try:
            buf.seek(0)
            xf = pd.ExcelFile(buf, engine=engine)
            return xf.sheet_names
        except Exception:
            continue
    return []


def find_sheet(sheet_names, keywords):
    """키워드를 포함하는 시트 이름 탐색 (대소문자/공백 무시)"""
    for name in sheet_names:
        if all(k in name for k in keywords):
            return name
    return None


def parse_pos_file(buf, store_name):
    sheets = get_sheet_names(buf)
    print(f'     📋 시트 목록: {sheets}')

    sheet_pay    = find_sheet(sheets, ['결제', '합계']) or '결제 합계'
    sheet_items  = find_sheet(sheets, ['상품', '주문', '합계']) or '상품 주문 합계'
    sheet_detail = find_sheet(sheets, ['결제', '상세']) or '결제 상세내역'
    sheet_item_detail = find_sheet(sheets, ['상품', '주문', '상세']) or '상품 주문 상세내역'
    print(f'     → 결제합계={sheet_pay}, 상품={sheet_items}, 상세={sheet_detail}')

    # ── 결제 상세내역 (모든 파일에서 잘 읽힘) ──────────────────────────
    df_detail = pd.DataFrame()
    try:
        df_detail = read_excel(buf, sheet_name=sheet_detail, header=0, skiprows=[1], usecols=[0, 1, 5, 7, 9])
        df_detail.columns = ['date', 'time', 'revenue', 'method', 'status']
        df_detail = df_detail.dropna(subset=['date']).copy()
        df_detail['revenue'] = pd.to_numeric(df_detail['revenue'], errors='coerce').fillna(0)
        df_detail['store'] = store_name
    except Exception as e:
        print(f'     ⚠️ 상세내역 파싱 실패: {e}')

    # ── 결제 합계 시트 시도 ───────────────────────────────────────────
    df_pay = pd.DataFrame()
    try:
        _tmp = read_excel(buf, sheet_name=sheet_pay, header=0, skiprows=[1], usecols=[0, 1, 3])
        _tmp.columns = ['date', 'revenue', 'tc']
        _tmp = _tmp.dropna(subset=['date']).copy()
        _tmp['date'] = pd.to_datetime(_tmp['date'], errors='coerce')
        _tmp = _tmp.dropna(subset=['date'])
        _tmp['revenue'] = pd.to_numeric(_tmp['revenue'], errors='coerce').fillna(0)
        _tmp['tc'] = pd.to_numeric(_tmp['tc'], errors='coerce').fillna(0)
        _tmp['store'] = store_name
        if len(_tmp) > 0:
            df_pay = _tmp
            print(f'     ✅ 결제합계 시트: {len(df_pay)}행')
    except Exception:
        pass

    # ── 결제합계 0행이면 상세내역에서 일별 집계 ──────────────────────
    if len(df_pay) == 0 and not df_detail.empty:
        print(f'     ⚠️ 결제합계 0행 → 상세내역 기반 일별 집계')
        df_a = df_detail.copy()
        if 'status' in df_a.columns:
            df_a = df_a[df_a['status'] == '승인']
        df_a['_date'] = pd.to_datetime(df_a['date'], errors='coerce').dt.normalize()
        daily = df_a.groupby('_date').agg(
            revenue=('revenue', 'sum'),
            tc=('revenue', 'count'),
        ).reset_index().rename(columns={'_date': 'date'})
        daily['store'] = store_name
        df_pay = daily
        print(f'     ✅ 일별 집계 완료: {len(df_pay)}행')

    # ── 상품 데이터 파싱 공통 함수 ────────────────────────────────────
    def parse_items_df(df):
        print(f'     📋 상품시트 컬럼: {list(df.columns[:8])}')
        # 컬럼명 정규화: 앞뒤 공백 제거
        df.columns = [str(c).strip() for c in df.columns]
        name_col = next((c for c in df.columns if '상품명' in c or '상품 명' in c), None)
        if not name_col:
            print(f'     ⚠️ 상품명 컬럼 없음')
            return pd.DataFrame()
        df = df.dropna(subset=[name_col]).copy()
        sales_col = next((c for c in df.columns if '실 판매' in c or '실판매' in c), None)
        if sales_col is None:
            sales_col = next((c for c in df.columns if '가격' in c), None)
        date_col  = next((c for c in df.columns if '기간' in c or '날짜' in c or 'date' in c.lower()), None)
        cat_col   = next((c for c in df.columns if '카테고리' in c), None)
        cnt_col   = next((c for c in df.columns if '판매건수' in c or '건수' in c), None)
        print(f'     → name={name_col}, date={date_col}, cat={cat_col}, cnt={cnt_col}, sales={sales_col}')
        if not all([date_col, cat_col, cnt_col, sales_col]):
            return pd.DataFrame()
        out = df[[date_col, name_col, cat_col, cnt_col, sales_col]].copy()
        out.columns = ['date', 'product', 'category', 'cnt', 'actual_price']
        out['date'] = pd.to_datetime(out['date'], errors='coerce')
        out = out.dropna(subset=['date'])
        out['cnt'] = pd.to_numeric(out['cnt'], errors='coerce').fillna(0)
        out['actual_price'] = pd.to_numeric(out['actual_price'], errors='coerce').fillna(0)
        out['store'] = store_name
        return out

    # ── 상품 주문 합계 (skiprows 없이 / 있이 둘 다 시도) ───────────────
    df_items = pd.DataFrame()
    for skip in [None, [1]]:
        try:
            kw = dict(sheet_name=sheet_items, header=0)
            if skip: kw['skiprows'] = skip
            _items = read_excel(buf, **kw)
            parsed = parse_items_df(_items)
            if len(parsed) > 0:
                df_items = parsed
                print(f'     ✅ 상품합계 시트: {len(df_items)}행 (skiprows={skip})')
                break
        except Exception as e:
            print(f'     ⚠️ 상품합계 파싱 실패(skip={skip}): {e}')

    # 상품합계도 0행이면 상품 주문 상세내역 시도
    if len(df_items) == 0:
        for skip in [None, [1]]:
            try:
                kw = dict(sheet_name=sheet_item_detail, header=0)
                if skip: kw['skiprows'] = skip
                _idet = read_excel(buf, **kw)
                parsed = parse_items_df(_idet)
                if len(parsed) > 0:
                    df_items = parsed
                    print(f'     ✅ 상품상세 시트 대체: {len(df_items)}행 (skiprows={skip})')
                    break
            except Exception:
                pass

    print(f'     📊 최종: 결제합계={len(df_pay)}행, 상품={len(df_items)}행, 상세={len(df_detail)}행')
    return df_pay, df_items, df_detail


def compute_detail(df_pay, df_items, df_detail):
    """카테고리 / 상품 / 시간대 / 요일별 지표 계산 (특정 매장 또는 전체 데이터 전달)"""
    # 카테고리별
    cat_grp = df_items.groupby('category').agg(
        revenue=('actual_price', 'sum'), cnt=('cnt', 'sum'),
    ).reset_index()
    total_cat = cat_grp['revenue'].sum()
    cat_grp['_ord'] = cat_grp['category'].map({c: i for i, c in enumerate(CAT_ORDER)}).fillna(99)
    cat_grp = cat_grp.sort_values('_ord').drop(columns='_ord')
    categories = [
        {'name': r['category'], 'revenue': int(r['revenue']),
         'cnt': int(r['cnt']), 'share': round(r['revenue'] / total_cat * 100, 1) if total_cat else 0}
        for _, r in cat_grp.iterrows()
    ]

    # 상품별
    prod_grp = df_items.groupby(['product', 'category']).agg(
        cnt=('cnt', 'sum'), revenue=('actual_price', 'sum'),
    ).reset_index().sort_values('cnt', ascending=False)
    products = [
        {'name': r['product'], 'category': r['category'],
         'cnt': int(r['cnt']), 'revenue': int(r['revenue'])}
        for _, r in prod_grp.iterrows()
    ]

    # 시간대별
    hourly = []
    if not df_detail.empty:
        df_a = df_detail[df_detail['status'] == '승인'].copy() if 'status' in df_detail.columns else df_detail.copy()
        df_a['hour'] = pd.to_datetime(df_a['time'], errors='coerce').dt.hour
        df_a['revenue'] = pd.to_numeric(df_a['revenue'], errors='coerce').fillna(0)
        hg = df_a.groupby('hour').agg(revenue=('revenue', 'sum'), tc=('revenue', 'count')).reset_index()
        hourly = [{'hour': int(r['hour']), 'revenue': int(r['revenue']), 'tc': int(r['tc'])}
                  for _, r in hg.sort_values('hour').iterrows()]

    # 요일별
    df_p = df_pay.copy()
    df_p['dow'] = df_p['date'].dt.dayofweek
    dg = df_p.groupby('dow').agg(total_rev=('revenue', 'sum'), days=('date', 'nunique')).reset_index()
    dow = [
        {'dow': int(r['dow']), 'name': DOW_NAMES[int(r['dow'])],
         'avg_revenue': round(r['total_rev'] / r['days']) if r['days'] > 0 else 0,
         'days': int(r['days'])}
        for _, r in dg.sort_values('dow').iterrows()
    ]

    return {'categories': categories, 'products': products, 'hourly': hourly, 'dow': dow}


def compute_metrics(df_pay, df_items, df_detail):
    start = df_pay['date'].min().strftime('%Y-%m-%d')
    end   = df_pay['date'].max().strftime('%Y-%m-%d')
    n_days = max((df_pay['date'].max() - df_pay['date'].min()).days + 1, 1)

    total_rev = int(df_pay['revenue'].sum())
    total_tc  = int(df_pay['tc'].sum())

    # 매장별 KPI
    stores_out = []
    for fname, info in STORE_FOLDERS.items():
        df = df_pay[df_pay['store'] == fname]
        if df.empty:
            continue
        rev = int(df['revenue'].sum())
        tc  = int(df['tc'].sum())
        n   = max((df['date'].max() - df['date'].min()).days + 1, 1)
        stores_out.append({
            'id': info['id'], 'name': fname, 'color': info['color'],
            'revenue': rev, 'tc': tc,
            'daily_avg': round(rev / n),
            'atv': round(rev / tc) if tc > 0 else 0,
            'share': round(rev / total_rev * 100, 1) if total_rev > 0 else 0,
        })

    # 일별
    daily_out = [
        {'date': r['date'].strftime('%Y-%m-%d'), 'store': r['store'],
         'revenue': int(r['revenue']), 'tc': int(r['tc'])}
        for _, r in df_pay.sort_values(['date', 'store']).iterrows()
    ]

    # by_store: 전체 + 각 매장별
    by_store = {'all': compute_detail(df_pay, df_items, df_detail)}
    for fname in STORE_FOLDERS:
        sp = df_pay[df_pay['store'] == fname]
        si = df_items[df_items['store'] == fname]
        sd = df_detail[df_detail['store'] == fname] if not df_detail.empty else pd.DataFrame()
        if not sp.empty:
            by_store[fname] = compute_detail(sp, si, sd)

    return {
        'updated_at': datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),
        'period': {'start': start, 'end': end},
        'summary': {
            'total_revenue': total_rev, 'total_tc': total_tc,
            'daily_avg': round(total_rev / n_days),
            'atv': round(total_rev / total_tc) if total_tc > 0 else 0,
        },
        'stores': stores_out,
        'daily': daily_out,
        'by_store': by_store,
    }


def main():
    print('🚀 Google Drive 데이터 수집 시작...')
    service   = get_service()
    folder_id = os.environ['GOOGLE_DRIVE_FOLDER_ID']

    all_pay, all_items, all_detail = [], [], []

    for folder in list_folder(service, folder_id):
        if folder['mimeType'] != 'application/vnd.google-apps.folder':
            continue
        fname = folder['name']
        if fname not in STORE_FOLDERS:
            print(f'  ⚠️  알 수 없는 폴더: {fname}')
            continue

        excels = [f for f in list_folder(service, folder['id'])
                  if f['name'].lower().endswith(('.xlsx', '.xls'))]
        if not excels:
            print(f'  ⚠️  {fname}: 엑셀 없음')
            continue

        print(f'  📂 {fname}: {len(excels)}개 파일')
        for f in excels:
            print(f'     ↳ {f["name"]}')
            buf = download_file(service, f['id'])
            buf = decrypt_if_needed(buf)
            try:
                pay, items, detail = parse_pos_file(buf, fname)
                print(f'     📊 결제합계: {len(pay)}행, 상품: {len(items)}행, 상세: {len(detail)}행')
                all_pay.append(pay)
                all_items.append(items)
                if not detail.empty:
                    all_detail.append(detail)
            except Exception as e:
                import traceback
                print(f'     ❌ {e}')
                print(traceback.format_exc())

    if not all_pay:
        raise RuntimeError('처리된 데이터가 없습니다.')

    df_pay    = pd.concat(all_pay, ignore_index=True)
    df_items  = pd.concat(all_items, ignore_index=True)
    df_detail = pd.concat(all_detail, ignore_index=True) if all_detail else pd.DataFrame()

    # 중복 날짜 제거 (동일 날짜 파일 중복 업로드 시 최신 우선)
    df_pay   = df_pay.sort_values('date').drop_duplicates(subset=['date', 'store'], keep='last')
    df_items = df_items.sort_values('date').drop_duplicates(subset=['date', 'product', 'store'], keep='last')

    print('📊 지표 계산 중...')
    result = compute_metrics(df_pay, df_items, df_detail)

    out = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'latest.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as fp:
        json.dump(result, fp, ensure_ascii=False, indent=2)

    print(f'✅ 완료: data/latest.json')
    print(f'   기간: {result["period"]["start"]} ~ {result["period"]["end"]}')
    print(f'   총 매출: {result["summary"]["total_revenue"]:,}원')


if __name__ == '__main__':
    main()
