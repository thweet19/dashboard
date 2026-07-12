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
        fields='files(id, name, mimeType, modifiedTime)', pageSize=200,
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


# 기대 시트명 키워드 (find_sheet 기준)
REQUIRED_SHEETS = [
    (['결제', '합계'],    '결제 합계'),
    (['결제', '상세'],    '결제 상세내역'),
    (['상품', '주문', '상세'], '상품 주문 상세내역'),
]
# 상품 주문 상세내역 기대 컬럼 (위치 → 포함 키워드)
EXPECTED_ITEM_COLS = {
    0:  ['일자'],
    1:  ['상태'],
    4:  ['번호'],
    5:  ['상품명', '상품'],
    7:  ['카테고리'],
    11: ['수량'],
    16: ['실판매금액', '판매금액'],
}


def validate_excel_structure(buf, store_name, file_name):
    """엑셀 파일 구조 검증. 이상 있으면 warning dict 리스트 반환."""
    issues = []

    def warn(detail, severity='error'):
        issues.append({
            'store': store_name,
            'file': file_name,
            'message': '매출 엑셀파일 구조가 상이합니다',
            'detail': detail,
            'severity': severity,
        })

    sheets = get_sheet_names(buf)

    # 1. 필수 시트 존재 확인
    for keywords, label in REQUIRED_SHEETS:
        if not find_sheet(sheets, keywords):
            warn(f"필수 시트 '{label}'를 찾을 수 없습니다 (발견된 시트: {', '.join(sheets)})")

    # 2. 상품 주문 상세내역 컬럼 구조 확인
    sheet_item_detail = find_sheet(sheets, ['상품', '주문', '상세'])
    if sheet_item_detail:
        try:
            hdr = read_excel(buf, sheet_name=sheet_item_detail, header=0, nrows=0)
            cols = list(hdr.columns)
            total = len(cols)
            for pos, keywords in EXPECTED_ITEM_COLS.items():
                if pos >= total:
                    warn(f"상품 주문 상세내역: 컬럼 수({total}개)가 부족합니다 — {pos+1}번째 컬럼({keywords[0]}) 없음")
                else:
                    col_name = str(cols[pos])
                    if not any(kw in col_name for kw in keywords):
                        warn(
                            f"상품 주문 상세내역 {pos+1}번째 컬럼: "
                            f"'{col_name}' (기대값: {'/'.join(keywords)})",
                            severity='warning',
                        )
        except Exception as e:
            warn(f"상품 주문 상세내역 시트 열기 실패: {e}")

    return issues


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
        detail_total = int(df_a['revenue'].sum())
        print(f'     ✅ 일별 집계 완료: {len(df_pay)}행, 결제상세내역 총합={detail_total:,}원')

    # ── 상품 주문 상세내역 RAW 데이터로 파싱 (모든 파일 공통) ──────────
    # 컬럼: A(0)주문기준일자, B(1)결제상태, E(4)주문번호, F(5)상품명, H(7)카테고리,
    #       L(11)수량, Q(16)실판매금액(할인+옵션 포함)
    df_items = pd.DataFrame()
    try:
        _raw = read_excel(buf, sheet_name=sheet_item_detail, header=0, skiprows=[1],
                          usecols=[0, 1, 4, 5, 7, 11, 16])
        _raw.columns = ['date', 'status', 'order_id', 'product', 'category', 'cnt', 'actual_price']
        statuses = _raw['status'].astype(str).str.strip().unique()
        print(f'     ℹ️  상품 상태값: {list(statuses)}')
        cancel_kw = {'취소', '환불', '부분취소', '결제취소', 'cancel'}
        _raw = _raw[~_raw['status'].astype(str).str.strip().str.lower().isin({s.lower() for s in cancel_kw})].copy()
        _raw['date'] = pd.to_datetime(_raw['date'], errors='coerce')
        _raw = _raw.dropna(subset=['date', 'product', 'category'])
        _raw['cnt'] = pd.to_numeric(_raw['cnt'], errors='coerce').fillna(1)
        _raw['actual_price'] = pd.to_numeric(_raw['actual_price'], errors='coerce').fillna(0)
        _raw['store'] = store_name
        df_items = _raw[['date', 'order_id', 'product', 'category', 'cnt', 'actual_price', 'store']].copy()
        cat_sum = df_items.groupby('category')['actual_price'].sum().sort_values(ascending=False)
        print(f'     📊 카테고리별 매출합계(실판매금액 기준):\n{cat_sum.to_string()}')
        print(f'     ✅ 상품 상세내역: {len(df_items)}행, 총합={int(df_items["actual_price"].sum()):,}원')
    except Exception as e:
        import traceback
        print(f'     ❌ 상품 상세내역 파싱 실패: {e}')
        print(traceback.format_exc())

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


def compute_monthly(df_pay, df_items, df_detail):
    """월별 stores + by_store 집계"""
    dp = df_pay.copy()
    di = df_items.copy()
    dp['_month'] = dp['date'].dt.strftime('%Y-%m')
    di['_month'] = pd.to_datetime(di['date'], errors='coerce').dt.strftime('%Y-%m')

    months = sorted(dp['_month'].unique())
    by_month = {}
    for month in months:
        mp = dp[dp['_month'] == month]
        mi = di[di['_month'] == month]
        md = pd.DataFrame()
        if not df_detail.empty:
            try:
                dd = df_detail.copy()
                dd['_month'] = pd.to_datetime(dd['date'], errors='coerce').dt.strftime('%Y-%m')
                md = dd[dd['_month'] == month].drop(columns='_month')
            except Exception:
                pass

        total_rev = int(mp['revenue'].sum())
        stores_m = []
        for fname, info in STORE_FOLDERS.items():
            df_s = mp[mp['store'] == fname]
            if df_s.empty:
                continue
            rev = int(df_s['revenue'].sum())
            tc  = int(df_s['tc'].sum())
            n   = max((df_s['date'].max() - df_s['date'].min()).days + 1, 1)
            stores_m.append({
                'id': info['id'], 'name': fname, 'color': info['color'],
                'revenue': rev, 'tc': tc,
                'daily_avg': round(rev / n),
                'atv': round(rev / tc) if tc > 0 else 0,
                'share': round(rev / total_rev * 100, 1) if total_rev > 0 else 0,
            })

        bstore_m = {'all': compute_detail(mp, mi, md)}
        for fname in STORE_FOLDERS:
            sp = mp[mp['store'] == fname]
            si = mi[mi['store'] == fname]
            sd = md[md['store'] == fname] if not md.empty else pd.DataFrame()
            if not sp.empty:
                bstore_m[fname] = compute_detail(sp, si, sd)

        by_month[month] = {'stores': stores_m, 'by_store': bstore_m}

    return by_month


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
        'by_month': compute_monthly(df_pay, df_items, df_detail),
    }


def main():
    print('🚀 Google Drive 데이터 수집 시작...')
    service   = get_service()
    folder_id = os.environ['GOOGLE_DRIVE_FOLDER_ID']

    all_pay, all_items, all_detail = [], [], []
    all_warnings = []
    file_order = 0

    for folder in list_folder(service, folder_id):
        if folder['mimeType'] != 'application/vnd.google-apps.folder':
            continue
        fname = folder['name']
        if fname not in STORE_FOLDERS:
            print(f'  ⚠️  알 수 없는 폴더: {fname}')
            continue

        # modifiedTime 오름차순 정렬 → 오래된 파일 먼저, 최신 파일 나중 (최신이 dedup에서 이김)
        excels = sorted(
            [f for f in list_folder(service, folder['id']) if f['name'].lower().endswith(('.xlsx', '.xls'))],
            key=lambda f: f.get('modifiedTime', '')
        )
        if not excels:
            print(f'  ⚠️  {fname}: 엑셀 없음')
            continue

        print(f'  📂 {fname}: {len(excels)}개 파일 (업로드순 처리)')
        for f in excels:
            print(f'     ↳ {f["name"]} (수정: {f.get("modifiedTime","?")[:10]})')
            buf = download_file(service, f['id'])
            buf = decrypt_if_needed(buf)
            try:
                # 구조 검증 먼저
                file_warnings = validate_excel_structure(buf, fname, f['name'])
                if file_warnings:
                    for w in file_warnings:
                        sev = '❌' if w['severity'] == 'error' else '⚠️'
                        print(f'     {sev} 구조경고: {w["detail"]}')
                    all_warnings.extend(file_warnings)

                pay, items, detail = parse_pos_file(buf, fname)
                print(f'     📊 결제합계: {len(pay)}행, 상품: {len(items)}행, 상세: {len(detail)}행')
                if not pay.empty:
                    pay['_forder'] = file_order
                if not items.empty:
                    items['_forder'] = file_order
                file_order += 1
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

    # 최신 업로드 파일 우선 덮어쓰기 (_forder 높을수록 최신)
    if '_forder' not in df_pay.columns:
        df_pay['_forder'] = 0
    if '_forder' not in df_items.columns:
        df_items['_forder'] = 0
    df_pay   = df_pay.sort_values(['date', '_forder']).drop_duplicates(subset=['date', 'store'], keep='last').drop(columns='_forder')
    # items dedup: 주문번호+상품+매장 기준 → 파일 중복 방지하면서 같은 날 같은 상품 다건 주문은 유지
    df_items = df_items.sort_values(['date', '_forder']).drop_duplicates(subset=['order_id', 'product', 'store'], keep='last').drop(columns='_forder')

    print('📊 지표 계산 중...')
    result = compute_metrics(df_pay, df_items, df_detail)
    result['warnings'] = all_warnings
    if all_warnings:
        errors = [w for w in all_warnings if w['severity'] == 'error']
        print(f'  ⚠️  구조 경고 {len(all_warnings)}건 (오류 {len(errors)}건) → JSON에 포함됨')

    out = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'latest.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as fp:
        json.dump(result, fp, ensure_ascii=False, indent=2)

    print(f'✅ 완료: data/latest.json')
    print(f'   기간: {result["period"]["start"]} ~ {result["period"]["end"]}')
    print(f'   총 매출: {result["summary"]["total_revenue"]:,}원')


if __name__ == '__main__':
    main()
