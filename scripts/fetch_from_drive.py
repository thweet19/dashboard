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
    # 시트 이름 확인
    sheets = get_sheet_names(buf)
    print(f'     📋 시트 목록: {sheets}')

    sheet_pay   = find_sheet(sheets, ['결제', '합계']) or '결제 합계'
    sheet_items = find_sheet(sheets, ['상품']) or '상품 주문 합계'
    sheet_detail = find_sheet(sheets, ['결제', '상세']) or '결제 상세내역'
    print(f'     → 결제합계={sheet_pay}, 상품={sheet_items}, 상세={sheet_detail}')

    df_pay = read_excel(buf, sheet_name=sheet_pay, header=0, skiprows=[1], usecols=[0, 1, 3])
    df_pay.columns = ['date', 'revenue', 'tc']
    df_pay = df_pay.dropna(subset=['date']).copy()
    df_pay['date'] = pd.to_datetime(df_pay['date'], errors='coerce')
    df_pay = df_pay.dropna(subset=['date'])
    df_pay['revenue'] = pd.to_numeric(df_pay['revenue'], errors='coerce').fillna(0)
    df_pay['tc'] = pd.to_numeric(df_pay['tc'], errors='coerce').fillna(0)
    df_pay['store'] = store_name

    df_items = read_excel(buf, sheet_name=sheet_items, header=0)
    df_items = df_items.dropna(subset=['상품명']).copy()
    sales_col = next((c for c in df_items.columns if '실 판매' in str(c)), '상품가격')
    df_items = df_items[['기간', '상품명', '카테고리', '판매건수', sales_col]].copy()
    df_items.columns = ['date', 'product', 'category', 'cnt', 'actual_price']
    df_items['date'] = pd.to_datetime(df_items['date'], errors='coerce')
    df_items = df_items.dropna(subset=['date'])
    df_items['cnt'] = pd.to_numeric(df_items['cnt'], errors='coerce').fillna(0)
    df_items['actual_price'] = pd.to_numeric(df_items['actual_price'], errors='coerce').fillna(0)
    df_items['store'] = store_name

    try:
        df_detail = read_excel(buf, sheet_name=sheet_detail, header=0, skiprows=[1], usecols=[0, 1, 5, 7, 9])
        df_detail.columns = ['date', 'time', 'revenue', 'method', 'status']
        df_detail = df_detail.dropna(subset=['date']).copy()
        df_detail['store'] = store_name
    except Exception:
        df_detail = pd.DataFrame()

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
