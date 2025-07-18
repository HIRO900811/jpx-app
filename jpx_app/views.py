import requests
from django.shortcuts import render
from .utils import get_latest_excel_url, extract_specific_cells_from_excel, get_latest_excel_url, fetch_nikkei225_from_api, analyze_market_trend

def index(request):
    excel_url = get_latest_excel_url()
    if not excel_url:
        return render(request, 'index.html', {'error': 'Excelファイルが見つかりません'})

    res = requests.get(excel_url)
    extracted_data = extract_specific_cells_from_excel(res.content)

    return render(request, 'index.html', {'extracted': extracted_data})

#JPXのリアルタイムデータAPIから「日経225」の先物データを取得する
def index(request):
    excel_url = get_latest_excel_url()
    if not excel_url:
        return render(request, 'index.html', {'error': 'Excelファイルが見つかりません'})

    res = requests.get(excel_url)
    extracted_data = extract_specific_cells_from_excel(res.content)
    nikkei_data = fetch_nikkei225_from_api()  # ← 追加

    return render(request, 'index.html', {
        'extracted': extracted_data,
        'nikkei': nikkei_data  # ← 追加
    })

#出力データの分析
def index(request):
    excel_url = get_latest_excel_url()
    if not excel_url:
        return render(request, 'index.html', {'error': 'Excelファイルが見つかりません'})

    res = requests.get(excel_url)
    extracted_data = extract_specific_cells_from_excel(res.content)
    nikkei_data = fetch_nikkei225_from_api()
    analysis_result = analyze_market_trend(nikkei_data, extracted_data)  # ← 追加

    return render(request, 'index.html', {
        'extracted': extracted_data,
        'nikkei': nikkei_data,
        'analysis': analysis_result  # ← 追加
    })
