from django.shortcuts import render
import pandas as pd
import os

def index(request):
    answer = request.GET.get('answer')
    
    # 5つのアカウント検索フォームの値を取得
    account_queries = [
        request.GET.get('account_query_1', ''),
        request.GET.get('account_query_2', ''),
        request.GET.get('account_query_3', ''),
        request.GET.get('account_query_4', ''),
        request.GET.get('account_query_5', '')
    ]
    
    # 1つ目のアカウント名を代表値として保持
    account_query = account_queries[0]

    context = {
        'select_answer': answer,
        'account_query': account_query,
        'account_queries': account_queries,
        'results': {}, 
    }

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(base_dir, 'final_transaction_history_v6.xlsx')
    
    if os.path.exists(excel_path):
        try:
            # Excelファイルをそのまま文字列として安全に読み込む
            df = pd.read_excel(excel_path, dtype=str)
            df = df.fillna('')
            
            # 【名称変更】列名を「通信環境」に変更
            if 'Wi-Fi / Mobile / VPN' in df.columns:
                df = df.rename(columns={'Wi-Fi / Mobile / VPN': '通信環境'})
            
            # 【表記補正】Excel内のヘッダー名を統一
            if '購入時間(JST)' in df.columns:
                df = df.rename(columns={'購入時間(JST)': '購入時間'})

            # 💡【修正箇所】インストール日から「 00:00:00」などの余計な時間をカットして日付だけに揃える
            if 'インストール日' in df.columns:
                df['インストール日'] = df['インストール日'].apply(lambda x: x.split(' ')[0].replace('-', '/') if ' ' in x else x.replace('-', '/'))

            # 🍏 【追加】購入日からも「 00:00:00」を完全に切り落としてスラッシュ区切りの日付に統一する
            if '購入日' in df.columns:
                df['購入日'] = df['購入日'].apply(lambda x: x.split(' ')[0].replace('-', '/') if ' ' in x else x.replace('-', '/'))

            # 指定された30列の順番に整理
            defined_order = [
                '名前', 'アカウント', 'アカウント数', '購入日', '購入時間', 
                '取引ID', 'アプリ名', 'アイテム名', '金額', '支払い方法', 
                'BIN', '決済状況', '居住地', '所在地', 'IPアドレス', 
                '年齢', '対象年齢', 'キャリア', 'デバイス名', 'IMEI', 
                '端末数', '通信環境', 'OSバージョン', '認証種別', '購入時認証', 
                '平均単価', 'インストール日', 'アラート', '調査依頼回数', '過去FF調査依頼回数'
            ]
            existing_columns = [col for col in defined_order if col in df.columns]
            df = df[existing_columns]

            # 出力用に、元の「購入時間(JST)」というヘッダー名に戻す
            if '購入時間' in df.columns:
                df = df.rename(columns={'購入時間': '購入時間(JST)'})

            # アカウントごとにデータを抽出して画面に送る
            for i, query in enumerate(account_queries, start=1):
                if query: 
                    is_target = df['アカウント'].str.contains(query, case=False, na=False)
                    df_target = df[is_target]
                    context['results'][f'account_{i}'] = df_target.to_dict(orient='records')
                else:
                    context['results'][f'account_{i}'] = df.to_dict(orient='records')
                    
        except Exception as e:
            context['excel_error'] = f"データ処理中にエラーが発生しました: {e}"
    else:
        context['excel_error'] = f"Excelファイルが見つかりません。パス: {excel_path}"

    return render(request, 'fraud_alerts/index.html', context)