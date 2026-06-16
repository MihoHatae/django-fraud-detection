import pandas as pd
import os
from datetime import datetime, timedelta

def run_excel_overwrite():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, 'final_transaction_history_v6.xlsx')
    
    if not os.path.exists(excel_path):
        excel_path = os.path.join(os.path.dirname(base_dir), 'final_transaction_history_v6.xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ Excelファイルが見つかりません。パスを確認してください: {excel_path}")
        return

    print(f"🔄 Excelファイルを読み込んでいます: {excel_path}")
    
    df = pd.read_excel(excel_path, dtype=str)
    df = df.fillna('')

    time_col = '購入時間(JST)' if '購入時間(JST)' in df.columns else '購入時間'
    
    df['購入日'] = df['購入日'].apply(lambda x: x.split(' ')[0].replace('-', '/') if ' ' in x else x.replace('-', '/'))

    # -------------------------------------------------------------------------
    # ① エクセルの全体を購入日と時間で「昇順（古い順）」に並び替え
    # -------------------------------------------------------------------------
    temp_datetime = pd.to_datetime(df['購入日'] + ' ' + df[time_col], errors='coerce')
    df['_sort_key'] = temp_datetime
    df = df.sort_values(by='_sort_key', ascending=True).reset_index(drop=True)
    df = df.drop(columns=['_sort_key'])

    # 計算用に金額を数値化
    df['_numeric_amount'] = pd.to_numeric(df['金額'], errors='coerce').fillna(0)

    # -------------------------------------------------------------------------
    # ② 平均単価の計算のし直し（取引発生時点での『過去の累計平均単価』を一行ずつ算出）
    # -------------------------------------------------------------------------
    print("② 取引発生日時の時点での平均単価を一行ずつ計算しています（FAC・チャージ除外）...")
    
    # 新しい平均単価を格納するためのリスト
    calculated_avg_prices = []
    
    # ユーザーごとに「過去の有効な購入金額のリスト」を記録するための辞書
    user_purchase_history = {}

    # 昇順に並んだデータを上から順（時系列順）に処理
    for idx, row in df.iterrows():
        user_name = row['名前']
        tx_id = str(row['取引ID'])
        payment_method = str(row['支払い方法'])
        current_amount = row['_numeric_amount']
        
        if user_name not in user_purchase_history:
            user_purchase_history[user_name] = []
            
        # 💡 [重要] この行が「アイテム購入」か「チャージ（除外対象）」かを判定
        is_charge = 'FAC.C' in tx_id or 'チャージ' in payment_method or 'ギフト' in payment_method or row['金額'] == ''
        
        if is_charge:
            # 該当行がチャージの場合は、平均単価の計算を汚さないよう「その時点の既存の平均単価」を入れる
            if len(user_purchase_history[user_name]) > 0:
                history = user_purchase_history[user_name]
                current_avg = sum(history) / len(history)
            else:
                current_avg = 0
        else:
            # 該当行が純粋なアイテム購入の場合、この現在の取引金額も含めてその時点の平均単価を計算
            user_purchase_history[user_name].append(current_amount)
            history = user_purchase_history[user_name]
            current_avg = sum(history) / len(history)
            
        # 計算結果（四捨五入して整数に変換）をリストに格納
        calculated_avg_prices.append(str(int(round(current_avg))))

    # 算出した時系列の平均単価をデータフレームに適用
    df['平均単価'] = calculated_avg_prices
    df = df.drop(columns=['_numeric_amount'])

    # -------------------------------------------------------------------------
    # ③ アプリのインストール日ロジックの適用
    # -------------------------------------------------------------------------
    grouped = df.groupby(['名前', 'アプリ名'])
    for (name, app_name), group_indices in grouped.groups.items():
        first_row_idx = group_indices[0]
        first_purchase_date_str = df.loc[first_row_idx, '購入日']
        
        try:
            first_purchase_date = datetime.strptime(first_purchase_date_str, '%Y/%m/%d')
            calculated_install_date = first_purchase_date - timedelta(days=1)
            install_date_str = calculated_install_date.strftime('%Y/%m/%d')
        except Exception:
            install_date_str = first_purchase_date_str
        
        df.loc[group_indices, 'インストール日'] = install_date_str

    # -------------------------------------------------------------------------
    # 💾 大元のExcelファイルへ「直接上書き保存」
    # -------------------------------------------------------------------------
    df.to_excel(excel_path, index=False)
    print("✨✨ 全員分のExcelデータの時系列クレンジングと直接書き換えが完璧に完了しました！ ✨✨")

if __name__ == '__main__':
    run_excel_overwrite()