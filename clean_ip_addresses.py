import random
import pandas as pd

# リアルなIPプール
IP_POOLS = {
    "NTT DOCOMO": ["49.98.132.45", "49.105.92.114", "114.162.7.88", "114.180.64.103"],
    "KDDI": ["1.66.12.45", "106.128.5.14", "106.130.33.67", "111.107.4.55"],
    "SoftBank": [
        "126.142.45.101",
        "126.158.201.34",
        "126.247.12.89",
        "126.255.90.142",
    ],
    "Rakuten Mobile": [
        "133.106.12.45",
        "133.106.99.182",
        "133.32.128.14",
        "115.124.33.88",
    ],
    "Wi-Fi": [
        "153.142.12.45",
        "153.150.33.64",
        "114.160.15.42",
        "121.80.12.98",
        "60.64.4.19",
    ],
    "VPN": ["160.86.231.42", "160.86.200.15", "160.86.231.43", "160.86.200.16"],
    "Korea": ["117.111.35.78", "117.111.90.23", "211.234.50.112", "211.36.133.45"],
}


def determine_correct_ip(row, carrier_idx, location_idx, comm_idx):
    # 位置（番号）で直接データを安全に取得
    carrier = str(row.iloc[carrier_idx]).strip()
    location = str(row.iloc[location_idx]).strip()
    comm_type = str(row.iloc[comm_idx]).strip()

    if "KR" in location or "Korea" in location or "韓国" in location:
        return random.choice(IP_POOLS["Korea"])
    if "VPN" in comm_type:
        return random.choice(IP_POOLS["VPN"])
    if "Wi-Fi" in comm_type or "Wi-Fi" in comm_type:
        return random.choice(IP_POOLS["Wi-Fi"])

    if "Mobile" in comm_type or "モバイル" in comm_type:
        if "DOCOMO" in carrier or "ドコモ" in carrier or "NTT" in carrier:
            return random.choice(IP_POOLS["NTT DOCOMO"])
        elif "KDDI" in carrier or "au" in carrier:
            return random.choice(IP_POOLS["KDDI"])
        elif "SoftBank" in carrier or "ソフトバンク" in carrier:
            return random.choice(IP_POOLS["SoftBank"])
        elif "Rakuten" in carrier or "楽天" in carrier:
            return random.choice(IP_POOLS["Rakuten Mobile"])

    return random.choice(IP_POOLS["Wi-Fi"])


def main():
    excel_file = "取引履歴.xlsx"
    print("Excelファイルを安全に読み込んでいます...")

    # header=Noneにすることで、文字の区切りトラブルを無視して全行を強制取得
    df = pd.read_excel(excel_file, header=None)

    # いただいた1行目の並び順から、列の「左からの位置（番号）」を絶対固定で定義します
    # Pythonは0から数えるので、A列=0, B列=1... となります
    carrier_idx = 7  # H列（キャリア）
    location_idx = 19  # T列（所在地）
    ip_idx = 20  # U列（IPアドレス）
    comm_idx = 28  # AC列（Wi-Fi / Mobile / VPN）

    print("インデックス位置でIPアドレスの条件判定と書き換えを実行中...")

    # 2行目（インデックス1）以降のデータ部分をコピーして処理
    data_df = df.iloc[1:].copy()

    # 安全にIPアドレスの列（20番目）を上書き
    data_df[ip_idx] = data_df.apply(
        lambda row: determine_correct_ip(
            row, carrier_idx, location_idx, comm_idx
        ),
        axis=1,
    )

    # 1行目のタイトルと、綺麗にした2行目以降のデータを結合
    final_df = pd.concat([df.iloc[:1], data_df])

    output_file = "取引履歴_新.xlsx"
    final_df.to_excel(output_file, index=False, header=False)
    print(f"【成功】IPアドレスの修正が完了しました！新ファイル: {output_file}")


if __name__ == "__main__":
    main()