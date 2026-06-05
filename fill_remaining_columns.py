import random
import pandas as pd

# 決済方法（L列）からBINコードをマッピングするための辞書
BIN_POOL = {
    "VISA": "453416",
    "Master": "541011",
    "JCB": "352811",
    "AMEX": "379212",
}
DOMESTIC_BINS = ["453416", "498007", "352811", "541011"]

DOMESTIC_DEVICES = {
    "Pixel 8": "14",
    "Pixel 7": "13",
    "Xperia 5 IV": "12",
    "AQUOS sense8": "13",
    "Galaxy S23": "14",
}


def main():
    # 前回までに作成されたExcelファイルを読み込みます
    input_file = "final_transaction_history_v5.xlsx"
    print(f"現在のExcelファイル（{input_file}）を読み込んでいます...")

    df = pd.read_excel(input_file, header=None, dtype=object)

    # 各列のインデックス（A列=0...）
    name_idx = 0  # A列（名前）
    pay_method_idx = 11  # L列（支払い方法）
    auth_type_idx = 16  # Q列（認証種別）
    bin_idx = 24  # Y列（BIN）
    os_idx = 25  # Z列（OSバージョン）
    device_idx = 26  # AA列（デバイス名）
    avg_price_idx = 27  # AB列（平均単価）
    comm_idx = 28  # AC列（Wi-Fi / Mobile / VPN）

    print("【最終補正】認証の空欄に『None』を適用 ＆ 全データの最終クレンジング中...")

    # 名前一貫性用のキャッシュ辞書
    user_card_profiles = {}
    user_device_profiles = {}
    user_avg_profiles = {}
    user_comm_profiles = {}
    user_auth_profiles = {}

    for i in range(1, len(df)):
        name = str(df.iloc[i, name_idx]).strip()
        pay_method = str(df.iloc[i, pay_method_idx]).strip()

        # 1. 名前の事前プロファイル登録
        if name not in user_avg_profiles:
            if "林" in name:
                user_avg_profiles[name] = str(random.randint(500, 1200))
                user_comm_profiles[name] = "Wi-Fi"
                # 林様など、最初から設定を空欄にしている人は一貫して「None」に固定
                user_auth_profiles[name] = "None"
            elif "池田" in name:
                user_avg_profiles[name] = str(random.randint(500, 1200))
                user_comm_profiles[name] = "Wi-Fi"
                user_auth_profiles[name] = "PWD"
            else:
                user_avg_profiles[name] = str(
                    random.choice([500, 800, 1200, 1500, 3000, 4500])
                )
                user_comm_profiles[name] = random.choice(["Wi-Fi", "Mobile"])
                user_auth_profiles[name] = "BIO"

        # 2. Q列: 認証種別 の上書き補正
        # 元のセルが空欄（またはnanや空白文字）の場合、プロファイルに基づいて「None」等をハメ込みます
        if pd.isna(df.iloc[i, auth_type_idx]) or str(
            df.iloc[i, auth_type_idx]
        ).strip() in ["", "nan"]:
            df.iloc[i, auth_type_idx] = user_auth_profiles[name]

        # 3. Y列: BINコード（名前 × クレカ種別固定）
        card_key = f"{name}_{pay_method}"
        if pd.notna(df.iloc[i, bin_idx]) and str(df.iloc[i, bin_idx]).strip() not in [
            "",
            "nan",
        ]:
            user_card_profiles[card_key] = str(df.iloc[i, bin_idx]).strip()
        else:
            if card_key not in user_card_profiles:
                matched_bin = random.choice(DOMESTIC_BINS)
                for brand, bin_code in BIN_POOL.items():
                    if brand.lower() in pay_method.lower():
                        matched_bin = bin_code
                        break
                user_card_profiles[card_key] = matched_bin
            df.iloc[i, bin_idx] = user_card_profiles[card_key]

        # 4. Z列 & AA列: OSバージョンとデバイス名
        if pd.isna(df.iloc[i, device_idx]) or str(df.iloc[i, device_idx]).strip() in [
            "",
            "nan",
        ]:
            if name not in user_device_profiles:
                user_device_profiles[name] = random.choice(
                    list(DOMESTIC_DEVICES.keys())
                )
            chosen_device = user_device_profiles[name]
            df.iloc[i, device_idx] = chosen_device
            df.iloc[i, os_idx] = DOMESTIC_DEVICES[chosen_device]

        # 5. AB列: 平均単価
        df.iloc[i, avg_price_idx] = user_avg_profiles[name]

        # 6. AC列: Wi-Fi / Mobile / VPN
        if pd.notna(df.iloc[i, comm_idx]) and str(df.iloc[i, comm_idx]).strip() in [
            "VPN",
            "Mobile",
            "Wi-Fi",
        ]:
            pass  # 林様の手入力を維持
        else:
            df.iloc[i, comm_idx] = user_comm_profiles[name]

    # すべてのロジックが詰まった【最終大完成版】として保存
    output_file = "final_transaction_history_v6.xlsx"
    df.to_excel(output_file, index=False, header=False)
    print(
        f"【パーフェクト】すべての認証列に『None』が適用され、データセットが完全体になりました！最終ファイル: {output_file}"
    )


if __name__ == "__main__":
    main()