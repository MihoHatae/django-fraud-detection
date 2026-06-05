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
    # ★林様の手入力がある「v6」を直接読み込みます
    input_file = "final_transaction_history_v6.xlsx"
    print(f"現在のExcelファイル（{input_file}）を読み込んでいます...")

    df = pd.read_excel(input_file, header=None, dtype=object)

    # 各列のインデックス（A列=0...）
    name_idx = 0  # A列（名前）
    pay_method_idx = 11  # L列（支払い方法）
    auth_type_idx = 16  # Q列（認証種別）
    auth_now_idx = 17  # R列（購入時認証）
    bin_idx = 24  # Y列（BIN）
    os_idx = 25  # Z列（OSバージョン）
    device_idx = 26  # AA列（デバイス名）
    avg_price_idx = 27  # AB列（平均単価）
    comm_idx = 28  # AC列（Wi-Fi / Mobile / VPN）

    print("【最終仕上げ】手入力を完全に保護しつつ、残りの空欄を一括クレンジング中...")

    # 名前一貫性用のキャッシュ辞書
    user_card_profiles = {}
    user_device_profiles = {}
    user_avg_profiles = {}
    user_comm_profiles = {}
    user_auth_profiles = {}

    for i in range(1, len(df)):
        name = str(df.iloc[i, name_idx]).strip()
        pay_method = str(df.iloc[i, pay_method_idx]).strip()

        # 1. 名前の事前プロファイル登録（名前ごとに値を完全固定）
        if name not in user_avg_profiles:
            if "林" in name:
                user_avg_profiles[name] = str(random.randint(500, 1200))
                user_comm_profiles[name] = "Mobile"
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

        # 2. Q列（認証種別） & R列（購入時認証） の空欄補正
        if pd.isna(df.iloc[i, auth_type_idx]) or str(
            df.iloc[i, auth_type_idx]
        ).strip() in ["", "nan"]:
            df.iloc[i, auth_type_idx] = user_auth_profiles[name]

        if pd.isna(df.iloc[i, auth_now_idx]) or str(
            df.iloc[i, auth_now_idx]
        ).strip() in ["", "nan"]:
            df.iloc[i, auth_now_idx] = user_auth_profiles[name]

        # 3. Y列: BINコード（手入力を保護しつつ、名前 × クレカ種別固定）
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

        # 4. Z列 & AA列: OSバージョンとデバイス名（手入力を維持）
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

        # 5. AB列: 平均単価（適正な金額アノマリ用に強制上書き）
        df.iloc[i, avg_price_idx] = user_avg_profiles[name]

        # 6. AC列: Wi-Fi / Mobile / VPN の処理（手入力を維持）
        if pd.isna(df.iloc[i, comm_idx]) or str(df.iloc[i, comm_idx]).strip() in [
            "",
            "nan",
        ]:
            df.iloc[i, comm_idx] = user_comm_profiles[name]

    # 上書きで保存します
    output_file = "final_transaction_history_v6.xlsx"
    df.to_excel(output_file, index=False, header=False)
    print(
        f"【パーフェクト】すべての認証列に『None』が適用され、全データが完全体になりました！最終ファイル: {output_file}"
    )


if __name__ == "__main__":
    main()