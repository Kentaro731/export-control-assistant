# 輸出貿易管理令 法令データ
# 経済産業省 外為法・輸出管理令に基づく参照データ
# 更新時は経産省通知に従い手動で修正すること

# 輸出令別表第3（ホワイト国リスト）— YESなら輸出OK（許可不要）
# 2023年時点。経産省通知で随時更新。
WHITE_COUNTRIES = {
    "アメリカ", "米国", "USA", "アメリカ合衆国",
    "イギリス", "英国", "UK",
    "ドイツ", "フランス", "イタリア", "スペイン", "オランダ",
    "ベルギー", "ルクセンブルク", "デンマーク", "スウェーデン",
    "フィンランド", "ノルウェー", "アイスランド", "ポルトガル",
    "オーストリア", "スイス", "アイルランド", "ギリシャ",
    "チェコ", "ハンガリー", "ポーランド", "スロバキア", "スロベニア",
    "エストニア", "ラトビア", "リトアニア", "マルタ", "キプロス",
    "ブルガリア", "ルーマニア", "クロアチア",
    "オーストラリア", "豪州", "ニュージーランド",
    "カナダ",
    "韓国", "大韓民国",
    "アルゼンチン",
}

WHITE_COUNTRIES_EN = {
    "United States", "USA", "US", "UK", "United Kingdom",
    "Germany", "France", "Italy", "Spain", "Netherlands",
    "Belgium", "Denmark", "Sweden", "Finland", "Norway",
    "Iceland", "Portugal", "Austria", "Switzerland", "Ireland",
    "Greece", "Czech Republic", "Hungary", "Poland", "Slovakia",
    "Slovenia", "Estonia", "Latvia", "Lithuania", "Malta",
    "Cyprus", "Bulgaria", "Romania", "Croatia", "Luxembourg",
    "Australia", "New Zealand", "Canada", "South Korea",
    "Republic of Korea", "Argentina",
}

# 輸出令別表第3の2（指定国・要注意国）— 通常兵器キャッチオール適用国
DESIGNATED_COUNTRIES = {
    "北朝鮮", "朝鮮民主主義人民共和国", "DPRK",
    "イラン", "Iran",
    "イラク", "Iraq",
    "リビア", "Libya",
    "コンゴ民主共和国", "DRC",
    "ソマリア", "Somalia",
    "スーダン", "Sudan",
    "南スーダン", "South Sudan",
    "中央アフリカ共和国",
    "ベラルーシ", "Belarus",
    "ミャンマー", "Myanmar",
    "ハイチ", "Haiti",
    "マリ", "Mali",
}

# 輸出令別表第1（1〜15項）主要規制カテゴリ
# 金属材料・プレートは通常これらに「非該当」
CONTROLLED_CATEGORIES_1_TO_15 = [
    "1項: 武器（銃砲弾薬、火薬、軍用爆発物）",
    "2項: 軍用火工品",
    "3項: 化学兵器関連（毒ガス等）",
    "4項: 生物兵器関連（病原体等）",
    "5項: 核兵器・核物質（ウラン、プルトニウム等）",
    "6項: 推進薬・爆薬の製造設備",
    "7項: 軍用爆発物の製造設備",
    "8項: 軍用ポンプ",
    "9項: 航空機用エンジン",
    "10項: 航空機・宇宙機器",
    "11項: 軍用艦船",
    "12項: 戦車・装甲車両",
    "13項: 軍用電子機器",
    "14項: 軍用誘導装置",
    "15項: 工作機械（精密度が一定水準以上のもの）",
]

# 16項（キャッチオール規制対象）の判定に用いる用途危険フラグ
# 主に大量破壊兵器（核・化学・生物・ミサイル）関連
CATCHALL_WARNING_KEYWORDS = [
    "核", "原子力", "ウラン", "プルトニウム",
    "化学兵器", "毒ガス", "神経剤",
    "生物兵器", "病原体", "細菌",
    "ミサイル", "弾道", "ロケット推進",
    "爆発物", "爆弾", "軍用",
    "兵器", "武器", "軍事",
]

# 通常兵器・軍事用途（リスト規制：別表第1の1〜15項）に直結しうる用途キーワード
# 用途にこれらが「明示」されている場合は、ホワイト国であっても
# 別表第1への該当性を厳格に評価し、輸出許可の要否を確認する必要がある。
MILITARY_END_USE_KEYWORDS = [
    "軍用", "軍事", "軍需", "国防", "防衛省", "自衛隊", "軍",
    "兵器", "武器", "弾薬", "弾頭",
    "戦闘機", "軍用機", "軍用航空機", "戦闘ヘリ", "無人機", "ドローン兵器",
    "戦車", "装甲車", "装甲", "砲", "火器", "銃",
    "潜水艦", "軍艦", "艦艇", "魚雷",
    "ミサイル", "弾道", "誘導弾", "ロケット弾",
    "レーダー", "軍事転用", "デュアルユース", "軍民融合",
]


def has_military_end_use(text: str) -> list:
    """テキストに通常兵器・軍事用途キーワードが含まれているか確認。
    検出された場合、ホワイト国向けでも別表第1該当性の精査が必要。"""
    found = []
    for kw in MILITARY_END_USE_KEYWORDS:
        if kw in text and kw not in found:
            found.append(kw)
    return found

# 取扱品目（金属材料）の一般的なHS分類（参考）
METAL_MATERIAL_HS = {
    "アルミ": "7604〜7616（アルミニウム製品）",
    "アルミニウム": "7604〜7616（アルミニウム製品）",
    "銅": "7407〜7419（銅製品）",
    "真鍮": "7407〜7419（銅合金製品）",
    "チタン": "8108（チタン及びその製品）",
    "ステンレス": "7219〜7228（鉄鋼製品）",
    "鉄鋼": "7200系（鉄及び鋼）",
    "ニッケル": "7505〜7508（ニッケル製品）",
    "タングステン": "8101（タングステン及びその製品）",
    "モリブデン": "8102（モリブデン及びその製品）",
}

def is_white_country(country: str) -> bool:
    """仕向け国がホワイト国（別表第3）に該当するか"""
    country_stripped = country.strip()
    return (country_stripped in WHITE_COUNTRIES or
            country_stripped in WHITE_COUNTRIES_EN)

def is_designated_country(country: str) -> bool:
    """仕向け国が別表第3の2（指定国）に該当するか"""
    country_stripped = country.strip()
    return country_stripped in DESIGNATED_COUNTRIES

def has_danger_keywords(text: str) -> list:
    """テキストに危険用途キーワードが含まれているか確認"""
    found = []
    for kw in CATCHALL_WARNING_KEYWORDS:
        if kw in text:
            found.append(kw)
    return found

def get_hs_hint(material: str) -> str:
    """材質からHS分類のヒントを取得"""
    for key, code in METAL_MATERIAL_HS.items():
        if key in material:
            return code
    return "7200〜8199番台（金属製品）— 詳細は品目により確認要"
