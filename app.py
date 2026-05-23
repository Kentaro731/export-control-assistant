import os
import json
from datetime import datetime
from flask import (
    Flask, render_template, jsonify, request,
    session, redirect, url_for,
)
import anthropic
from export_law_data import (
    is_white_country, is_designated_country,
    has_danger_keywords, has_military_end_use, get_hs_hint,
    CONTROLLED_CATEGORIES_1_TO_15,
)

app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# === 簡易パスワード認証（社内共有用）===
# 環境変数 APP_PASSWORD が設定されている場合のみ認証を有効化する。
# （ローカル開発で未設定なら従来どおりパスワードなしで動作）
app.secret_key = os.environ.get("APP_SECRET_KEY") or os.urandom(24).hex()
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

LOGIN_PAGE = """<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>輸出判定ツール ログイン</title>
<style>
  body{margin:0;height:100vh;display:flex;align-items:center;justify-content:center;
       background:#0f1419;color:#e6e6e6;font-family:"Yu Gothic",sans-serif;}
  .box{background:#1a2230;padding:40px 36px;border-radius:14px;width:320px;
       box-shadow:0 8px 30px rgba(0,0,0,.4);}
  h1{font-size:18px;margin:0 0 4px;}
  p.sub{font-size:12px;color:#8b9bb4;margin:0 0 24px;}
  input{width:100%;box-sizing:border-box;padding:12px;border-radius:8px;border:1px solid #34425a;
        background:#0f1419;color:#fff;font-size:15px;margin-bottom:14px;}
  button{width:100%;padding:12px;border:0;border-radius:8px;background:#2f6df6;color:#fff;
         font-size:15px;font-weight:bold;cursor:pointer;}
  .err{color:#ff6b6b;font-size:13px;margin-bottom:12px;}
  .note{font-size:11px;color:#6b7a93;margin-top:18px;line-height:1.6;}
</style></head>
<body><form class="box" method="post">
  <h1>輸出判定ツール</h1>
  <p class="sub">社内利用者専用</p>
  {ERR}
  <input type="password" name="password" placeholder="合言葉（パスワード）" autofocus>
  <button type="submit">ログイン</button>
  <p class="note">本ツールはAIによる一次確認の補助です。最終的な該非判断は輸出管理担当者が行ってください。</p>
</form></body></html>"""


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            return redirect(url_for("index"))
        return LOGIN_PAGE.replace("{ERR}", '<p class="err">パスワードが違います。</p>')
    return LOGIN_PAGE.replace("{ERR}", "")


@app.route("/logout")
def logout():
    session.pop("authed", None)
    return redirect(url_for("login"))


@app.before_request
def require_login():
    # パスワード未設定時は認証なし（ローカル開発用）
    if not APP_PASSWORD:
        return None
    # ログイン画面と静的ファイルは認証不要
    if request.endpoint in ("login", "static"):
        return None
    if not session.get("authed"):
        return redirect(url_for("login"))
    return None


HISTORY_FILE = os.path.join(os.path.dirname(__file__), "judgment_history.json")

SYSTEM_PROMPT = """あなたは日本の安全保障輸出管理（外為法・輸出貿易管理令）の専門家AIです。
非鉄金属卸（アルミ・銅・真鍮・チタン等の材料プレート・切り板を扱う企業）の
輸出担当者を支援します。

## 判定フロー（5ステップ）

STEP1: 輸出令別表第1の1〜15項への該当確認
- 1〜15項は主に武器・核・化学・生物兵器、軍用装備品等
- 金属材料・プレート・切り板は通常「非該当」だが、用途・仕様によっては確認要
- 高強度・特殊合金で軍事用途が疑われる場合は要注意
- 【重要】用途・商流・顧客に軍事・防衛・兵器関連の用途（軍用／軍事／防衛省／自衛隊／戦闘機・軍用機／戦車・装甲車／潜水艦・軍艦／ミサイル・弾道／砲・銃・弾薬等）が明示・示唆されている場合は、別表第1（特に9項航空機用エンジン・10項航空機宇宙機器・11項軍用艦船・12項戦車装甲車両・13〜14項軍用電子/誘導装置）への該当可能性が高い。この場合STEP1の結果を「非該当」にしてはならない。最低でも「要確認」、用途が明確な軍事用途であれば「該当」とすること。

STEP2: 16項（キャッチオール規制）への該当確認
- 輸出令別表第1の1〜15項非該当でも、大量破壊兵器等の開発・製造に使われるおそれがある場合は規制対象
- 仕向け国・用途・顧客の総合判断が必要

STEP3: 仕向地が輸出令別表第3（ホワイト国）か確認
- ホワイト国へのSTEP1・2非該当品 → 輸出許可不要（OK）
- 非ホワイト国 → STEP4へ
- 【重要】ホワイト国向けであっても、別表第1（1〜15項）に該当する貨物（リスト規制品）は輸出許可が必要。ホワイト国であることはリスト規制（STEP1）の免除理由にはならない。免除されるのはキャッチオール規制（STEP2・4）のみ。

STEP4: キャッチオール規制の用途確認（非ホワイト国向けの場合のみ）
①大量破壊兵器等の開発・製造・使用・貯蔵の用途かどうか
②おそれ省令別表に掲げる行為（輸出令※5）に当たるかどうか
③仕向地が輸出令別表第3の2（北朝鮮・イラン等）かつ通常兵器への用途かどうか
- ホワイト国向けでSTEP1が「非該当」の場合のみ「スキップ」可。STEP1が「該当」または「要確認」の場合は、ホワイト国でもスキップせず用途を精査すること。

STEP5: STEP4に該当する場合、用途・仕向地・エンドユーザーで個別判断

## 総合判断（overall）の整合性ルール（厳守）
- STEP1が「該当」、またはSTEP1〜4のいずれかが「該当」 → overallは「要許可」
- STEP1〜4のいずれかが「要確認」 → overallは最低でも「要確認」（「OK」にしてはならない）
- 明確な軍事用途が示されている場合 → 安全側に倒し「要許可」とすること
- 全ステップが「非該当／ホワイト国／問題なし」のときのみ overallを「OK」にできる

## 出力形式
必ず以下のJSON形式のみで回答してください。説明文は不要。

{
  "overall": "OK" または "要確認" または "要許可",
  "steps": [
    {"step": 1, "title": "別表第1（1〜15項）該当確認", "result": "非該当" または "要確認" または "該当", "reason": "判断根拠を簡潔に"},
    {"step": 2, "title": "16項（キャッチオール）確認", "result": "非該当" または "要確認" または "該当", "reason": "判断根拠を簡潔に"},
    {"step": 3, "title": "仕向地（別表第3）確認", "result": "ホワイト国" または "非ホワイト国" または "指定国", "reason": "判断根拠を簡潔に"},
    {"step": 4, "title": "用途確認（キャッチオール）", "result": "非該当" または "要確認" または "該当" または "スキップ", "reason": "判断根拠または「ホワイト国のためスキップ」"},
    {"step": 5, "title": "総合判断", "result": "問題なし" または "要精査" または "輸出不可", "reason": "最終的な根拠"}
  ],
  "recommendation": "担当者への具体的な推奨アクション（1〜3文）",
  "hs_hint": "推定HSコード（参考）",
  "caution": "注意事項や免責文言"
}"""


def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_history(entry):
    history = load_history()
    history.insert(0, entry)
    history = history[:100]  # 最新100件のみ保持
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/judge", methods=["POST"])
def judge():
    data = request.json
    customer = data.get("customer", "")
    material = data.get("material", "")
    dimensions = data.get("dimensions", "")
    purpose = data.get("purpose", "")
    trade_flow = data.get("trade_flow", "")
    destination = data.get("destination", "")

    # 事前チェック（ルールベース補助）
    is_white = is_white_country(destination)
    is_designated = is_designated_country(destination)
    scan_text = purpose + " " + trade_flow + " " + customer + " " + material
    danger_kws = has_danger_keywords(scan_text)
    military_kws = has_military_end_use(scan_text)
    hs_hint = get_hs_hint(material)

    user_message = f"""以下の輸出貨物について、5ステップの該非判定を行ってください。

【顧客名】{customer}
【材質】{material}
【寸法】{dimensions}
【用途】{purpose}
【商流（中間業者等）】{trade_flow}
【仕向け国】{destination}

【補助情報（システム事前チェック）】
- 別表第3ホワイト国: {"YES" if is_white else "NO"}
- 別表第3の2指定国: {"YES" if is_designated else "NO"}
- 大量破壊兵器関連の危険キーワード検出: {", ".join(danger_kws) if danger_kws else "なし"}
- ⚠️通常兵器・軍事用途キーワード検出: {", ".join(military_kws) if military_kws else "なし"}
{"  → 【警告】軍事用途キーワードを検出。ホワイト国向けでも別表第1（リスト規制）該当性を厳格に評価し、STEP1を非該当にせず、overallは要許可（または最低でも要確認）とすること。" if military_kws else ""}
- 推定HSコード: {hs_hint}
- 輸出令別表第1（1〜15項）主要カテゴリ:
{chr(10).join(CONTROLLED_CATEGORIES_1_TO_15)}

JSON形式のみで回答してください。"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}],
        )
        raw = response.content[0].text.strip()
        # JSON部分のみ抽出
        if "```json" in raw:
            raw = raw.split("```json")[1].split("```")[0].strip()
        elif "```" in raw:
            raw = raw.split("```")[1].split("```")[0].strip()

        result = json.loads(raw)
        result["hs_hint"] = result.get("hs_hint") or hs_hint

        # 安全ネット（ガードレール）：軍事用途キーワードを検出しているのに
        # AIが「OK」と返した場合は、安全側に倒して最低「要確認」へ強制的に引き上げる。
        if military_kws and result.get("overall") == "OK":
            result["overall"] = "要確認"
            warn = f"※システム自動補正：軍事用途キーワード（{', '.join(military_kws)}）を検出したため、判定を『要確認』に引き上げました。別表第1（リスト規制）該当性を担当者が必ず確認してください。"
            result["recommendation"] = warn + " " + result.get("recommendation", "")

        entry = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S"),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "customer": customer,
            "material": material,
            "destination": destination,
            "purpose": purpose,
            "overall": result.get("overall", "要確認"),
            "recommendation": result.get("recommendation", ""),
        }
        save_history(entry)

        return jsonify({"status": "ok", "result": result})

    except json.JSONDecodeError:
        return jsonify({"status": "ok", "result": {
            "overall": "要確認",
            "steps": [
                {"step": i, "title": t, "result": "確認中", "reason": "AI応答の解析に失敗しました。再度お試しください。"}
                for i, t in enumerate([
                    "別表第1（1〜15項）該当確認",
                    "16項（キャッチオール）確認",
                    "仕向地（別表第3）確認",
                    "用途確認（キャッチオール）",
                    "総合判断",
                ], 1)
            ],
            "recommendation": "AI応答の解析に失敗しました。入力内容を確認の上、再度お試しください。",
            "hs_hint": hs_hint,
            "caution": "本ツールはAI補助判定です。最終判断は担当者が行ってください。",
        }})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/history", methods=["GET"])
def get_history():
    return jsonify(load_history())


@app.route("/api/history/<entry_id>", methods=["DELETE"])
def delete_history(entry_id):
    history = load_history()
    history = [h for h in history if h.get("id") != entry_id]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5002))
    app.run(debug=False, host="0.0.0.0", port=port)
