import re
import streamlit as st
import anthropic
from google import genai
from google.genai import types
import os
import json
from dotenv import load_dotenv
from datetime import datetime
from pathlib import Path
import openpyxl

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"), override=True)

# ------------------------------------------------------------------ #
# システムプロンプト（Claudeプロジェクトの指示 + 参考台本）
# ------------------------------------------------------------------ #
SYSTEM_PROMPT = """以下の構成でYouTube Shorts用の台本を作成してください:

【テーマ】
日本の高度経済成長期に活躍した日本企業や会社員の美談を紹介する動画

【感情を動かすキーポイント】
- 「現在にはないやる気と根性で経済を好調させるすごさ」
- 「低迷が続く日本にもこんな輝かしい時代があったんだという感動」
- 「日本はほんとはこんなもんじゃないという気概」

【構成】

1. **タイトル(1行目)**:
   日本の黎明期を支えたヒーロー的企業や会社員を
   衝撃的なフレーズで表現し、視聴意欲を喚起

2. **問題提起(2行目)**:
   日本人に持っている固定概念などの現状を
   視聴者が共感できる形で提示

3. **固定概念の裏切り(3行目)**:
   日本人が抱いていた固定概念を裏切る、またはその通り過ぎて逆に面白いという事実、発言で期待感を高める

4. **具体的行動1(4行目)**:
   相手を知ることによってやっと始めの土台に立つことが出来る具体的で心温まるエピソードを描写

以下の構成でYouTube Shorts用の台本を作成してください:

1. **タイトル(1行目)**: 結末を示唆する衝撃的なフレーズで視聴意欲を喚起
2. **問題提起(2行目)**: 視聴者が共感できる社会問題や困りごとを提示し、敵役/課題を明確化
3. **ヒーロー登場(3行目)**: 問題を解決する主役を登場させ、期待感を高める
4. **具体的戦術1(4行目)**: 主役の賢い対策や行動を具体例で示す
5. **対立の激化(5行目)**: 敵役の反撃や状況悪化を示し、緊張感を維持
6. **具体的戦術2-3(6-7行目)**: さらに進化した対策を複数示し、本気度や徹底ぶりを伝える
7. **展開の転換点(8行目)**: 「これで終わりではない!」など、まだ続きがあることを示唆
8. **新展開(9行目)**: 新キャラ登場や予想外の展開で物語を拡大
9. **逆転・因果応報(10行目)**: 敵役の失敗や主役の成功を痛快に描写
10. **クライマックス(11行目)**: タイトルで示唆した「末路」や結末を具体的に提示しカタルシスを提供
11. **メッセージ(12行目)**: 主役の姿勢や価値観を示し、感動と共感で締める
各行は簡潔に、テンポよく展開させ、視聴者が最後まで離脱しない構成にしてください。

---

【台本分析：視聴者を引きつける要素】
| 要素 | 効果 |
|------|------|
| 冒頭の情報過多 | 反射的注意を引く |
| 「実は〜だけではない」構文 | 最後まで見る動機を作る |
| 対比構造 | 感情的な落差を最大化する |
| 具体的な数字・年・固有名詞 | 信憑性と臨場感を高める |
| 「なんと」「さらに」「信じられないことに」 | 感情の盛り上がりを段階的に演出する |
| どんでん返し | クライマックスで感情的ピークを作る |
| 海外・第三者の反応引用 | 信頼性を補強する |
| 余韻ある締め | コメント・拡散を自然に誘導する |

---

【参考台本 - 文章の構成・話し方・言い回しのみ参考にし、内容は参考にしないこと】

■台本1
世界で唯一日本人だけが読めない3文字に世界が大爆笑w
日本語は海外の人にとって
習得が難しい言語だ
しかし、そんな難しい言語を使いこなす
日本人だけが読めない文字があるらしい
海外のネット上で
「日本人はマジで読めないの？噓でしょ？」
と話題になっている
ではみんなには
この文字が読めるだろうか？
正解はPEN
驚くことに
この文字は日本人以外には普通に読めるようだ！
ではこれはどうだろう？
正解はHELLO
申し訳ないが
もう一問付き合ってもらおう！
正解はGOOD　MORNING！
混乱させて悪かったな！
おそらくこれを読めた人は日本人じゃないと思う！
実はこの文字の正体は
「エレクトロハーモニクス」という
日本語をモチーフにしたアルファベットのフォント
日本人がみると
どうしても日本語に見えてしまい
読みづらいのだ
「普通に読めるよ！
日本人には本当に読めないの？」と
海外の人には面白いようだが
何回見ても読めんｗ

■台本2
ヨドバシの罠にまんまとはまった中国人転売ヤ—の末路
レア商品を根こそぎ買い占め
高値で転売する中国人転売ヤ—の存在が
社会的な問題に
そんな迷惑転売ヤ—の企みを阻止したのが
日本のヨドバシだった
レア商品を売る際商品名を答えさせたり
スマホを使い画像を検索させたりするなどして
転売ヤ—を炙りだしたのだ
転売ヤーもすぐさま対応するが
ヨドバシの対策の方が勝っていた
日本語が分からないと
誤魔化す中国人に対処するため
中国語が話せる店員を配置
さらには商品の箱に店の捺印をすることで
転売しにくくするなど
徹底的な対策を施し
悪質な転売ヤ—を撃退したのだ
これで終わりではない！転売ヤーの悪夢は続く
買占めに起こった本当のファン達が
「一泡吹かせてやろう」と
結束してデマを流すと
転売ヤーはまんまと引っかかり
全く売れない商品を大量購入w
大損害を受ける羽目に
商売が成り立たなくなり
廃業に追い込まれる者まで現れた！
転売対策には費用も人員もかかるのに
「真のファンに買ってもらえてこそ！」
ヨドバシの誠意が十分に伝わる

■台本3
「絶対CG!」暇つぶしで日本人が作った雪だるまに世界が大爆笑w
海外では
「スノーマン」と呼ばれる
雪だるま
マフラーを巻いたり帽子を被せたり
にんじんを刺したり
人に見立てて様々な工夫が
凝らされている
そんななか...
海外のネット上で日本人が作った
あまりにも独創的な
雪だるまの写真が話題に！
まず...ピカチュウの雪だるま！
可愛らしいだるまの形に
尻尾や耳が作り込まれ
表情もにっこり
次に少し意外な...
仏像の雪だるま！
溶けて消えてしまう雪で
象ったことで「諸行無常」を
表現したのかもしれない
そして欧米でも大人気の
ゴジラ！
雪とは思えない
皮膚のゴツゴツとした質感や
背中の突起物
鋭い牙など見事に表現！
さらに翼を失った天使
割れた氷がガラスのように見え
ストーリー性を感じられる作品...
まさにアートだ
その他ジブリ作品の雪だるまとして
トトロやカオナシ
王蟲など
様々な写真が登場！
日本独特の雪だるまに
外国人の反応は...
「なんてクリエイティブな国だ！」
「俺の国のスノーマンとは
全然違う...」
発想力もだが手先が器用すぎるw

■台本4
「独立させろ」「どうぞ×２」韓国ダイソーが日本から強引に独立した結果...
韓国のダイソーが完全に
韓国企業に...
韓国国内のダイソーは
韓国企業のアソンダイソーが
日本の大創産業の協力を得て
設立店舗数を伸ばした
ところが2023年までに
アソンダイソーが
大創産業が保有していた株を全取得
韓国で日本製品の不買運動が
起こるたびに
槍玉に上げられた韓国ダイソーだったが
悲願の日本離れを達成！
日本がいなくなってハッピーｗ
と思いきや一筋縄ではいかず...
良いものを安く提供することは
並大抵の企業努力では叶わない
独立した韓国ダイソーには
100円から500円の
幅広い価格帯の商品が並び
コンセプトがブレブレ
100円のイメージが強い
日本のダイソーとは異なる
また世界的知名度の差も圧倒的
韓国のダイソーはあくまで韓国国内の
ダイソーであるのに対し
日本のダイソーは世界25ヵ国に
1000店舗以上を展開
ちなみに
大創産業側は
約20年前に4億円ほどで
入手したアソンダイソーの株が
何百倍にもなって戻ってきたという
なので痛くも痒くもないご様子

■台本5
「日本へ復讐」中国で"日本旅行しない運動"がトレンド入りで日本が異常事態...
中国人にとって
人気の観光先である日本
円安などを背景に
訪日中国人は増加してきたが...
そんな中中国SNSで突如
日本旅行への
ボイコット運動が活発に！
そのきっかけは
元迷惑系Youtuberの
"へずまりゅう"氏
彼は2024年に中国人観光客が
奈良公園の
鹿を蹴る動画を発見し
"鹿を守る活動"を開始
マナーを守らない観光客に対し
注意を促す姿が話題となった
これに対し中国SNSで
「日本のインフルエンサーが
反中国的発言を繰り返している」
「日本特に奈良には
行かない方がいい」という
日本旅行しない運動が拡散したのだ
規模は次第に大きくなり...
日本経済を揺るがす問題に
日本人は大慌て...かと思いきや
奈良をはじめ
各地でマナー違反に
悩まされてきた日本人からは
「願ったり叶ったり！」と
むしろ歓迎の声が上がっている
しかも実際に日本旅行を
ボイコットする中国人はごく僅か
実際の統計では過去最高ペースで
中国人観光客が増加しているという
文句言いたかっただけなのか？
"""


# ------------------------------------------------------------------ #
# API 関数
# ------------------------------------------------------------------ #
def _get_secret(key: str) -> str:
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, "")


def get_gemini_client():
    api_key = _get_secret("GEMINI_API_KEY")
    if not api_key:
        st.error("GEMINI_API_KEY が設定されていません。")
        st.stop()
    return genai.Client(api_key=api_key)


def get_claude_client():
    api_key = _get_secret("CLAUDE_API_KEY")
    if not api_key:
        st.error("CLAUDE_API_KEY が設定されていません。")
        st.stop()
    return anthropic.Anthropic(api_key=api_key)


def step1_generate_cases(theme: str) -> str:
    client = get_gemini_client()
    used = load_used_cases(theme)
    exclude_text = ""
    if used:
        exclude_text = "\n\n【除外する事例（すでに台本化済み）】\n" + "\n".join(f"- {c}" for c in used) + "\n上記の事例は絶対に含めないでください。"

    prompt = f"""あなたはYouTubeショート動画のプロデューサーです。
テーマ「{theme}」について、視聴者が「知らなかった！」「すごい！」と感じる
感動的・驚き・誇りを感じられる具体的な事例を5つ提案してください。{exclude_text}

各事例について以下の形式で記載してください：

【事例1】
- 会社名/人物名：
- 時代（年代）：
- エピソード概要（2〜3文）：
- 動画タイトル案（衝撃的なフレーズ）：

【事例2】
...（同様に5つ）"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )
    return response.text


def step2_research_case(case_description: str) -> str:
    client = get_gemini_client()
    prompt = f"""以下の事例について、YouTubeショート動画の台本作成に使える詳細情報をウェブ検索で調査してください。

【事例】
{case_description}

以下を収集・整理してください：
- 正確な年代・数字・固有名詞
- 具体的なエピソード・逸話（感動的なシーン）
- 関係者の発言・名言（あれば）
- 驚きや感動を与える事実
- 当時の社会的背景・時代感

※事実確認済みの情報のみ記載。不確かな情報には「要確認」と付記してください。"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())]
        ),
    )
    return response.text


def step1b_generate_hypothesis(case_description: str) -> str:
    client = get_gemini_client()
    prompt = f"""あなたはYouTubeショート動画のプロデューサーです。
以下の事例に対して、動画の仮説を一文で作成してください。

【事例】
{case_description}

【テンプレート】
「この動画は、【誰に】／【どんな感情】を起こすために、【冒頭の一言・見せ方】で掴み、【どんな結末や転換】で心を動かし、最後に【どんな問い】で言葉を引き出す。」

条件：
- "なぜその視聴者の心が動くのか"を必ず明示すること
- 「話題だから伸びそう」「有名人だから伸びそう」といった表面的な理由はNG
- 具体的なターゲット視聴者・感情・冒頭の掴み・転換・締めの問いを盛り込む
- 一文（または2〜3文の箇条書き形式）で言い切ること

仮説のみを出力してください。前置きや説明は不要です。"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )
    return response.text.strip()


def step2b_generate_framework(case: str, research: str, hypothesis: str) -> dict:
    client = get_gemini_client()
    prompt = f"""あなたはYouTubeショート動画のプロデューサーです。
以下の情報をもとに、動画制作の4点フレームワークを生成してください。

【事例】
{case}

【調査結果】
{research}

【動画の仮説】
{hypothesis}

以下の4点それぞれについて、具体的な案を1〜2文で記載してください。

**ファーストビュー**：主語と状況が一目で分かる・動きがある・強い言葉を1つ置く
例：「電車内で"まさか"の一言」「外国人が引くほど驚いた日本の○○」

**約束**：「この先で○○が分かります／ひっくり返ります」と宣言する（わざと約束をしないパターンも可）

**見せ場**：事実や描写をサクッと重ね、意外性や逆転で気持ちを動かす

**余韻**：二択・体験共有・価値判断のどれかでコメントを誘う
例：「注意すべき／静観すべき、あなたはどっち？」

以下のJSON形式のみを出力してください（説明不要）：
{{
  "first_view": "...",
  "promise": "...",
  "highlight": "...",
  "lingering": "..."
}}"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )
    text = response.text.strip()
    text = re.sub(r'^```(?:json)?\n?', '', text)
    text = re.sub(r'\n?```$', '', text)
    return json.loads(text)


def step3_generate_script(case_outline: str, research: str, hypothesis: str = "", framework: dict | None = None) -> str:
    client = get_claude_client()

    hypothesis_block = f"\n\n【動画の仮説（一文）】\n{hypothesis}" if hypothesis.strip() else ""

    framework_block = ""
    if framework:
        framework_block = f"""

【4点フレームワーク（制作指針）】
- ファーストビュー：{framework.get('first_view', '')}
- 約束：{framework.get('promise', '')}
- 見せ場：{framework.get('highlight', '')}
- 余韻：{framework.get('lingering', '')}"""

    user_message = f"""という内容のショート動画の台本を以下の記事を参考にして生成してください。字数は400字程度にしてください。{hypothesis_block}{framework_block}

【選択した事例】
{case_outline}

【詳細調査結果（Gemini + Google検索）】
{research}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    return message.content[0].text


EXCEL_PATH = Path(__file__).parent / "台本履歴.xlsx"
EXCEL_HEADERS = ["保存日時", "テーマ", "事例", "台本"]
USED_CASES_PATH = Path(__file__).parent / "used_cases.json"


def load_used_cases(theme: str) -> list[str]:
    if not USED_CASES_PATH.exists():
        return []
    data = json.loads(USED_CASES_PATH.read_text(encoding="utf-8"))
    return data.get(theme, [])


def add_used_case(theme: str, case: str) -> None:
    data = {}
    if USED_CASES_PATH.exists():
        data = json.loads(USED_CASES_PATH.read_text(encoding="utf-8"))
    data.setdefault(theme, [])
    if case not in data[theme]:
        data[theme].append(case)
    USED_CASES_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create_excel_bytes(theme: str, case: str, script: str) -> bytes:
    import io
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "台本履歴"
    ws.append(EXCEL_HEADERS)
    for col in range(1, 5):
        ws.cell(1, col).font = openpyxl.styles.Font(bold=True)
    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), theme, case, script])
    col_widths = [20, 30, 50, 60]
    for col, width in zip(range(1, 5), col_widths):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save_to_excel(theme: str, case: str, script: str) -> None:
    if EXCEL_PATH.exists():
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "台本履歴"
        ws.append(EXCEL_HEADERS)
        for col in range(1, 5):
            ws.cell(1, col).font = openpyxl.styles.Font(bold=True)

    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        theme,
        case,
        script,
    ])

    # 列幅を自動調整
    col_widths = [20, 30, 50, 60]
    for col, width in zip(range(1, 5), col_widths):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    # 台本列は折り返し表示
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")

    wb.save(EXCEL_PATH)


def save_to_notion(theme: str, case: str, script: str) -> None:
    from notion_client import Client
    token = _get_secret("NOTION_TOKEN")
    database_id = _get_secret("NOTION_DATABASE_ID")
    if not token or not database_id:
        raise ValueError("NOTION_TOKEN または NOTION_DATABASE_ID が設定されていません。")
    notion = Client(auth=token)

    def rich_text(content: str) -> list:
        return [{"text": {"content": content[:2000]}}]

    notion.pages.create(
        parent={"database_id": database_id},
        properties={
            "台本タイトル": {"title": rich_text(theme)},
            "保存日時": {"rich_text": rich_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))},
            "テーマ": {"rich_text": rich_text(theme)},
            "事例": {"rich_text": rich_text(case[:2000])},
            "台本": {"rich_text": rich_text(script[:2000])},
        },
    )


def refine_script(current_script: str, instructions: str) -> str:
    client = get_claude_client()

    user_message = f"""以下の台本を指示に従って推敲してください。字数は400字程度を維持してください。

【現在の台本】
{current_script}

【推敲の指示】
{instructions}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    return message.content[0].text


# ------------------------------------------------------------------ #
# Streamlit UI
# ------------------------------------------------------------------ #
st.set_page_config(page_title="ショート動画台本生成", layout="wide")
st.title("ショート動画台本生成ツール")
st.caption("Gemini（事例立案・調査）× Claude（台本生成）")

# APIキー確認
if not _get_secret("GEMINI_API_KEY"):
    st.warning("GEMINI_API_KEY が設定されていません。Streamlit Cloud の Secrets に追加してください。")
if not _get_secret("CLAUDE_API_KEY"):
    st.warning("CLAUDE_API_KEY が設定されていません。Streamlit Cloud の Secrets に追加してください。")

st.divider()

# テーマ入力
theme = st.text_input(
    "テーマを入力",
    value="高度経済成長期に活躍した企業・会社員の美談",
    help="Geminiに渡すテーマです。自由に変更できます。",
)

st.divider()

def _parse_cases(cases_text: str) -> list[str]:
    parts = re.split(r'【事例\d+】', cases_text)
    return [p.strip() for p in parts if p.strip()]

def _case_label(case_text: str, index: int) -> str:
    for line in case_text.split('\n'):
        if '動画タイトル案' in line:
            title = line.split('：', 1)[-1].strip().lstrip('「').rstrip('」')
            if title:
                return f"事例{index + 1}：{title}"
    return f"事例{index + 1}"


# ---- Phase 1: 事例立案 ----
if st.button("事例を立案する", type="primary", key="btn_step1", use_container_width=True):
    for k in ("cases", "case_options", "selected_case", "hypothesis_draft", "research", "script"):
        st.session_state.pop(k, None)

    try:
        with st.status("Step 1　事例を立案中（Gemini）...", expanded=True) as status:
            cases = step1_generate_cases(theme)
            st.session_state["cases"] = cases
            st.session_state["case_options"] = _parse_cases(cases)
            status.update(label="Step 1　事例の立案完了", state="complete")
    except Exception as e:
        st.error(f"エラー: {e}")

# ---- Phase 1.5: 事例選択 + 仮説入力 ----
if st.session_state.get("case_options"):
    st.divider()
    st.subheader("事例を選択")

    options = st.session_state["case_options"]
    labels = [_case_label(opt, i) for i, opt in enumerate(options)]
    selected_idx = st.radio("使用する事例を選んでください", range(len(options)), format_func=lambda x: labels[x], key="case_radio")
    selected_case = options[selected_idx]
    st.session_state["selected_case"] = selected_case

    with st.expander("選択した事例の詳細", expanded=False):
        st.text(selected_case)

    st.divider()
    st.subheader("動画の仮説（投稿前チェック）")
    st.caption(
        "テンプレート：「この動画は、【誰に】／【どんな感情】を起こすために、【冒頭の一言・見せ方】で掴み、"
        "【どんな結末や転換】で心を動かし、最後に【どんな問い】で言葉を引き出す。」"
    )

    col_hyp_btn, _ = st.columns([1, 3])
    with col_hyp_btn:
        if st.button("仮説を自動生成", key="btn_gen_hypothesis"):
            with st.spinner("仮説を生成中..."):
                try:
                    st.session_state["hypothesis_input"] = step1b_generate_hypothesis(selected_case)
                    st.rerun()
                except Exception as e:
                    st.error(f"エラー: {e}")

    hypothesis = st.text_area(
        "仮説を入力（または自動生成してから編集）",
        height=110,
        placeholder='例：「通勤中の流し見層に"驚き"を起こす。冒頭で断定の一言→中盤で意外な事実→最後は"あなたはどう思う？"の二択でコメントを促す。」',
        key="hypothesis_input",
    )

    # ---- Phase 2: 調査 ----
    st.divider()
    if st.button("調査する", type="primary", key="btn_step2", use_container_width=True, disabled=not hypothesis.strip()):
        st.session_state.pop("research", None)
        st.session_state.pop("framework", None)
        st.session_state.pop("script", None)

        try:
            with st.status("Step 2　事例を調査中（Gemini + Google検索）...", expanded=True) as status:
                research = step2_research_case(selected_case)
                st.session_state["research"] = research
                status.update(label="Step 2　調査完了", state="complete")
        except Exception as e:
            st.error(f"エラー: {e}")

if st.session_state.get("research"):
    with st.expander("調査結果（Gemini + Google検索）", expanded=False):
        st.text(st.session_state["research"])

    # ---- Phase 2.5: 4点フレームワーク ----
    st.divider()
    st.subheader("4点フレームワーク（制作チェック）")

    col_fw_btn, _ = st.columns([1, 3])
    with col_fw_btn:
        if st.button("フレームワークを自動生成", key="btn_gen_framework"):
            with st.spinner("フレームワークを生成中..."):
                try:
                    fw = step2b_generate_framework(
                        st.session_state.get("selected_case", ""),
                        st.session_state["research"],
                        st.session_state.get("hypothesis_input", ""),
                    )
                    st.session_state["fw_first_view"] = fw.get("first_view", "")
                    st.session_state["fw_promise"] = fw.get("promise", "")
                    st.session_state["fw_highlight"] = fw.get("highlight", "")
                    st.session_state["fw_lingering"] = fw.get("lingering", "")
                    st.rerun()
                except Exception as e:
                    st.error(f"エラー: {e}")

    fw_first_view = st.text_area(
        "ファーストビュー（最初のひと言と見せ方）",
        height=80,
        placeholder='例：「外国人が引くほど驚いた日本の○○」「電車内で"まさか"の一言」',
        key="fw_first_view",
    )
    fw_promise = st.text_area(
        "約束（何が起きる？何が得られる？）",
        height=80,
        placeholder='例：「この先でひっくり返ります」※約束しないパターンも可',
        key="fw_promise",
    )
    fw_highlight = st.text_area(
        "見せ場（期待の回収・転換）",
        height=80,
        placeholder="事実や描写をサクッと重ね、意外性・逆転で気持ちを動かす",
        key="fw_highlight",
    )
    fw_lingering = st.text_area(
        "余韻（言葉を引き出す問い）",
        height=80,
        placeholder='例：「注意すべき／静観すべき、あなたはどっち？」',
        key="fw_lingering",
    )

    # ---- Phase 3: 台本生成 ----
    st.divider()
    current_framework = {
        "first_view": fw_first_view,
        "promise": fw_promise,
        "highlight": fw_highlight,
        "lingering": fw_lingering,
    }
    if st.button("台本を生成する", type="primary", key="btn_step3", use_container_width=True):
        st.session_state.pop("script", None)
        hypothesis_val = st.session_state.get("hypothesis_input", "")
        try:
            with st.status("Step 3　台本を執筆中（Claude）...", expanded=True) as status:
                script = step3_generate_script(
                    st.session_state.get("selected_case", ""),
                    st.session_state["research"],
                    hypothesis_val,
                    current_framework,
                )
                st.session_state["script"] = script
                status.update(label="Step 3　台本生成完了", state="complete")
        except Exception as e:
            st.error(f"エラー: {e}")

# ---- 台本 ----
if st.session_state.get("script"):
    st.divider()
    st.subheader("生成された台本")
    st.text_area("台本（400字程度）", st.session_state["script"], height=420)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button(
            label="テキストでダウンロード",
            data=st.session_state["script"],
            file_name="script.txt",
            mime="text/plain",
        )
    with col2:
        case = st.session_state.get("selected_case", "")
        st.download_button(
            label="Excelでダウンロード",
            data=create_excel_bytes(theme, case, st.session_state["script"]),
            file_name="台本履歴.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col3:
        if st.button("Notionに保存", key="btn_notion"):
            try:
                case = st.session_state.get("selected_case", "")
                save_to_notion(theme, case, st.session_state["script"])
                add_used_case(theme, case)
                st.success("Notionに保存しました！")
            except Exception as e:
                st.error(f"保存エラー: {e}")

    st.divider()

    # ---- 推敲 ----
    st.subheader("推敲（Claude）")
    refine_instructions = st.text_area(
        "修正の指示を入力",
        height=100,
        placeholder="例：タイトルをもっと衝撃的にして／字数を減らして／感動的な締めにして",
        key="refine_instructions",
    )
    if st.button("推敲する", type="primary", key="btn_refine") and refine_instructions.strip():
        with st.spinner("Claudeが推敲中..."):
            try:
                refined = refine_script(st.session_state["script"], refine_instructions)
                st.session_state["script"] = refined
                st.session_state.pop("script_display", None)
                st.rerun()
            except Exception as e:
                st.error(f"エラー: {e}")
