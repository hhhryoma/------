# -*- coding: utf-8 -*-
import sys, os, html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sample_data import *

OUT = "sample.html"

CSS = """
@page { size: A4 landscape; margin: 11mm 11mm 12mm; }
* { box-sizing: border-box; }
body { font-family: "IPAPGothic","IPAGothic","Noto Sans CJK JP",sans-serif;
       font-size: 8.6pt; color: #1a1a1a; margin: 0; line-height: 1.5; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }
.ttl { font-size: 15pt; font-weight: bold; color: #1F4E79; margin: 0 0 2mm; }
.sub { font-size: 8pt; color: #6b6b6b; margin: 0 0 3mm; }
.bar { display: flex; align-items: baseline; justify-content: space-between;
       border-bottom: 2.2pt solid #1F4E79; padding-bottom: 1.6mm; margin-bottom: 3.2mm; }
.bar h2 { font-size: 12.5pt; color: #1F4E79; margin: 0; }
.bar .meta { font-size: 7.6pt; color: #6b6b6b; }
.badge { display: inline-block; background: #C00000; color: #fff; font-size: 7.2pt;
         font-weight: bold; padding: 0.6mm 2mm; border-radius: 1mm; letter-spacing: .04em; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
th { background: #1F4E79; color: #fff; font-weight: bold; font-size: 7.8pt;
     padding: 1.4mm 1.2mm; border: 0.4pt solid #9dbbd6; text-align: center; vertical-align: middle; }
td { border: 0.4pt solid #bfbfbf; padding: 1.3mm 1.4mm; vertical-align: top; font-size: 8pt;
     word-wrap: break-word; overflow-wrap: anywhere; }
td.c { text-align: center; vertical-align: middle; }
td.r { text-align: right; }
tr:nth-child(even) td { background: #f7fafd; }
.lbl { background: #DDEBF7 !important; font-weight: bold; color: #12385c; text-align: center; vertical-align: middle; }
.stats { display: flex; gap: 3mm; margin-bottom: 3mm; }
.stat { flex: 1; border: 0.5pt solid #9dbbd6; border-top: 2pt solid #1F4E79; padding: 1.8mm 2mm; background: #f7fafd; }
.stat .k { font-size: 7.4pt; color: #4a6d8c; }
.stat .v { font-size: 13pt; font-weight: bold; color: #1F4E79; line-height: 1.25; }
.sec { font-size: 9.6pt; font-weight: bold; color: #1F4E79; background: #DDEBF7;
       padding: 1.3mm 2mm; margin: 4mm 0 2mm; border-left: 3pt solid #1F4E79; }
.sec:first-of-type { margin-top: 0; }
.note { font-size: 7.4pt; color: #6b6b6b; margin-top: 1.8mm; }
.cover-wrap { display: flex; flex-direction: column; height: 172mm; }
.cover-mid { flex: 1; display: flex; flex-direction: column; justify-content: center; }
.cover-ttl { font-size: 24pt; font-weight: bold; color: #1F4E79; line-height: 1.35; margin-bottom: 3mm; }
.cover-sub { font-size: 12pt; color: #333; margin-bottom: 8mm; }
.cover-tbl { width: 62%; }
.cover-tbl td { font-size: 9.4pt; padding: 1.9mm 2.4mm; }
.two { display: flex; gap: 5mm; }
.two > div { flex: 1; }
.narr h3 { font-size: 9.2pt; color: #1F4E79; margin: 3mm 0 1.2mm;
           border-bottom: 0.8pt solid #9dbbd6; padding-bottom: 0.8mm; }
.narr p { margin: 0; font-size: 8.4pt; line-height: 1.62; }
.warn { border: 0.8pt solid #C00000; background: #fdf2f2; color: #8a1010;
        padding: 2mm 2.6mm; font-size: 8pt; margin-top: 3mm; }
ul.pl { margin: 0; padding-left: 5mm; font-size: 8.4pt; }
ul.pl li { margin-bottom: 1mm; }
.k-hi { color: #C00000; font-weight: bold; }
.tight td { padding: 0.95mm 1.1mm; font-size: 7.5pt; line-height: 1.42; }
.tight th { padding: 1.05mm 1mm; font-size: 7.3pt; }
.tight .sec { margin: 2.6mm 0 1.6mm; font-size: 9pt; padding: 1mm 2mm; }
.tight .stats { margin-bottom: 2.4mm; }
.tight .stat { padding: 1.4mm 2mm; }
.tight .stat .v { font-size: 12pt; }
"""

def esc(s):
    return html.escape(str(s))

def table(headers, rows, widths=None, center_cols=(), raw=False):
    cg = ""
    if widths:
        cg = "<colgroup>" + "".join('<col style="width:%s%%">' % w for w in widths) + "</colgroup>"
    h = "".join("<th>%s</th>" % esc(x) for x in headers)
    body = []
    for row in rows:
        tds = []
        for i, v in enumerate(row):
            cls = ' class="c"' if i in center_cols else ""
            tds.append("<td%s>%s</td>" % (cls, v if raw else esc(v)))
        body.append("<tr>%s</tr>" % "".join(tds))
    return "<table>%s<thead><tr>%s</tr></thead><tbody>%s</tbody></table>" % (cg, h, "".join(body))

def stats(items):
    return '<div class="stats">%s</div>' % "".join(
        '<div class="stat"><div class="k">%s</div><div class="v">%s</div></div>' % (esc(k), esc(v))
        for k, v in items)

def head(no, name, extra=""):
    return ('<div class="bar"><h2>%s　%s</h2>'
            '<div class="meta"><span class="badge">記入例（サンプル）</span>　%s／%s　%s</div></div>'
            % (esc(no), esc(name), esc(META["文書番号"]), esc(META["版数"]), esc(extra)))

P = []

# ---------- 表紙 ----------
cover_rows = [("案件名", META["案件名"]), ("お客様名", META["お客様名"]),
              ("契約・注文番号", META["契約・注文番号"]), ("報告日", META["報告日"]),
              ("文書番号／版数", "%s／%s" % (META["文書番号"], META["版数"])),
              ("作成部署", META["作成部署"]), ("作成者", META["作成者"])]
cover_tbl = "".join('<tr><td class="lbl" style="width:34%%">%s</td><td>%s</td></tr>' % (esc(k), esc(v))
                    for k, v in cover_rows)
P.append("""
<div class="page">
  <div class="cover-wrap">
    <div><span class="badge">記入例（サンプル）</span>
      <span class="sub" style="display:inline; margin-left:3mm;">
      本書は「スケジュール遅延リカバリ計画テンプレート」の記入例です。架空の案件・数値を用いています。</span></div>
    <div class="cover-mid">
      <div class="cover-ttl">スケジュール遅延に関するご報告<br>およびリカバリ計画</div>
      <div class="cover-sub">%s</div>
      <table class="cover-tbl">%s</table>
    </div>
    <div>
      <div class="sec" style="margin-bottom:2mm;">本書の構成（Excelテンプレートの各シートに対応）</div>
      <div class="two">
        <div><ul class="pl">
          <li><b>1. サマリ</b>　現状の数値と、ご報告・ご依頼事項</li>
          <li><b>2. 遅延状況一覧</b>　遅延タスクと遅延日数</li>
          <li><b>3. 原因分析</b>　事象→直接原因→真因→再発防止策</li>
          <li><b>4. リカバリ施策</b>　施策一覧と案A/B/Cの比較</li>
        </ul></div>
        <div><ul class="pl">
          <li><b>5. 新スケジュール</b>　変更後マイルストーンと前提条件</li>
          <li><b>6. 影響評価とリスク</b>　納期・コスト・品質への影響</li>
          <li><b>7. アクション管理</b>　誰が・何を・いつまでに</li>
          <li><b>8. 週次モニタリング</b>　回復状況の定点報告</li>
        </ul></div>
      </div>
    </div>
  </div>
</div>""" % (esc(META["案件名"]), cover_tbl))

# ---------- 1. サマリ ----------
sm_rows = "".join('<tr><td class="lbl" style="width:42%%">%s</td><td class="c" style="font-weight:bold;">%s</td></tr>'
                  % (esc(k), esc(v)) for k, v, _ in SUMMARY)
narr = "".join('<h3>%s</h3><p>%s</p>' % (esc(t), b) for t, b in NARRATIVE)
P.append("""
<div class="page">
  %s
  <div class="two">
    <div style="flex:0 0 40%%;">
      <div class="sec">1. 現状サマリ</div>
      <table>%s</table>
      <div class="warn"><b>ご判断のお願い：</b>与信チェック機能の第2次リリース分割を
      <b>7月17日まで</b>にご承認いただけますと、残る5日を吸収し当初どおり
      <b>2026年9月30日</b>の納品が可能です。</div>
      <div class="note">※ Excelテンプレートでは、これらの数値は各シートから自動参照されます（手入力は「現時点の完了見込」のみ）。</div>
    </div>
    <div class="narr">%s</div>
  </div>
</div>""" % (head("1.", "サマリ"), sm_rows, narr))

# ---------- 2. 遅延状況 ----------
P.append("""
<div class="page">
  %s
  %s
  %s
  <div class="note">※ CP＝クリティカルパス。「遅延日数」「件数」「最大遅延」「平均進捗率」はExcel上で自動計算されます。</div>
</div>""" % (head("2.", "遅延状況一覧"), stats(DELAY_STAT),
             table(DELAY_HEAD, DELAY, widths=[3,5,15,9,7,7,8,8,5,4,5,10,14],
                   center_cols=(0,1,4,5,6,7,8,9,10))))

# ---------- 3. 原因分析 ----------
P.append("""
<div class="page">
  %s
  %s
  <div class="note">※「責任区分」は社内整理用の項目です。契約・請求に直結するため、社外提出時は営業／契約担当と取扱いを事前確認してください。</div>
</div>""" % (head("3.", "原因分析（なぜなぜ・再発防止）"),
             table(CAUSE_HEAD, CAUSE, widths=[3,4,13,11,11,11,13,6,5,14,5,4],
                   center_cols=(0,1,7,8,10,11))))

# ---------- 4. リカバリ施策 ----------
P.append("""
<div class="page tight">
  %s
  %s
  %s
  <div class="sec">リカバリ案の比較（ご判断の材料）</div>
  %s
</div>""" % (head("4.", "リカバリ施策と案の比較"), stats(MEASURE_STAT),
             table(MEASURE_HEAD, MEASURE, widths=[3,12,4,16,5,8,13,14,5,6,7,7],
                   center_cols=(0,2,4,5,8,9,10,11)),
             table(COMPARE_HEAD, COMPARE, widths=[12,22,22,22,22], center_cols=())))

# ---------- 5. 新スケジュール ----------
prem = "".join("<li>%s</li>" % esc(x) for x in PREMISE)
P.append("""
<div class="page">
  %s
  %s
  %s
  <div class="sec">本スケジュールの前提条件（崩れた場合は再計画が必要です）</div>
  <ul class="pl">%s</ul>
</div>""" % (head("5.", "新スケジュール（変更後マイルストーン）"),
             stats([("最終納品日（当初）", "2026/09/30"), ("最終納品日（変更後）", "2026/10/05"),
                    ("差異", "5日"), ("案Bご承認時", "9/30を維持")]),
             table(MS_HEAD, MS, widths=[4,20,10,10,6,7,43], center_cols=(0,2,3,4,5)),
             prem))

# ---------- 6. 影響・リスク ----------
P.append("""
<div class="page">
  %s
  <div class="sec">1. 遅延による影響評価</div>
  %s
  <div class="sec">2. リカバリ計画のリスク一覧（スコア＝発生確率×影響度、最大9）</div>
  %s
</div>""" % (head("6.", "影響評価とリスク"),
             table(IMPACT_HEAD, IMPACT, widths=[12,34,18,36], center_cols=(0,)),
             table(RISK_HEAD, RISK, widths=[3,28,6,6,5,33,8,7,6], center_cols=(0,2,3,4,6,7,8))))

# ---------- 7. アクション管理 ----------
P.append("""
<div class="page">
  %s
  %s
  %s
  <div class="note">※「期限超過」は本日基準で自動判定されます（本サンプルは2026年7月17日時点。No.2が期限超過）。</div>
</div>""" % (head("7.", "アクション管理"), stats(ACT_STAT),
             table(ACT_HEAD, ACT, widths=[3,9,6,36,13,9,7,9,8], center_cols=(0,1,2,5,6,7,8))))

# ---------- 8. 週次モニタリング ----------
P.append("""
<div class="page">
  %s
  %s
  <div class="note">※「達成率」「累計」「差異」は自動計算されます。「遅延回復日数」は当初計画に対する遅れの増減をプラス（回復）／マイナス（悪化）で記入します。</div>
  <div class="sec">本サンプルについて</div>
  <ul class="pl">
    <li>本書は「スケジュール遅延リカバリ計画テンプレート（Excel）」の記入例であり、<b>架空の案件・数値</b>を用いています。実在の組織・案件とは関係ありません。</li>
    <li>実際の運用では、Excelテンプレートの黄色いセルに記入すると、遅延日数・合計・件数・達成率などが自動計算されます。</li>
    <li>お客様へ提出される際は、記入例の行をすべて削除し、「責任区分」など社内整理用の項目の取扱いをご確認ください。</li>
  </ul>
</div>""" % (head("8.", "週次モニタリング（回復状況の定点報告）"),
             table(WEEK_HEAD, WEEK, widths=[4,9,7,7,5,7,7,7,8,39], center_cols=(0,1,2,3,4,5,6,7,8))))

doc = "<!doctype html><html lang='ja'><head><meta charset='utf-8'><title>スケジュール遅延リカバリ計画（記入例）</title><style>%s</style></head><body>%s</body></html>" % (CSS, "".join(P))
open(OUT, "w", encoding="utf-8").write(doc)
print("html written:", OUT, len(doc))
