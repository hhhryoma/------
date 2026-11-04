"""最小のエージェント・ハーネス（ステップ 1）

モデルを呼ぶ → モデルが頼んだツールを実行する → 結果を返す、を繰り返します。

使い方:
    pip install -r requirements.txt
    python harness.py                        # 起動してからタスクを入力する
    python harness.py このフォルダのファイルを数えて   # タスクを直接渡す

API キーは環境変数 ANTHROPIC_API_KEY に入れておきます（README を見てください）。
"""

import subprocess
import sys

import anthropic

MODEL = "claude-opus-5-5"
MAX_TURNS = 30  # ループの上限。モデルがツールを呼び続けて止まらないときの安全装置

SYSTEM = (
    "あなたはユーザーの PC で作業するアシスタントです。"
    "必要に応じてツールを使い、終わったら結果を短く報告してください。"
)

# ① ツールの定義: モデルに「こういう道具がある」と伝える
TOOLS = [
    {
        "name": "read_file",
        "description": "テキストファイルの中身を読む",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "読むファイルのパス"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "run_command",
        "description": "シェルのコマンドを実行して、その出力を返す",
        "input_schema": {
            "type": "object",
            "properties": {
                "command": {"type": "string", "description": "実行するコマンド"},
            },
            "required": ["command"],
        },
    },
]


# ② ツールの実行: モデルは「使いたい」と言うだけで、実際に動かすのはここ
def execute_tool(name, args, ask=input):
    """ツールを実行して (結果の文字列, エラーかどうか) を返す"""
    try:
        if name == "read_file":
            with open(args["path"], encoding="utf-8", errors="replace") as f:
                return f.read()[:20000], False  # 長すぎる結果は切る

        if name == "run_command":
            # ③ 権限の管理: コマンドは実行する前に人に確かめる
            answer = ask(f"  実行してよいですか? $ {args['command']} [y/N] ")
            if answer.strip().lower() != "y":
                return "ユーザーが実行を断りました", True
            r = subprocess.run(
                args["command"],
                shell=True,
                capture_output=True,
                text=True,
                errors="replace",
                timeout=60,
            )
            output = (r.stdout + r.stderr)[-10000:] or "(出力なし)"
            return output, r.returncode != 0

        return f"知らないツールです: {name}", True

    except Exception as e:
        # エラーも結果としてモデルに返す。モデルはそれを見て自分で直そうとする
        return f"エラー: {e}", True


# ④ エージェントのループ
def run_agent(task, client, ask=input, max_turns=MAX_TURNS):
    """タスクが終わるまでループし、会話の履歴を返す"""
    messages = [{"role": "user", "content": task}]

    for _ in range(max_turns):
        response = client.beta.messages.create(
            model=MODEL,
            max_tokens=16000,
            system=SYSTEM,
            tools=TOOLS,
            messages=messages,
            output_config={"effort": "high"},
            # 安全のフィルタで断られたとき、別のモデルで自動的にやり直す
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
        )

        # モデルの返事は手を加えず、そのまま履歴に足す
        messages.append({"role": "assistant", "content": response.content})

        for block in response.content:
            if block.type == "text":
                print(f"Claude: {block.text}")

        # ツールを使わないなら終わり（end_turn / max_tokens / refusal など）
        if response.stop_reason != "tool_use":
            print(f"[終了: {response.stop_reason}]")
            return messages

        # 頼まれたツールをすべて実行し、結果を 1 つのメッセージにまとめて返す
        results = []
        for block in response.content:
            if block.type == "tool_use":
                print(f"  [ツール] {block.name} {block.input}")
                output, is_error = execute_tool(block.name, block.input, ask)
                results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,  # どの依頼への結果かを対応づける
                    "content": output,
                    "is_error": is_error,
                })
        messages.append({"role": "user", "content": results})

    print(f"[終了: {max_turns} 回で打ち切りました]")
    return messages


def main():
    client = anthropic.Anthropic()  # 環境変数 ANTHROPIC_API_KEY を使う
    task = " ".join(sys.argv[1:]) or input("タスク> ")
    run_agent(task, client)


if __name__ == "__main__":
    main()
