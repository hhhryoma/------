"""API を呼ばずに、ハーネスのループの動きを確かめるテスト

    python -m unittest test_harness.py -v

決めておいた返事を順に返す「偽のクライアント」を使うので、API キーはいりません。
"""

import os
import sys
import tempfile
import unittest
from types import SimpleNamespace as NS

import harness


def text(t):
    return NS(type="text", text=t)


def tool_use(id, name, input):
    return NS(type="tool_use", id=id, name=name, input=input)


def reply(stop_reason, *blocks):
    return NS(stop_reason=stop_reason, content=list(blocks))


class FakeClient:
    """決めておいた返事を順に返し、送られてきたリクエストを記録する"""

    def __init__(self, *responses):
        self.responses = list(responses)
        self.requests = []
        self.beta = NS(messages=NS(create=self._create))

    def _create(self, **kwargs):
        # 履歴はあとで書き足されるので、送られた時点の中身を写しておく
        self.requests.append({**kwargs, "messages": list(kwargs["messages"])})
        if len(self.responses) == 1:
            return self.responses[0]  # 最後の返事はくり返し返す
        return self.responses.pop(0)


def never_ask(prompt):
    raise AssertionError("確認を求めないはずの場面で確認を求めました")


class HarnessTest(unittest.TestCase):
    def test_ツールの結果をモデルに返してから終わる(self):
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".txt", delete=False) as f:
            f.write("こんにちは")
        self.addCleanup(os.remove, f.name)

        first = reply("tool_use", text("読みます"), tool_use("t1", "read_file", {"path": f.name}))
        client = FakeClient(first, reply("end_turn", text("中身は「こんにちは」です")))

        messages = harness.run_agent("ファイルを読んで", client, ask=never_ask)

        self.assertEqual(len(client.requests), 2)
        # モデルの返事は手を加えずに履歴に入っている
        self.assertIs(messages[1]["content"], first.content)
        # 2 回目のリクエストの最後に、ツールの結果が入っている
        result = client.requests[1]["messages"][-1]
        self.assertEqual(result["role"], "user")
        self.assertEqual(result["content"], [{
            "type": "tool_result",
            "tool_use_id": "t1",
            "content": "こんにちは",
            "is_error": False,
        }])

    def test_同時に頼まれたツールの結果は1つのメッセージで返す(self):
        client = FakeClient(
            reply("tool_use",
                  tool_use("a", "read_file", {"path": "ない1.txt"}),
                  tool_use("b", "read_file", {"path": "ない2.txt"})),
            reply("end_turn", text("どちらもありません")),
        )

        harness.run_agent("2 つ読んで", client, ask=never_ask)

        results = client.requests[1]["messages"][-1]["content"]
        self.assertEqual([r["tool_use_id"] for r in results], ["a", "b"])
        self.assertTrue(all(r["is_error"] for r in results))

    def test_コマンドは許可されたときだけ実行する(self):
        command = f'"{sys.executable}" -c "print(123)"'

        output, is_error = harness.execute_tool("run_command", {"command": command}, ask=lambda p: "n")
        self.assertTrue(is_error)
        self.assertIn("断りました", output)

        output, is_error = harness.execute_tool("run_command", {"command": command}, ask=lambda p: "y")
        self.assertFalse(is_error)
        self.assertEqual(output.strip(), "123")

    def test_失敗したコマンドはエラーとして返す(self):
        command = f'"{sys.executable}" -c "import sys; sys.exit(1)"'
        _, is_error = harness.execute_tool("run_command", {"command": command}, ask=lambda p: "y")
        self.assertTrue(is_error)

    def test_知らないツールはエラーとして返す(self):
        output, is_error = harness.execute_tool("delete_everything", {}, ask=never_ask)
        self.assertTrue(is_error)
        self.assertIn("知らないツール", output)

    def test_ツールを呼び続けても上限で止まる(self):
        client = FakeClient(reply("tool_use", tool_use("x", "read_file", {"path": "ない.txt"})))

        harness.run_agent("ずっと読んで", client, ask=never_ask, max_turns=3)

        self.assertEqual(len(client.requests), 3)

    def test_断られたなど_tool_use_以外で終わったら止まる(self):
        client = FakeClient(reply("refusal"), reply("end_turn", text("呼ばれないはず")))

        harness.run_agent("何か", client, ask=never_ask)

        self.assertEqual(len(client.requests), 1)


if __name__ == "__main__":
    unittest.main()
