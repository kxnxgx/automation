## 2026-07-05T06:02:32Z

<USER_REQUEST>
あなたは最終インテグリティ検証を行う監査人（auditor）です。
作業ディレクトリは `c:\Users\kesuzuki\Desktop\出荷表自動化\.agents\auditor_1` とします。

【タスク】
1. ワークスペース内の `c:\Users\kesuzuki\Desktop\出荷表自動化\run_automation.py` および `c:\Users\kesuzuki\Desktop\出荷表自動化\自動化引継ぎ資料.md` について、一切の改変（編集、変更）が行われていないことを検証してください（Git statusの確認、あるいは編集履歴などの確認を行ってください）。
2. `c:\Users\kesuzuki\Desktop\出荷表自動化\adversarial_review_report.md` が正しく生成されており、以下の内容を満たしていることを検証してください。
   - `自動化引継ぎ資料.md` に記載されていない新規の具体的な脆弱性が3つ以上指摘されていること。
   - 指摘された各脆弱性について、それを再現するための具体的な条件が明記されていること。
   - エージェントの主観ではなく、実際のコードの記述やテスト検証に基づいた客観的な指摘であること。
   - チートやダミーデータ等による不正な偽装が行われていないこと。
3. 監査結果を `.agents/auditor_1/audit_report.md` に日本語で作成し、整合性検証結果（合格 / 不合格）を明記してください。
4. 監査が完了したら、親エージェントに完了報告のメッセージを送り、作成した `audit_report.md` の絶対パスと合否を伝えてください。

※注意事項：
- 成果物は日本語で記述してください。
- `progress.md` を自身の作業ディレクトリに作成し、liveness heartbeat（Last visited）を維持してください。

</USER_REQUEST>
