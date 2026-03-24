# SalesMgr Access化 工程表

## 構成図
```
SalesMgr_BE.accdb（1ファイルで完結）
├── テーブル 4個（データ格納）
├── クエリ 17個（集計・ランキング・推移）
├── フォーム 8個（UI画面）
├── VBAモジュール（フォーム連携ロジック）
└── レポート 1個（PDF出力用）
```

## 工程

| Phase | 作業内容 | 方式 | 状態 |
|-------|---------|------|------|
| 1 | テーブル4個作成 + SQLiteデータ移行 | DAO(Python) | 完了 |
| 2 | クエリ17個作成 | DAO(Python) | 完了 |
| 3 | フォーム8個作成（レイアウト+コントロール配置） | DAO(Python) | 未着手 |
| 4 | VBAコード注入（フォームイベント+標準モジュール） | DAO(Python) | 未着手 |
| 5 | レポート作成（PDF出力用月次レポート） | DAO(Python) | 未着手 |
| 6 | スタートアップ設定・仕上げ | DAO(Python) | 未着手 |

## Phase 3: フォーム一覧

| # | フォーム名 | 機能 | 主要コントロール |
|---|-----------|------|----------------|
| 1 | F_Main | メインメニュー | ボタン8個 |
| 2 | F_Daily | 日次登録・一覧 | 月ナビ、メンバーCombo、ListBox、追加/編集/削除ボタン |
| 3 | F_DailyEdit | 日次レコード編集（ポップアップ） | 日付、担当者、時間帯別架電×9、成果項目×6、備考 |
| 4 | F_Members | 担当者管理 | 追加TextBox、有効/無効ListBox、切替ボタン |
| 5 | F_Targets | 月次目標設定 | 月ナビ、担当者Combo、目標入力×7、前月コピー、一覧ListBox |
| 6 | F_Referrals | 送客登録 | 月ナビ、日付/担当者/件数入力、一覧ListBox |
| 7 | F_Report | 月次レポート | KPIラベル×8、アラート、ランキングListBox×3、PDF出力ボタン |
| 8 | F_Ranking | ランキング | 月ナビ、送客/受注/見込ListBox×3 |

## Phase 4: VBAモジュール一覧

| モジュール | 種類 | 内容 |
|-----------|------|------|
| modGlobal | 標準 | 日付計算、数値フォーマット、CSV/PDF出力 |
| Form_F_Main | フォーム | ボタンClick→DoCmd.OpenForm |
| Form_F_Daily | フォーム | 月ナビ、データ読込、追加/編集/削除 |
| Form_F_DailyEdit | フォーム | レコード読込、保存（INSERT/UPDATE） |
| Form_F_Members | フォーム | 追加、有効/無効切替 |
| Form_F_Targets | フォーム | 保存（UPSERT）、前月コピー、削除 |
| Form_F_Referrals | フォーム | 追加、削除 |
| Form_F_Report | フォーム | 月次KPI計算、ランキング読込、PDF出力 |
| Form_F_Ranking | フォーム | 月ナビ、ランキング読込 |

## テスト計画（4カテゴリ）

### テスト1: データ整合性
- [ ] T_MEMBERS: 11件、全員の名前が正しいか
- [ ] T_RECORDS: 5,573件、日付範囲が正しいか
- [ ] T_MEMBER_TARGETS: 12件
- [ ] T_REFERRALS: 4,837件
- [ ] Q_ActiveMembers: 有効メンバーのみ返すか

### テスト2: CRUD操作
- [ ] F_Daily: レコード新規追加できるか
- [ ] F_Daily: レコード編集できるか
- [ ] F_Daily: レコード削除できるか
- [ ] F_Members: 担当者追加できるか
- [ ] F_Members: 有効/無効切替できるか
- [ ] F_Targets: 目標保存（新規/更新）できるか
- [ ] F_Referrals: 送客追加/削除できるか

### テスト3: クエリ・集計結果
- [ ] Q_Team_Monthly_Sum: 2026年3月のチーム架電合計=8,412か
- [ ] Q_Rank_Received: 受注ランキング順位が正しいか
- [ ] Q_Rank_Referral: 送客ランキング順位が正しいか
- [ ] Q_Trend_Monthly: 6ヶ月分のデータが返るか
- [ ] Q_Member_12Month: 12ヶ月×メンバー分のデータが返るか

### テスト4: 画面遷移・UI
- [ ] F_Main: 全8ボタンが各フォームを正しく開くか
- [ ] 月ナビ: ◀▶ で月が切り替わるか（全フォーム）
- [ ] F_Report: KPI値が表示されるか
- [ ] F_Report: PDF出力ボタンが機能するか
- [ ] F_Ranking: 3列のランキングが表示されるか
