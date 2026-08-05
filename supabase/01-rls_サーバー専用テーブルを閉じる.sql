-- 01-rls_サーバー専用テーブルを閉じる.sql
-- 対象プロジェクト: xenzkjsptyqfzseqdwny（相乗りの共有DB）
-- ★ RoutMe の rvbxbkrupamsqdppesrd では絶対に実行しない。
--
-- ■ 何をするか
--   「ブラウザから触る必要が一切ないテーブル」を Data API から締め出す。
--   RLS を有効化し、**ポリシーを1本も作らない**（＝ anon / authenticated からは1行も見えず1行も書けない）。
--   さらに権限そのものを剥がす（REVOKE）。ポリシーの有無に関係なく拒否されるので二重で確実。
--
-- ■ 誰が・どの行に・何をできるか（このファイル適用後）
--   ・anon（＝公開されている publishable キー。誰でも入手できる）      … 何もできない
--   ・authenticated（ログイン済みユーザー）                            … 何もできない
--   ・service_role（サーバー側の鍵。RLS をバイパスする）               … 従来どおり全部できる
--   ・テーブル所有者 / postgres                                        … 従来どおり
--
-- ■ アプリへの影響： なし
--   daily_sales / weekly_sales に書いているのは aggregate.py だけで、
--   SUPABASE_SERVICE_ROLE_KEY を使っている（aggregate.py:565, 621）。service_role は RLS を
--   バイパスするので upsert は通り続ける。読み手はローカルの HTML 生成のみで DB を読まない。
--
-- ■ なぜ必要か（実測）
--   公開 config.js から取れる anon キーで、daily_sales 10,386行 / weekly_sales 3,019行が
--   そのまま全件読めた。約184店舗ぶんの卸売上（週1000万円規模）。
--
-- ■ 手動適用: Supabase SQL Editor で xenzkjsptyqfzseqdwny を選んで実行。
--   ★ SQL Editor は長文の貼り付けが途中で切れることがある。BEGIN;…COMMIT; で包んであるので、
--     切れた場合は半端に適用されず全部ロールバックされる（そのまま貼り直せばよい）。

SET lock_timeout = '5s';

BEGIN;

-- ── ① 旬彩坊 週次集計（service_role のバッチ専用） ──────────────────
-- 売上金額そのもの。ブラウザからは一切触らない。
ALTER TABLE public.daily_sales  ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON TABLE public.daily_sales  FROM anon, authenticated;

ALTER TABLE public.weekly_sales ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON TABLE public.weekly_sales FROM anon, authenticated;

-- ── ② レガシー（manual_stops / delivery_logs / order_imports / incidents）は対象外にした ──
-- 当初は「0行だから塞いでおく」つもりだったが、Security Advisor の画面で確認したところ
-- Critical(RLS Disabled) に挙がっているのは routes / weekly_sales / daily_sales /
-- media_stores / pv_records の5つだけで、この4つは出ていない。
-- 代わりに「Auth RLS Initialization Plan」（ポリシーがある表にだけ出る性能警告）側に載っている
-- ＝ すでに RLS 有効かつポリシーあり。触る必要がない。
--
-- routes だけは兄弟テーブル（stores / drivers / contacts）と同じ扱いに揃えたいので、
-- それらの既存ポリシーを 00-確認_RLSの現状.sql の③で見てから別ファイルで対応する。
-- ここで安易に REVOKE すると、万一まだ動いている旧アプリで routes だけが落ちて挙動が不揃いになる。

COMMIT;

-- PostgREST にスキーマ再読込を通知（これを忘れると Data API の一覧に残り続ける）
NOTIFY pgrst, 'reload schema';


-- ═══════════════════════════════════════════════════════════════════════════
-- ■ 適用後の確認（SQL Editor で実行）
-- ═══════════════════════════════════════════════════════════════════════════
--   ▼ 2テーブルとも rls有効=true・ポリシー数=0 になっていること
--     SELECT c.relname, c.relrowsecurity,
--            (SELECT count(*) FROM pg_policies p WHERE p.tablename = c.relname) AS policies
--     FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
--     WHERE n.nspname='public'
--       AND c.relname IN ('daily_sales','weekly_sales');
--
--   ▼ anon / authenticated の権限が1行も残っていないこと（期待: 0件）
--     SELECT table_name, grantee, privilege_type FROM information_schema.role_table_grants
--     WHERE table_schema='public' AND grantee IN ('anon','authenticated')
--       AND table_name IN ('daily_sales','weekly_sales');
--
--   ▼ バッチが壊れていないこと（ターミナルで手動実行）
--     cd ~/Documents/旬彩坊_週次集計 && .venv/bin/python aggregate.py
--
-- ═══════════════════════════════════════════════════════════════════════════
-- ■ ロールバック（元に戻す）
-- ═══════════════════════════════════════════════════════════════════════════
--   BEGIN;
--   GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.daily_sales, public.weekly_sales
--     TO anon, authenticated;
--   ALTER TABLE public.daily_sales   DISABLE ROW LEVEL SECURITY;
--   ALTER TABLE public.weekly_sales  DISABLE ROW LEVEL SECURITY;
--   COMMIT;
--   NOTIFY pgrst, 'reload schema';
