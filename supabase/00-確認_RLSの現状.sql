-- 00-確認_RLSの現状.sql
--
-- 共有Supabaseプロジェクト xenzkjsptyqfzseqdwny の RLS 現状を一覧する。**読み取りのみ。**
-- 相乗りしているアプリ: 厨房マニュアル(krm_) / デシャボ(dsb_) / 開店チェックリスト(ocl_)
--                       / 媒体PV測定(media_stores, pv_records) / 旬彩坊週次集計(daily_sales, weekly_sales)
--                       / 旧・旬彩坊配送アプリのレガシー(routes, stores, drivers, contacts, users, tenants ほか)
--
-- ★ RoutMe(partner-a-delivery-app) は別プロジェクト rvbxbkrupamsqdppesrd なので無関係。ここでは触らない。
--
-- 使い方: Supabase SQL Editor で **xenzkjsptyqfzseqdwny** を選び、そのまま実行する。
--         結果の「rls有効 = false」の行が、Security Advisor が毎週指摘しているテーブル。

-- ── ① テーブルごとの RLS とポリシー数 ────────────────────────────────
SELECT
  c.relname                                   AS "テーブル",
  c.relrowsecurity                            AS "rls有効",
  c.relforcerowsecurity                       AS "所有者にも強制",
  (SELECT count(*) FROM pg_policies p
     WHERE p.schemaname = 'public' AND p.tablename = c.relname) AS "ポリシー数",
  (SELECT reltuples::bigint FROM pg_class x WHERE x.oid = c.oid) AS "概算行数"
FROM pg_class c
JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE n.nspname = 'public' AND c.relkind = 'r'
ORDER BY c.relrowsecurity, c.relname;


-- ── ② anon / authenticated に付いているテーブル権限 ──────────────────
--    RLS を有効にしてもポリシーが無ければ拒否されるが、権限そのものを剥がすほうが確実。
--    「クライアントから触る必要がないテーブル」はここが空になっているのが理想。
SELECT
  table_name  AS "テーブル",
  grantee     AS "ロール",
  string_agg(privilege_type, ', ' ORDER BY privilege_type) AS "権限"
FROM information_schema.role_table_grants
WHERE table_schema = 'public'
  AND grantee IN ('anon', 'authenticated')
GROUP BY table_name, grantee
ORDER BY table_name, grantee;


-- ── ③ anon に全開になっているポリシーを洗い出す ──────────────────────
--    ★これが今回いちばん見落としやすいところ。
--    RLS が有効でも `for all to anon using (true) with check (true)`（pilot_all 等）が付いていると
--    実質 RLS 無効と同じ。Security Advisor は「RLS有効」と見なすので**警告に出てこない**。
SELECT
  tablename   AS "テーブル",
  policyname  AS "ポリシー名",
  cmd         AS "対象操作",
  roles       AS "対象ロール",
  qual        AS "USING句",
  with_check  AS "WITH CHECK句"
FROM pg_policies
WHERE schemaname = 'public'
  AND 'anon' = ANY (roles)
ORDER BY tablename, policyname;
