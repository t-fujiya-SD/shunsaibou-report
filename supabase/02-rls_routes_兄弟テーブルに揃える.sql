-- 02-rls_routes_兄弟テーブルに揃える.sql
-- 対象プロジェクト: xenzkjsptyqfzseqdwny（相乗りの共有DB）
-- ★ RoutMe の rvbxbkrupamsqdppesrd では実行しない。
--
-- ■ 何をするか
--   旧・旬彩坊配送アプリのテーブルのうち、routes だけ RLS が無効で置き去りになっていた
--   （Security Advisor の Critical に public.routes として出ている）。
--   兄弟の stores / drivers / contacts と**同じ形**のテナント分離ポリシーを付けて揃える。
--
-- ■ 誰が・どの行に・何をできるか（適用後）
--   ・anon（公開されている publishable キー）
--       auth.uid() が NULL → 副問い合わせが NULL → tenant_id = NULL は真にならない
--       → **1行も見えず、1行も書けない**（stores などと同じ挙動。実測でも stores は anon から0件）
--   ・ログイン済みユーザー
--       自分の users.tenant_id と一致する行だけ、SELECT / INSERT / UPDATE / DELETE できる
--       （FOR ALL で WITH CHECK を省略すると USING が書き込み側の検査にも使われる＝兄弟と同じ）
--   ・service_role … RLS をバイパスするので従来どおり全部
--
-- ■ 兄弟テーブルの既存ポリシー（これに合わせている）
--   tenant_isolation_stores / _drivers / _contacts
--     FOR ALL TO public
--     USING (tenant_id = (SELECT users.tenant_id FROM users WHERE users.id = auth.uid()))
--
-- ■ 1点だけ書き方を変えている（意味は同じ）
--   auth.uid() を (SELECT auth.uid()) で包んだ。行ごとに評価されるのを防ぐための定石で、
--   包まないと Advisor の「Auth RLS Initialization Plan」警告が1件増える
--   （兄弟3つは今まさにその警告が出ている）。判定結果は完全に同じ。
--
-- ■ 適用前に確認済み
--   routes 15行すべて tenant_id='shunsaibo'、users も 'shunsaibo' の1件のみ。
--   → ポリシーを付けても、ログイン済みユーザーからは今までどおり15行すべて見える。
--     （tenant_id が NULL の行があると誰からも見えなくなるので、ここは必ず先に確認する）
--
-- ■ 手動適用: Supabase SQL Editor で xenzkjsptyqfzseqdwny を選んで実行。

SET lock_timeout = '5s';

BEGIN;

ALTER TABLE public.routes ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS tenant_isolation_routes ON public.routes;
CREATE POLICY tenant_isolation_routes ON public.routes
  FOR ALL
  TO public
  USING (
    tenant_id = (
      SELECT users.tenant_id FROM public.users
      WHERE users.id = (SELECT auth.uid())
    )
  );

COMMIT;

NOTIFY pgrst, 'reload schema';


-- ═══════════════════════════════════════════════════════════════════════════
-- ■ 適用後の確認
-- ═══════════════════════════════════════════════════════════════════════════
--   ▼ Security Advisor の Critical から public.routes が消えること（残るのは media_stores / pv_records）
--
--   ▼ ポリシーが兄弟と同じ形で入ったこと
--     SELECT tablename, policyname, cmd, roles, qual FROM pg_policies
--     WHERE schemaname='public' AND tablename IN ('routes','stores') ORDER BY tablename;
--
--   ▼ service_role では今までどおり15行見えること
--     SELECT count(*) FROM public.routes;   -- 期待: 15
--
-- ═══════════════════════════════════════════════════════════════════════════
-- ■ ロールバック
-- ═══════════════════════════════════════════════════════════════════════════
--   BEGIN;
--   DROP POLICY IF EXISTS tenant_isolation_routes ON public.routes;
--   ALTER TABLE public.routes DISABLE ROW LEVEL SECURITY;
--   COMMIT;
--   NOTIFY pgrst, 'reload schema';
