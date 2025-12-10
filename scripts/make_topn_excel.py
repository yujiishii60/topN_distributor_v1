import argparse
import sys
from importlib import import_module

# 既存の軽量版を呼び出して処理を実行するだけのスタブ
# --template / --template-sheet は受け取るが、現段階では未使用でスルー

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--event-name", default="")
    p.add_argument("--title-template", default=None)
    p.add_argument("--no-date-in-title", action="store_true")
    p.add_argument("--category", required=True)
    p.add_argument("--dates", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--split-by-store", action="store_true")
    p.add_argument("--split-dir", default=None)
    # 将来のための受け口（今は未使用）
    p.add_argument("--template", default=None)
    p.add_argument("--template-sheet", default=None)
    args = p.parse_args()

    # 既存の軽量版へフォワード
    mod = import_module("scripts.make_topn_simple_refactor")
    # そのまま sys.argv を差し替えて実行（未対応引数は軽量版が無視 or argparseで弾かないようにする）
    # → argparseが厳密な場合は、必要な引数のみを再構築して main に渡す。
    if hasattr(mod, "main"):
        # 軽量版mainに必要な引数だけを構築
        fwd = [
            "--event-name", args.event_name or "",
            "--category", args.category,
            "--dates", args.dates,
            "--out", args.out,
        ]
        if args.split_by_store:
            fwd += ["--split-by-store"]
        if args.split_dir:
            fwd += ["--split-dir", args.split_dir]
        # 軽量版の main を直接呼ぶ（想定: def main(argv=None)）
        mod.main(fwd)
    else:
        # もし main() を公開していない場合は -m 実行にフォールバック
        import runpy
        sys.argv = [
            "python", "-m", "scripts.make_topn_simple_refactor",
            "--event-name", args.event_name or "",
            "--category", args.category,
            "--dates", args.dates,
            "--out", args.out,
        ]
        if args.split_by_store:
            sys.argv.append("--split-by-store")
        if args.split_dir:
            sys.argv += ["--split-dir", args.split_dir]
        runpy.run_module("scripts.make_topn_simple_refactor", run_name="__main__")

if __name__ == "__main__":
    main()
