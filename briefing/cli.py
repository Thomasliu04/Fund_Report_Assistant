"""命令行入口"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="季度简报辅助系统")
    sub = parser.add_subparsers(dest="command")

    gen = sub.add_parser("generate", help="生成 HTML 简报")
    gen.add_argument("--config", required=True, help="报告配置 YAML")
    gen.add_argument("--data", required=True, help="数据目录")
    gen.add_argument("--output", required=True, help="输出 HTML 路径")

    p1 = sub.add_parser("phase1", help="Phase1：灌入 PPT 第1-2页（文字+表图）")
    p1.add_argument("--config", default="config/report_25h1.yaml")
    p1.add_argument("--data", default="data/sample_25h1")
    p1.add_argument("--slide-map", default="config/slide_map.yaml")
    p1.add_argument("--template", default="templates/25H1_template.pptx")
    p1.add_argument("--output", default="output/25H1_phase1.pptx")

    loop = sub.add_parser("excel-loop", help="最小闭环：Excel+表图 → PPT第2页")
    loop.add_argument("--config", default="config/report_25h1.yaml")
    loop.add_argument("--data", default="data/sample_25h1")
    loop.add_argument("--template", default="templates/25H1_template.pptx")
    loop.add_argument("--output", default="output/25H1_excel_loop.pptx")

    full = sub.add_parser("full-deck", help="全册10页：文字+表图灌入 PPT")
    full.add_argument("--config", default="config/report_25h1.yaml")
    full.add_argument("--data", default="data/sample_25h1")
    full.add_argument("--template", default="templates/25H1_template.pptx")
    full.add_argument("--slide-map", default="config/slide_map.yaml")
    full.add_argument("--output", default="output/25H1_full_deck.pptx")
    full.add_argument("--draft-xlsx", default="", help="底稿终表 xlsx；提供则以终表生成表图")

    imp = sub.add_parser("import-xlsx", help="从简报底稿 xlsx 导入 fund_metrics.csv")
    imp.add_argument("--xlsx", required=True, help="底稿 xlsx 路径")
    imp.add_argument("--out-dir", required=True, help="输出数据目录（写入 fund_metrics.csv）")

    tables = sub.add_parser("export-tables", help="以底稿终表为基准导出表图 Excel")
    tables.add_argument("--xlsx", required=True, help="底稿 xlsx（含终表）")
    tables.add_argument("--output", required=True, help="输出 xx_tables.xlsx")
    tables.add_argument("--period", default="", help="期别标签，如 25Q4")
    tables.add_argument("--focus", default="银华", help="高亮公司简称")

    serve = sub.add_parser("serve", help="启动网页操作台")
    serve.add_argument("--host", default="127.0.0.1")
    serve.add_argument("--port", type=int, default=8765)
    serve.add_argument("--open", action="store_true", help="启动后自动打开浏览器")

    args = parser.parse_args(argv)
    if args.command == "generate":
        from briefing.pipeline import generate_briefing

        out = generate_briefing(args.config, args.data, args.output)
        print(f"已生成: {args.output}")
        print(f"  标题: {out.config.title}")
        print(f"  章节: {len(out.sections)}")
        return 0

    if args.command == "phase1":
        from briefing.phase1 import run_phase1

        out = run_phase1(
            config_path=args.config,
            data_dir=args.data,
            slide_map_path=args.slide_map,
            template_path=args.template,
            output_path=args.output,
        )
        print(f"已生成 Phase1 PPT: {out}")
        return 0

    if args.command == "excel-loop":
        from briefing.excel_ppt_loop import run_excel_ppt_loop

        info = run_excel_ppt_loop(
            config_path=args.config,
            data_dir=args.data,
            template_path=args.template,
            output_pptx=args.output,
        )
        print("Excel→PPT 最小闭环完成:")
        for k, v in info.items():
            print(f"  {k}: {v}")
        return 0

    if args.command == "full-deck":
        from briefing.full_deck import run_full_deck

        info = run_full_deck(
            config_path=args.config,
            data_dir=args.data,
            template_path=args.template,
            slide_map_path=args.slide_map,
            output_pptx=args.output,
            draft_xlsx=args.draft_xlsx or None,
        )
        print("全册 PPT 生成完成:")
        for k, v in info.items():
            if k == "images":
                print(f"  images: {len(v)} 张")
            else:
                print(f"  {k}: {v}")
        return 0

    if args.command == "import-xlsx":
        from pathlib import Path

        from briefing.importers.xlsx_draft import detect_dates, import_draft_xlsx

        out_dir = Path(args.out_dir)
        out_csv = out_dir / "fund_metrics.csv"
        df = import_draft_xlsx(args.xlsx, output_csv=out_csv)
        dates = detect_dates(args.xlsx)
        print("底稿导入完成:")
        print(f"  xlsx: {args.xlsx}")
        print(f"  csv: {out_csv}")
        print(f"  rows: {len(df)}")
        print(f"  dates: {dates}")
        print(f"  categories: {sorted(df['category'].unique())}")
        print("请确认 config 中 dates 与上述一致后执行 full-deck。")
        return 0

    if args.command == "export-tables":
        from briefing.deck.final_tables import export_tables_from_final_xlsx

        out = export_tables_from_final_xlsx(
            args.xlsx,
            args.output,
            period_label=args.period,
            focus_company=args.focus,
        )
        print("终表导出完成:")
        print(f"  底稿: {args.xlsx}")
        print(f"  表图: {out}")
        return 0

    if args.command == "serve":
        import threading
        import time
        import webbrowser

        import uvicorn

        url = f"http://{args.host}:{args.port}"
        print(f"网页操作台：{url}")
        print("按 Ctrl+C 结束")

        if args.open:
            def _open():
                time.sleep(0.8)
                webbrowser.open(url)

            threading.Thread(target=_open, daemon=True).start()

        uvicorn.run(
            "briefing.web.app:app",
            host=args.host,
            port=args.port,
            reload=False,
        )
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
