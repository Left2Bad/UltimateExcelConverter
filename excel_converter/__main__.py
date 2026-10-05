import argparse
import sys
from pathlib import Path

from .core import convert, export_result, headers, load_profile, load_source, suggest_header, suggest_mapping, write_report
from .demo import create_demo
from .example_pack import create_example_pack


def main():
    parser = argparse.ArgumentParser(description="Локальная конвертация Excel. Без аргументов — графический интерфейс.")
    sub = parser.add_subparsers(dest="command")
    demo = sub.add_parser("demo", help="Создать синтетические примеры")
    demo.add_argument("directory", type=Path)
    examples = sub.add_parser('examples', help='Создать расширенный набор из 24 учебных файлов')
    examples.add_argument('directory', type=Path)
    cli = sub.add_parser("convert", help="Проверить и преобразовать один лист")
    cli.add_argument("input", type=Path)
    cli.add_argument("output", type=Path)
    cli.add_argument("--profile", type=Path)
    cli.add_argument("--sheet")
    cli.add_argument("--header", type=int)
    cli.add_argument("--report", type=Path, help="Сохранить отчёт проверки, в том числе при ошибках; новый файл")
    cli.add_argument("--accept-warnings", action="store_true", help="Подтвердить сохранение при предупреждениях")
    args = parser.parse_args()
    try:
        if args.command == 'examples':
            paths = create_example_pack(args.directory)
            print(f'Учебных файлов: {len(paths)}. Каталог: {args.directory / "КАТАЛОГ.md"}')
        elif args.command == "demo":
            for path in create_demo(args.directory):
                print(path)
        elif args.command == "convert":
            profile = load_profile(args.profile) if args.profile else load_profile()
            source = load_source(args.input)
            if not args.sheet and len(source.sheets) > 1:
                raise ValueError("В файле несколько листов. Укажите --sheet.")
            sheet = args.sheet or next(iter(source.sheets))
            if sheet not in source.sheets:
                raise ValueError("Лист не найден. Доступны: " + ", ".join(source.sheets))
            header = args.header if args.header is not None else suggest_header(source.sheets[sheet], profile)
            if header is None:
                raise ValueError("Строка заголовков не определена однозначно. Укажите --header.")
            mapping = suggest_mapping(headers(source.sheets[sheet], header), profile)
            result = convert(source, sheet, header, profile, mapping)
            for issue in result.issues:
                print(f"{issue.severity}: строка {issue.row or '—'}, {issue.field}: {issue.message}")
            print(f"Записей: {result.read_rows}; корректных: {len(result.records)}; ошибок: {len(result.errors)}")
            if args.report:
                write_report(result, args.report)
            if result.errors:
                return 2
            if result.issues and not args.accept_warnings:
                print("Просмотрите предупреждения. Для выгрузки добавьте --accept-warnings.")
                return 3
            output, report = export_result(result, args.output)
            print(f"Сохранено: {output}\nОтчёт: {report}")
        else:
            from .gui import launch
            launch()
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
