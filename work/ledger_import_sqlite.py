import argparse
import datetime as dt
import sqlite3
import sys
from pathlib import Path

from ledger_stats import extract_rows, number_value, sort_rows_by_detection_date


COMMON_COLUMNS = [
    "来源文件",
    "工作表",
    "检测类型",
    "检测单位",
    "施工单位",
    "原始行号",
    "委托单位",
    "工程名称",
    "单位工程",
    "分部工程",
    "单元工程",
    "工程部位",
    "委托编号",
    "报告编号",
    "注浆日期",
    "委托日期",
    "检测日期",
    "报告日期",
    "批准日期",
    "检测结果",
    "施工数量",
    "检测数量",
    "备注",
]


def to_db_value(value):
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if value == "":
        return None
    return str(value)


def connect(db_path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(conn):
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS import_batch (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT NOT NULL,
            project_name TEXT,
            section_name TEXT,
            source_type TEXT,
            discipline TEXT,
            imported_at TEXT NOT NULL,
            total_rows INTEGER NOT NULL DEFAULT 0,
            skipped_sheets INTEGER NOT NULL DEFAULT 0,
            remark TEXT
        );

        CREATE TABLE IF NOT EXISTS detection_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL REFERENCES import_batch(id) ON DELETE CASCADE,
            project_name TEXT,
            section_name TEXT,
            source_type TEXT,
            discipline TEXT,
            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER,
            detection_type TEXT,
            detection_unit TEXT,
            construction_unit TEXT,
            entrust_unit TEXT,
            project_full_name TEXT,
            unit_project TEXT,
            sub_project TEXT,
            item_project TEXT,
            work_part TEXT,
            entrust_no TEXT,
            report_no TEXT,
            grouting_date TEXT,
            entrust_date TEXT,
            detection_date TEXT,
            report_date TEXT,
            approval_date TEXT,
            result TEXT,
            construction_qty REAL,
            detection_qty REAL,
            remark TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS detection_extra_field (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ledger_id INTEGER NOT NULL REFERENCES detection_ledger(id) ON DELETE CASCADE,
            field_name TEXT NOT NULL,
            field_value TEXT,
            source_col INTEGER,
            source_col_letter TEXT,
            sort_order INTEGER NOT NULL DEFAULT 0
        );

        CREATE INDEX IF NOT EXISTS idx_detection_ledger_date
            ON detection_ledger(detection_date);
        CREATE INDEX IF NOT EXISTS idx_detection_ledger_type
            ON detection_ledger(detection_type);
        CREATE INDEX IF NOT EXISTS idx_detection_ledger_section
            ON detection_ledger(section_name);
        CREATE INDEX IF NOT EXISTS idx_detection_ledger_report
            ON detection_ledger(report_no);
        """
    )
    ensure_column(conn, "detection_extra_field", "source_col", "INTEGER")
    ensure_column(conn, "detection_extra_field", "source_col_letter", "TEXT")


def ensure_column(conn, table_name, column_name, column_type):
    columns = {row[1] for row in conn.execute(f"PRAGMA table_info({table_name})")}
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")


def create_batch(conn, input_path, args, row_count, skipped_count):
    now = dt.datetime.now().isoformat(timespec="seconds")
    cur = conn.execute(
        """
        INSERT INTO import_batch (
            source_file, project_name, section_name, source_type, discipline,
            imported_at, total_rows, skipped_sheets, remark
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(input_path),
            args.project,
            args.section,
            args.source_type,
            args.discipline,
            now,
            row_count,
            skipped_count,
            args.remark,
        ),
    )
    return cur.lastrowid, now


def insert_ledger_rows(conn, batch_id, rows, args, created_at):
    sql = """
        INSERT INTO detection_ledger (
            batch_id, project_name, section_name, source_type, discipline,
            source_file, source_sheet, source_row, detection_type,
            detection_unit, construction_unit, entrust_unit, project_full_name,
            unit_project, sub_project, item_project, work_part,
            entrust_no, report_no, grouting_date, entrust_date, detection_date,
            report_date, approval_date, result, construction_qty, detection_qty,
            remark, created_at
        )
        VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
    """
    inserted = 0
    for row in sort_rows_by_detection_date(rows):
        cur = conn.execute(
            sql,
            (
                batch_id,
                args.project,
                args.section,
                args.source_type,
                args.discipline,
                to_db_value(row.get("来源文件")),
                to_db_value(row.get("工作表")),
                row.get("原始行号"),
                to_db_value(row.get("检测类型")),
                to_db_value(row.get("检测单位")),
                to_db_value(row.get("施工单位")),
                to_db_value(row.get("委托单位")),
                to_db_value(row.get("工程名称")),
                to_db_value(row.get("单位工程")),
                to_db_value(row.get("分部工程")),
                to_db_value(row.get("单元工程")),
                to_db_value(row.get("工程部位")),
                to_db_value(row.get("委托编号")),
                to_db_value(row.get("报告编号")),
                to_db_value(row.get("注浆日期")),
                to_db_value(row.get("委托日期")),
                to_db_value(row.get("检测日期")),
                to_db_value(row.get("报告日期")),
                to_db_value(row.get("批准日期")),
                to_db_value(row.get("检测结果")),
                number_value(row.get("施工数量")) or None,
                number_value(row.get("检测数量")) or None,
                to_db_value(row.get("备注")),
                created_at,
            ),
        )
        ledger_id = cur.lastrowid
        insert_extra_fields(conn, ledger_id, row)
        inserted += 1
    return inserted


def insert_extra_fields(conn, ledger_id, row):
    extras = row.get("扩展字段", [])
    for index, extra in enumerate(extras, start=1):
        conn.execute(
            """
            INSERT INTO detection_extra_field (
                ledger_id, field_name, field_value, source_col, source_col_letter, sort_order
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                ledger_id,
                extra.get("字段名"),
                to_db_value(extra.get("字段值")),
                extra.get("列号"),
                extra.get("列名"),
                index,
            ),
        )


def print_summary(conn, batch_id):
    total = conn.execute(
        "SELECT COUNT(*) FROM detection_ledger WHERE batch_id = ?",
        (batch_id,),
    ).fetchone()[0]
    print(f"本批次入库明细：{total}")

    print("按检测类型统计：")
    for row in conn.execute(
        """
        SELECT detection_type, COUNT(*), SUM(CASE WHEN result LIKE '%合格%' AND result NOT LIKE '%不%' THEN 1 ELSE 0 END),
               SUM(CASE WHEN result LIKE '%不%' THEN 1 ELSE 0 END)
        FROM detection_ledger
        WHERE batch_id = ?
        GROUP BY detection_type
        ORDER BY detection_type
        """,
        (batch_id,),
    ):
        print(f"  {row[0]}：{row[1]} 条，合格 {row[2] or 0}，不合格 {row[3] or 0}")


def main():
    parser = argparse.ArgumentParser(description="Excel检测台账导入SQLite数据库")
    parser.add_argument("input", help="输入Excel台账文件路径")
    parser.add_argument("--db", default="outputs/ledger.db", help="SQLite数据库输出路径")
    parser.add_argument("--project", default="湖南安化抽水蓄能电站", help="项目名称")
    parser.add_argument("--section", default="Q2标", help="标段名称")
    parser.add_argument("--source-type", default="施工检测", help="资料来源类型，例如：施工检测/监理检测")
    parser.add_argument("--discipline", default="物探", help="检测专业，例如：物探/试验/测量")
    parser.add_argument("--remark", default="", help="导入备注")
    args = parser.parse_args()

    input_path = Path(args.input)
    db_path = Path(args.db)
    if not input_path.exists():
        print(f"输入文件不存在：{input_path}", file=sys.stderr)
        sys.exit(1)

    rows, skipped = extract_rows(input_path)
    conn = connect(db_path)
    try:
        init_db(conn)
        with conn:
            batch_id, created_at = create_batch(conn, input_path, args, len(rows), len(skipped))
            inserted = insert_ledger_rows(conn, batch_id, rows, args, created_at)
        print(f"导入批次ID：{batch_id}")
        print(f"解析明细记录：{len(rows)}")
        print(f"实际入库记录：{inserted}")
        print(f"跳过工作表：{len(skipped)}")
        print(f"数据库文件：{db_path.resolve()}")
        print_summary(conn, batch_id)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
