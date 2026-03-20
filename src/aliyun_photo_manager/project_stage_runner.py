from __future__ import annotations

import json
import sys
from pathlib import Path

from .project_stage_report import (
    StageServerConfig,
    dump_status_query_payload,
    query_project_stages,
    summary_to_dict,
)


def main() -> int:
    if len(sys.argv) != 3:
        raise SystemExit("usage: python -m aliyun_photo_manager.project_stage_runner <input.json> <output.json>")

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    payload = json.loads(input_path.read_text(encoding="utf-8"))
    servers = [StageServerConfig(**item) for item in payload.get("servers", [])]
    summary = query_project_stages(
        servers=servers,
        status_filter=str(payload.get("status_filter", "正在进行 + 即将开始")),
        stage_keyword=str(payload.get("stage_keyword", "")),
        project_keyword=str(payload.get("project_keyword", "")),
    )
    output_path.write_text(
        json.dumps(summary_to_dict(summary), ensure_ascii=False),
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
