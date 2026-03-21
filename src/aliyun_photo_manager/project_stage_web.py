from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, Field

from .project_stage_report import (
    StageServerConfig,
    dump_status_query_payload,
    export_project_stages,
    summary_from_dict,
    summary_to_dict,
)


HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>项目阶段汇总服务</title>
  <style>
    body { font-family: "Microsoft YaHei", sans-serif; background:#f6f7fb; color:#0f172a; margin:0; }
    .page { max-width: 1280px; margin: 0 auto; padding: 24px; }
    h1 { margin: 0 0 16px; font-size: 28px; }
    .grid { display:grid; grid-template-columns: 340px 1fr; gap: 20px; }
    .card { background:#fff; border:1px solid #d8dee8; border-radius:12px; padding:16px; box-shadow:0 4px 18px rgba(15,23,42,.06); }
    .section-title { margin:0 0 12px; font-size:18px; font-weight:700; }
    label { display:block; margin:10px 0 6px; font-size:14px; }
    input, select, button { font:inherit; }
    input, select { width:100%; box-sizing:border-box; padding:10px 12px; border:1px solid #c4ccda; border-radius:8px; background:#fff; }
    button { padding:10px 14px; border:none; border-radius:8px; background:#2563eb; color:#fff; cursor:pointer; }
    button.secondary { background:#64748b; }
    button.ghost { background:#e5e7eb; color:#111827; }
    button:disabled { opacity:.5; cursor:not-allowed; }
    .actions { display:flex; flex-wrap:wrap; gap:10px; margin-top:12px; }
    .server-item { margin-top:10px; padding:12px; border:1px solid #d8dee8; border-radius:10px; background:#f8fafc; }
    .server-head { display:flex; justify-content:space-between; align-items:center; gap:10px; }
    .server-meta { margin-top:6px; color:#475569; font-size:13px; }
    .server-controls { display:flex; gap:8px; margin-top:10px; }
    .server-controls button { flex:1; }
    .toolbar { display:grid; grid-template-columns: repeat(3, minmax(160px, 1fr)); gap:10px; margin-bottom:12px; }
    .summary { display:grid; grid-template-columns: repeat(5, minmax(100px, 1fr)); gap:10px; margin-bottom:14px; }
    .metric { background:#f8fafc; border:1px solid #d8dee8; border-radius:10px; padding:12px; }
    .metric b { display:block; margin-top:6px; font-size:22px; }
    .status { min-height:24px; margin-top:10px; white-space:pre-wrap; color:#334155; }
    table { width:100%; border-collapse:collapse; }
    th, td { border:1px solid #d8dee8; padding:8px 10px; font-size:14px; text-align:left; }
    th { background:#f8fafc; position:sticky; top:0; }
    .results { max-height:620px; overflow:auto; }
    .muted { color:#64748b; font-size:13px; }
  </style>
</head>
<body>
  <div class="page">
    <h1>项目阶段汇总服务</h1>
    <div class="grid">
      <div class="card">
        <div class="section-title">服务器配置</div>
        <label>服务器名称</label>
        <input id="server-name" placeholder="例如：162">
        <label>数据库地址</label>
        <input id="server-host" placeholder="例如：192.168.1.162">
        <label>端口</label>
        <input id="server-port" value="1433">
        <label>用户名</label>
        <input id="server-user">
        <label>密码</label>
        <input id="server-password" type="password">
        <div class="actions">
          <button id="save-server">新增/更新</button>
          <button id="clear-server" class="ghost">清空表单</button>
        </div>
        <div class="actions">
          <button id="test-server" class="secondary">测试连接</button>
        </div>
        <div class="status" id="server-status"></div>
        <div class="section-title" style="margin-top:18px;">服务器列表</div>
        <div id="server-list"></div>
      </div>
      <div class="card">
        <div class="section-title">查询条件</div>
        <div class="toolbar">
          <select id="status-filter">
            <option>正在进行 + 即将开始</option>
            <option>全部</option>
            <option>只看正在进行</option>
            <option>只看即将开始</option>
          </select>
          <input id="stage-keyword" placeholder="阶段关键字，例如：报名">
          <input id="project-keyword" placeholder="项目关键字">
        </div>
        <div class="actions">
          <button id="query-btn">开始查询</button>
          <button id="export-btn" class="secondary" disabled>导出 Excel</button>
        </div>
        <div class="status" id="query-status">未开始查询</div>
        <div class="summary" id="summary-panel"></div>
        <div class="muted">结果只依赖同一套业务表，适合快速查看多服务器上的项目阶段状态。</div>
        <div class="results" style="margin-top:14px;">
          <table>
            <thead>
              <tr>
                <th>服务器</th>
                <th>数据库</th>
                <th>项目名称</th>
                <th>阶段名称</th>
                <th>开始时间</th>
                <th>结束时间</th>
                <th>当前状态</th>
              </tr>
            </thead>
            <tbody id="result-body">
              <tr><td colspan="7" class="muted">暂无结果</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
  <script>
    const storageKey = "project-stage-web-servers";
    let selectedIndex = null;
    let lastPayload = null;

    function loadServers() {
      try { return JSON.parse(localStorage.getItem(storageKey) || "[]"); }
      catch { return []; }
    }

    function saveServers(servers) {
      localStorage.setItem(storageKey, JSON.stringify(servers));
    }

    function fillForm(server) {
      document.getElementById("server-name").value = server?.name || "";
      document.getElementById("server-host").value = server?.host || "";
      document.getElementById("server-port").value = server?.port || "1433";
      document.getElementById("server-user").value = server?.username || "";
      document.getElementById("server-password").value = server?.password || "";
    }

    function renderServers() {
      const container = document.getElementById("server-list");
      const servers = loadServers();
      container.innerHTML = "";
      if (!servers.length) {
        container.innerHTML = '<div class="muted">还没有添加服务器。</div>';
        return;
      }
      servers.forEach((server, index) => {
        const item = document.createElement("div");
        item.className = "server-item";
        item.innerHTML = `
          <div class="server-head">
            <strong>${server.name}</strong>
            <label><input type="checkbox" data-action="toggle" data-index="${index}" ${server.enabled ? "checked" : ""}> 启用</label>
          </div>
          <div class="server-meta">${server.host}:${server.port} / ${server.username}</div>
          <div class="server-controls">
            <button class="ghost" data-action="edit" data-index="${index}">编辑</button>
            <button class="ghost" data-action="delete" data-index="${index}">删除</button>
          </div>`;
        container.appendChild(item);
      });
    }

    function setServerStatus(text) {
      document.getElementById("server-status").textContent = text;
    }

    function setQueryStatus(text) {
      document.getElementById("query-status").textContent = text;
    }

    function enabledServers() {
      return loadServers().filter(item => item.enabled);
    }

    function buildPayload() {
      return {
        servers: enabledServers(),
        status_filter: document.getElementById("status-filter").value,
        stage_keyword: document.getElementById("stage-keyword").value.trim(),
        project_keyword: document.getElementById("project-keyword").value.trim()
      };
    }

    function renderSummary(summary) {
      document.getElementById("summary-panel").innerHTML = `
        <div class="metric">启用服务器<b>${summary.enabled_servers}</b></div>
        <div class="metric">遍历数据库<b>${summary.visited_databases}</b></div>
        <div class="metric">匹配业务库<b>${summary.matched_databases}</b></div>
        <div class="metric">正在进行<b>${summary.ongoing_count}</b></div>
        <div class="metric">即将开始<b>${summary.upcoming_count}</b></div>`;
    }

    function renderRecords(records) {
      const body = document.getElementById("result-body");
      if (!records.length) {
        body.innerHTML = '<tr><td colspan="7" class="muted">没有符合条件的结果</td></tr>';
        return;
      }
      body.innerHTML = records.map(item => `
        <tr>
          <td>${item.server_name}</td>
          <td>${item.database_name}</td>
          <td>${item.project_name}</td>
          <td>${item.stage_name}</td>
          <td>${item.start_time.replace("T", " ").slice(0, 19)}</td>
          <td>${item.end_time.replace("T", " ").slice(0, 19)}</td>
          <td>${item.status}</td>
        </tr>`).join("");
    }

    async function postJson(url, payload) {
      const resp = await fetch(url, {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(payload)
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data.detail || "请求失败");
      return data;
    }

    document.getElementById("save-server").addEventListener("click", () => {
      const server = {
        name: document.getElementById("server-name").value.trim(),
        host: document.getElementById("server-host").value.trim(),
        port: Number(document.getElementById("server-port").value.trim() || "1433"),
        username: document.getElementById("server-user").value.trim(),
        password: document.getElementById("server-password").value.trim(),
        enabled: true
      };
      if (!server.name || !server.host || !server.username || !server.password) {
        setServerStatus("请完整填写服务器信息。");
        return;
      }
      const servers = loadServers();
      if (selectedIndex === null) servers.push(server);
      else servers[selectedIndex] = {...servers[selectedIndex], ...server};
      saveServers(servers);
      selectedIndex = null;
      fillForm(null);
      setServerStatus("服务器配置已保存。");
      renderServers();
    });

    document.getElementById("clear-server").addEventListener("click", () => {
      selectedIndex = null;
      fillForm(null);
      setServerStatus("");
    });

    document.getElementById("server-list").addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) return;
      const action = target.dataset.action;
      const index = Number(target.dataset.index);
      const servers = loadServers();
      if (Number.isNaN(index) || !servers[index]) return;
      if (action === "edit") {
        selectedIndex = index;
        fillForm(servers[index]);
      } else if (action === "delete") {
        servers.splice(index, 1);
        saveServers(servers);
        renderServers();
      }
    });

    document.getElementById("server-list").addEventListener("change", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.dataset.action !== "toggle") return;
      const index = Number(target.dataset.index);
      const servers = loadServers();
      if (Number.isNaN(index) || !servers[index]) return;
      servers[index].enabled = target.checked;
      saveServers(servers);
    });

    document.getElementById("test-server").addEventListener("click", async () => {
      try {
        setServerStatus("正在测试连接...");
        const result = await postJson("/api/test", {
          server: {
            name: document.getElementById("server-name").value.trim(),
            host: document.getElementById("server-host").value.trim(),
            port: Number(document.getElementById("server-port").value.trim() || "1433"),
            username: document.getElementById("server-user").value.trim(),
            password: document.getElementById("server-password").value.trim(),
            enabled: true
          }
        });
        setServerStatus(`连接成功。遍历数据库 ${result.visited_databases} 个，匹配业务库 ${result.matched_databases} 个。`);
      } catch (error) {
        setServerStatus(error.message);
      }
    });

    document.getElementById("query-btn").addEventListener("click", async () => {
      try {
        lastPayload = buildPayload();
        setQueryStatus("正在查询，请稍候...");
        const result = await postJson("/api/query", lastPayload);
        renderSummary(result.summary);
        renderRecords(result.summary.records);
        document.getElementById("export-btn").disabled = result.summary.records.length === 0;
        setQueryStatus(`查询完成：启用服务器 ${result.summary.enabled_servers} 台，遍历数据库 ${result.summary.visited_databases} 个，匹配业务库 ${result.summary.matched_databases} 个。`);
      } catch (error) {
        document.getElementById("export-btn").disabled = true;
        setQueryStatus(error.message);
      }
    });

    document.getElementById("export-btn").addEventListener("click", async () => {
      if (!lastPayload) return;
      try {
        const resp = await fetch("/api/export", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify(lastPayload)
        });
        if (!resp.ok) {
          const data = await resp.json().catch(() => ({}));
          throw new Error(data.detail || "导出失败");
        }
        const blob = await resp.blob();
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "项目阶段汇总.xlsx";
        link.click();
      } catch (error) {
        setQueryStatus(error.message);
      }
    });

    renderServers();
  </script>
</body>
</html>
"""


class ServerInput(BaseModel):
    name: str
    host: str
    port: int = 1433
    username: str
    password: str
    enabled: bool = True


class TestRequest(BaseModel):
    server: ServerInput


class QueryRequest(BaseModel):
    servers: list[ServerInput] = Field(default_factory=list)
    status_filter: str = "正在进行 + 即将开始"
    stage_keyword: str = ""
    project_keyword: str = ""


app = FastAPI(title="项目阶段汇总服务")


def _run_query(payload: QueryRequest):
    with tempfile.TemporaryDirectory(prefix="project_stage_web_") as temp_dir:
        temp_path = Path(temp_dir)
        input_path = temp_path / "query.json"
        output_path = temp_path / "result.json"
        dump_status_query_payload(
            servers=[StageServerConfig(**item.model_dump()) for item in payload.servers],
            status_filter=payload.status_filter,
            stage_keyword=payload.stage_keyword,
            project_keyword=payload.project_keyword,
            output_path=input_path,
        )
        runner_path = Path(__file__).resolve().with_name("project_stage_runner.py")
        result = subprocess.run(
            [sys.executable, str(runner_path), str(input_path), str(output_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            stderr_text = (result.stderr or "").strip()
            stdout_text = (result.stdout or "").strip()
            details = [f"返回码: {result.returncode}"]
            if stderr_text:
                details.append(f"stderr: {stderr_text}")
            if stdout_text:
                details.append(f"stdout: {stdout_text}")
            if not stderr_text and not stdout_text:
                details.append("未捕获到标准输出，请检查本机 Python/ODBC 环境。")
            raise RuntimeError("项目阶段汇总查询失败\n" + "\n".join(details))
        return summary_from_dict(json.loads(output_path.read_text(encoding="utf-8")))


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return HTML_PAGE


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/test")
def test_connection(payload: TestRequest) -> dict:
    try:
        summary = _run_query(QueryRequest(servers=[payload.server], status_filter="全部"))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "visited_databases": summary.visited_databases,
        "matched_databases": summary.matched_databases,
    }


@app.post("/api/query")
def query(payload: QueryRequest) -> dict:
    if not payload.servers:
        raise HTTPException(status_code=400, detail="请至少添加一台服务器。")
    try:
        summary = _run_query(payload)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"summary": summary_to_dict(summary)}


@app.post("/api/export")
def export(payload: QueryRequest):
    if not payload.servers:
        raise HTTPException(status_code=400, detail="请至少添加一台服务器。")
    try:
        summary = _run_query(payload)
        temp_dir = Path(tempfile.mkdtemp(prefix="project_stage_export_"))
        output_path = temp_dir / "项目阶段汇总.xlsx"
        export_project_stages(summary, output_path)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="项目阶段汇总.xlsx",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="项目阶段汇总服务")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    import uvicorn

    uvicorn.run(app, host=args.host, port=args.port)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
