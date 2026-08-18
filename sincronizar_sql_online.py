from __future__ import annotations

import argparse
import base64
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
GH = BASE_DIR / ".tools" / "bin" / "gh.exe"
REPO = "JoaoCanhadas/acompanhamento-diario"
BRANCH = "main"
FILES_TO_SYNC = [
    "data.json",
    "geral.json",
    "keys.json",
    "positivacao_milho.json",
]


def load_env_file():
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        os.environ[name.strip()] = value.strip()


def gh_api(*args, stdin_data=None):
    env = os.environ.copy()

    env.pop("GH_TOKEN", None)
    env.pop("GITHUB_TOKEN", None)

    result = subprocess.run(
        [str(GH), "api", *args],
        capture_output=True,
        text=True,
        input=stdin_data,
        env=env,
    )

    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout)

    return result.stdout


def gh_json(*args, stdin_data=None):
    return json.loads(gh_api(*args, stdin_data=stdin_data))


def generate_files():
    result = subprocess.run([sys.executable, str(BASE_DIR / "exportar_dados.py")])
    if result.returncode != 0:
        raise RuntimeError("Falha ao gerar os dados pelo SQL.")


def publish_json_files():
    changed = [filename for filename in FILES_TO_SYNC if remote_content(filename) != (BASE_DIR / filename).read_bytes()]
    if not changed:
        return None

    ref = gh_json(f"repos/{REPO}/git/ref/heads/{BRANCH}")
    head_sha = ref["object"]["sha"]
    head_commit = gh_json(f"repos/{REPO}/git/commits/{head_sha}")
    base_tree = head_commit["tree"]["sha"]

    tree_items = []
    for filename in changed:
        local_path = BASE_DIR / filename
        content = base64.b64encode(local_path.read_bytes()).decode("ascii")
        blob = gh_json(
            f"repos/{REPO}/git/blobs",
            "--method",
            "POST",
            "--input",
            "-",
            stdin_data=json.dumps({"content": content, "encoding": "base64"}),
        )
        tree_items.append(
            {
                "path": filename,
                "mode": "100644",
                "type": "blob",
                "sha": blob["sha"],
            }
        )

    tree = gh_json(
        f"repos/{REPO}/git/trees",
        "--method",
        "POST",
        "--input",
        "-",
        stdin_data=json.dumps({"base_tree": base_tree, "tree": tree_items}),
    )
    message = f"Dados SQL sincronizados em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} [skip render]"
    commit = gh_json(
        f"repos/{REPO}/git/commits",
        "--method",
        "POST",
        "--input",
        "-",
        stdin_data=json.dumps(
            {
                "message": message,
                "tree": tree["sha"],
                "parents": [head_sha],
            }
        ),
    )
    gh_api(
        f"repos/{REPO}/git/refs/heads/{BRANCH}",
        "--method",
        "PATCH",
        "--input",
        "-",
        stdin_data=json.dumps({"sha": commit["sha"]}),
    )
    return commit["sha"]


def remote_content(filename):
    try:
        item = gh_json(f"repos/{REPO}/contents/{filename}")
        return base64.b64decode(item["content"])
    except Exception:
        return None


def sync_once():
    generate_files()
    commit_sha = publish_json_files()
    if commit_sha:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Online atualizado: {commit_sha[:7]}", flush=True)
    else:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Sem alteracoes nos dados.", flush=True)


def main():
    load_env_file()
    parser = argparse.ArgumentParser(description="Sincroniza SQL local com o dashboard online.")
    parser.add_argument("--once", action="store_true", help="Executa uma sincronizacao e sai.")
    parser.add_argument("--interval", type=int, default=60, help="Intervalo em segundos.")
    args = parser.parse_args()

    if args.once:
        sync_once()
        return 0

    print("Sincronizacao SQL -> dashboard online iniciada.", flush=True)
    print(f"Intervalo: {args.interval} segundos", flush=True)
    print("Pressione Ctrl+C para parar.", flush=True)
    while True:
        try:
            sync_once()
        except Exception as exc:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO: {exc}", flush=True)
        time.sleep(max(args.interval, 15))


if __name__ == "__main__":
    sys.exit(main())
