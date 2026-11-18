# -*- coding: utf-8 -*-
"""Проверка сборщика после изменений: собирает все примеры/*/report.md и печатает
итог. Запускать на Windows с Word (Проверка.bat). Ожидается «Замечаний нет»."""
import glob, os, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
ok = True
for src in sorted(glob.glob(os.path.join(HERE, "примеры", "*", "report.md"))):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    t = time.time()
    r = subprocess.run([sys.executable, os.path.join(HERE, "build.py"), src, "--pdf"],
                       capture_output=True, env=env)
    out = (r.stdout + r.stderr).decode("utf-8", "replace")
    with open(os.path.join(os.path.dirname(src), "_log.txt"), "w", encoding="utf-8") as f:
        f.write(out)
    clean = "Замечаний нет" in out
    ok &= clean
    name = os.path.basename(os.path.dirname(src))
    page = next((l for l in out.splitlines() if l.startswith("Страниц:")), "")
    print(f"{name}: {'OK' if clean else 'ЕСТЬ ЗАМЕЧАНИЯ'} ({time.time() - t:.0f} с) {page}")
    if not clean:
        print("   подробности: " + os.path.join(os.path.dirname(src), "_log.txt"))
print("\nВсё чисто." if ok else "\nЕсть замечания — смотрите _log.txt в папках примеров.")
