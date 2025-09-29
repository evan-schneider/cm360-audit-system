from pathlib import Path
lines = Path(r"c:\Users\EvSchneider\cm360-audit\Code.js").read_text(encoding="utf-8").splitlines()
for i, line in enumerate(lines):
    if line.strip().startswith('auditConfigs.forEach('):
        # find end of duplicate function (look for blank line or 'function')
        j = i
        while j < len(lines) and not lines[j].startswith('function '):
            j += 1
        del lines[i:j]
        break
Path(r"c:\Users\EvSchneider\cm360-audit\Code.js").write_text('\n'.join(lines) + '\n', encoding="utf-8")
