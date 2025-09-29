from pathlib import Path
path = Path(r"c:\Users\EvSchneider\cm360-audit\Code.js")
text = path.read_text(encoding='utf-8')
old = "function runDailyAuditByName(configName) {\n if (!checkDriveApiEnabled()) return;\n const config = auditConfigs.find(c => c.name === configName);\n if (!config) {\n\tLogger.log(`? Config \"${configName}\" not found.`);\n return;\n }\n executeAudit(config);\n}\n\nfunction runAuditBatch(configs, isFinal = false) {"
if old not in text:
    raise SystemExit('Exact old block not found')
new = "function runDailyAuditByName(configName) {\n\tif (!checkDriveApiEnabled()) return;\n\tconst config = getAuditConfigByName(configName);\n\tif (!config) {\n\t\tLogger.log(`?? Config \"${configName}\" not found.`);\n\t\treturn;\n\t}\n\texecuteAudit(config);\n}\n\nfunction runAuditBatch(configs, isFinal = false) {"
text = text.replace(old, new, 1)
path.write_text(text, encoding='utf-8')
