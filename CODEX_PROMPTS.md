# 🧠 CODEX PROMPTS — DOCX TABLE ENGINE

> Use tasks sequentially. Do NOT skip stages.
> Always verify output before moving to next stage.

---

# 🔧 GLOBAL RULES

- NEVER use doc.tables directly
- ALWAYS use TableManager abstraction
- Support nested tables via XPath
- Preserve backward compatibility
- Do not break existing mapping logic
- Log all structural changes

---

# 🥇 STAGE 1 — TableManager Core

## 🎯 Goal
Create a unified table access layer supporting nested tables.

## 📂 Target file
table_manager.py

## 📌 Tasks
- Implement TableManager class
- Add:
  - flat_tables (XPath)
  - top_level_tables
  - get_table(index)
  - print_table_map()
- Add signature system:
  - shape (rows x cols)
  - preview (first 50 chars)
  - hash (md5)

## ✅ Done when
- Can print all tables including nested
- get_table(i) works without errors
- No existing files modified

## ▶️ Codex prompt
Implement table_manager.py as described above. Do not modify other files.

---

# 🥈 STAGE 2 — Data Filler Refactor

## 🎯 Goal
Replace doc.tables usage with TableManager safely.

## 📂 Target file
data_filler_v2.py

## 📌 Tasks
- Replace all doc.tables calls
- Use:
  tm = TableManager(doc, include_nested=True)
- Replace:
  doc.tables[i] → tm.get_table(i)
- Add fallback:
  - if table not found → match by signature
- Add debug mode:
  - print_table_map()

## ✅ Done when
- Script produces identical output as before
- No missing tables
- Debug prints correct structure

## ▶️ Codex prompt
Refactor data_filler_v2.py to use TableManager instead of doc.tables. Preserve existing behavior and add fallback via signature.

---

# 🥉 STAGE 3 — Comparison Engine

## 🎯 Goal
Eliminate false diffs caused by index mismatch.

## 📂 Target files
- compare_documents.py
- analyze_tables.py

## 📌 Tasks
- Use TableManager in both documents
- Compare tables:
  1. by index
  2. fallback by signature
- Improve compare_tables():
  - normalize text (strip)
- Ensure nested tables included

## ✅ Done when
- No false diffs for identical docs
- Nested tables are compared
- Output diff is stable

## ▶️ Codex prompt
Update comparison logic to use TableManager and signature-based matching. Ensure stable and accurate diff results.

---

# 🏆 STAGE 4 — Enterprise Features

## 🎯 Goal
Make system self-healing and traceable.

## 📂 New files
- table_mapper.py
- change_logger.py
- visual_diff.py

## 📌 Tasks

### 1. Table Mapper
- Match tables between documents
- Strategy:
  - exact hash match
  - fallback: shape + preview

### 2. Change Logger
- Log:
  - table index
  - row, col
  - old value → new value
  - timestamp
- Save to JSON

### 3. Visual Diff
- Human-readable output:
  [row,col] "old" → "new"

## ✅ Done when
- Tables auto-remap correctly
- JSON log is generated
- Diff is readable

## ▶️ Codex prompt
Implement enterprise features: table_mapper, change_logger, and visual_diff. Integrate with existing pipeline.

---

# 🧪 STAGE 5 — Verification

## 🎯 Goal
Ensure system correctness vs baseline document.

## 📌 Tasks
- Run generation
- Compare with reference doc
- Ensure:
  - no diffs
  - identical structure

## ✅ Done when
- Diff result is empty
- Hashes match

## ▶️ Codex prompt
Run comparison between generated document and reference. Report differences if any.

---

# 🚀 STAGE 6 — Stress Test

## 🎯 Goal
Validate robustness.

## 📌 Tasks
- Add nested table inside cell
- Re-run pipeline
- Verify:
  - no crashes
  - correct mapping

## ▶️ Codex prompt
Modify document by adding nested tables and verify system stability.

---

# 💡 OPTIONAL — OPTIMIZATION

## 🎯 Goal
Improve performance

## 📌 Tasks
- Cache table signatures
- Avoid duplicate XPath scans
- Minimize repeated loops

## ▶️ Codex prompt
Optimize table processing for performance without changing behavior.

---

# ⚠️ FINAL RULE

At each stage:
1. Implement
2. Run
3. Verify
4. Commit

DO NOT PROCEED if errors exist.