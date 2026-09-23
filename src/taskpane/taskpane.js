/* global Office, Excel */

import JSZip from "jszip";
import * as XLSX from "xlsx";
import { getOpenWorkbookBytes, extractSheetAsWorkbookBuffer } from "./ooxmlExport";

Office.onReady(() => {
  document.getElementById("openExportToolBtn").onclick = () => showView("exportView");
  document.getElementById("openLockToolBtn").onclick = () => showView("lockView");
  document.getElementById("openMergeDataBtn").onclick = () => showView("mergeDataView");
  document.getElementById("openMergeWorkbooksBtn").onclick = () => showView("mergeWorkbooksView");

  document.getElementById("backFromExport").onclick = () => showView("homeView");
  document.getElementById("backFromLock").onclick = () => showView("homeView");
  document.getElementById("backFromMergeData").onclick = () => showView("homeView");
  document.getElementById("backFromMergeWorkbooks").onclick = () => showView("homeView");

  document.getElementById("loadSheetsBtn").onclick = () => loadSheets("sheetList", "selectAll");
  document.getElementById("selectAll").onchange = () => toggleSelectAll("sheetList", "selectAll");
  document.getElementById("exportSheetsBtn").onclick = exportSelectedSheets;

  document.getElementById("loadSheetsLockBtn").onclick = () =>
    loadSheets("sheetListLock", "selectAllLock");
  document.getElementById("selectAllLock").onchange = () =>
    toggleSelectAll("sheetListLock", "selectAllLock");
  document.getElementById("lockSheetsBtn").onclick = () => lockUnlockSheets(true);
  document.getElementById("unlockSheetsBtn").onclick = () => lockUnlockSheets(false);

  document.getElementById("loadSheetsMergeBtn").onclick = () =>
    loadSheets("sheetListMerge", "selectAllMerge");
  document.getElementById("selectAllMerge").onchange = () =>
    toggleSelectAll("sheetListMerge", "selectAllMerge");
  document.getElementById("mergeSheetsDataBtn").onclick = mergeSheetsData;

  document.getElementById("mergeWorkbooksBtn").onclick = mergeWorkbooks;
});

function showView(viewId) {
  document.querySelectorAll(".view").forEach((v) => (v.style.display = "none"));
  document.getElementById(viewId).style.display = "block";
}

async function loadSheets(listId, selectAllId) {
  const listDiv = document.getElementById(listId);
  listDiv.innerHTML = "<p style='color:#6366f1'>Loading...</p>";
  document.getElementById(selectAllId).checked = false;

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      listDiv.innerHTML = "";

      if (sheets.items.length === 0) {
        listDiv.innerHTML = "<p style='color:#94a3b8'>No sheets found.</p>";
        return;
      }

      sheets.items.forEach((sheet) => {
        const item = document.createElement("div");
        item.className = "sheet-item";
        item.innerHTML = `<label>
          <input type="checkbox" class="sheet-cb-${listId}" value="${escapeHtml(sheet.name)}">
          ${escapeHtml(sheet.name)}
        </label>`;
        listDiv.appendChild(item);
      });
    });
  } catch (e) {
    listDiv.innerHTML = `<p style='color:red'>Error loading sheets: ${e.message || e}</p>`;
  }
}

function toggleSelectAll(listId, selectAllId) {
  const checked = document.getElementById(selectAllId).checked;
  document.querySelectorAll(`.sheet-cb-${listId}`).forEach((cb) => (cb.checked = checked));
}

function getSelectedSheets(listId) {
  return Array.from(document.querySelectorAll(`.sheet-cb-${listId}:checked`)).map((cb) => cb.value);
}

// 1. EXPORT SHEETS
// Exports each selected sheet as a standalone .xlsx by pruning the real OOXML
// package of the currently open workbook (see ooxmlExport.js), instead of
// rebuilding a new sheet from Office.js range values.
async function exportSelectedSheets() {
  const selected = getSelectedSheets("sheetList");
  const result = document.getElementById("exportResult");
  const folderName = document.getElementById("projectName").value.trim() || "JSGS Export";

  if (selected.length === 0) {
    result.innerHTML = "<p style='color:red'>Please select at least one sheet.</p>";
    return;
  }

  result.innerHTML = "<p style='color:#6366f1'>Exporting...</p>";

  try {
    const sourceBytes = await getOpenWorkbookBytes();
    const sourceZip = await JSZip.loadAsync(sourceBytes);
    const zip = new JSZip();
    const warnings = [];

    for (const sheetName of selected) {
      try {
        const { buffer, droppedDefinedNames } = await extractSheetAsWorkbookBuffer(
          sourceZip,
          sheetName,
          JSZip
        );
        zip.file(cleanFileName(sheetName) + ".xlsx", buffer);
        if (droppedDefinedNames > 0) {
          warnings.push(
            `"${sheetName}": تم حذف ${droppedDefinedNames} اسم معرّف (defined name) لأنها قد تشير إلى شيتات أخرى غير موجودة في الملف المصدَّر.`
          );
        }
      } catch (sheetError) {
        warnings.push(`"${sheetName}": فشل التصدير — ${sheetError.message || sheetError}`);
      }
    }

    const blob = await zip.generateAsync({ type: "blob" });
    triggerDownload(blob, cleanFileName(folderName) + ".zip");

    const warningHtml = warnings.length
      ? `<p style='color:#b45309;font-size:12px'>${warnings.map(escapeHtml).join("<br>")}</p>`
      : "";
    result.innerHTML = `<p style='color:green'><b>Done!</b> ${selected.length} sheet(s) exported.</p>${warningHtml}`;
  } catch (e) {
    result.innerHTML = `<p style='color:red'>Error: ${e.message || e}</p>`;
  }
}

// 2. LOCK / UNLOCK SHEETS
async function lockUnlockSheets(lock) {
  const selected = getSelectedSheets("sheetListLock");
  const password = document.getElementById("lockPassword").value;
  const result = document.getElementById("lockResult");
  const action = lock ? "Locking" : "Unlocking";

  if (selected.length === 0) {
    result.innerHTML = "<p style='color:red'>Please select at least one sheet.</p>";
    return;
  }

  result.innerHTML = `<p style='color:#6366f1'>${action}...</p>`;

  try {
    await Excel.run(async (context) => {
      for (const sheetName of selected) {
        const sheet = context.workbook.worksheets.getItem(sheetName);

        if (lock) {
          sheet.protection.protect(
            {
              allowFormatCells: false,
              allowFormatColumns: false,
              allowFormatRows: false,
              allowInsertColumns: false,
              allowInsertRows: false,
              allowInsertHyperlinks: false,
              allowDeleteColumns: false,
              allowDeleteRows: false,
              allowSort: false,
              allowAutoFilter: false,
              allowPivotTables: false,
            },
            password || undefined
          );
        } else {
          sheet.protection.unprotect(password || undefined);
        }
      }

      await context.sync();
    });

    const label = lock ? "🔒 Locked" : "🔓 Unlocked";
    result.innerHTML = `<p style='color:green'><b>${label}</b> ${selected.length} sheet(s) successfully.</p>`;
  } catch (e) {
    result.innerHTML = `<p style='color:red'>Error: ${e.message || e}</p>`;
  }
}

// 3. MERGE SHEETS DATA — each source sheet keeps its own formatting
// Every selected sheet is copied as its own independent block via Excel's
// native Range.copyFrom(..., Excel.RangeCopyType.all), which carries values,
// formulas AND formatting together straight from the source range — no sheet
// is ever used as a formatting template for another sheet's data.
async function mergeSheetsData() {
  const selected = getSelectedSheets("sheetListMerge");
  const destName = (document.getElementById("mergeDestName").value.trim() || "Merged Data").slice(
    0,
    31
  );
  const skipHeaders = document.getElementById("skipHeadersMerge").checked;
  const result = document.getElementById("mergeDataResult");

  if (selected.length === 0) {
    result.innerHTML = "<p style='color:red'>Please select at least one sheet.</p>";
    return;
  }

  result.innerHTML = "<p style='color:#6366f1'>Merging...</p>";

  let sourcesWithData = 0;
  let mergedAreaCount = 0;

  try {
    await Excel.run(async (context) => {
      let destSheet;
      try {
        destSheet = context.workbook.worksheets.getItem(destName);
        const existingRange = destSheet.getUsedRangeOrNullObject();
        existingRange.load("isNullObject");
        await context.sync();
        if (!existingRange.isNullObject) existingRange.clear();
      } catch {
        destSheet = context.workbook.worksheets.add(destName);
      }
      await context.sync();

      let destRowCursor = 0;

      for (let i = 0; i < selected.length; i++) {
        const sourceSheet = context.workbook.worksheets.getItem(selected[i]);
        const usedRange = sourceSheet.getUsedRangeOrNullObject();
        usedRange.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync();

        if (usedRange.isNullObject || usedRange.rowCount === 0 || usedRange.columnCount === 0)
          continue;

        const skipRow = skipHeaders && sourcesWithData > 0 && usedRange.rowCount > 1;
        const srcTopRow = usedRange.rowIndex + (skipRow ? 1 : 0);
        const srcRowCount = usedRange.rowCount - (skipRow ? 1 : 0);
        if (srcRowCount <= 0) continue;

        const srcRange = sourceSheet.getRangeByIndexes(
          srcTopRow,
          usedRange.columnIndex,
          srcRowCount,
          usedRange.columnCount
        );
        const destRange = destSheet.getRangeByIndexes(
          destRowCursor,
          0,
          srcRowCount,
          usedRange.columnCount
        );
        destRange.copyFrom(srcRange, Excel.RangeCopyType.all, false, false);

        // Recreate merged cells that fall inside the copied block (formats
        // copy does not carry merge state). Skipped safely on failure —
        // never left half-applied.
        try {
          const mergedAreas = srcRange.getMergedAreasOrNullObject();
          mergedAreas.load(
            "isNullObject,areas/items/rowIndex,areas/items/columnIndex,areas/items/rowCount,areas/items/columnCount"
          );
          await context.sync();

          if (!mergedAreas.isNullObject) {
            mergedAreas.areas.items.forEach((area) => {
              const relRow = area.rowIndex - srcTopRow;
              const relCol = area.columnIndex - usedRange.columnIndex;
              destSheet
                .getRangeByIndexes(destRowCursor + relRow, relCol, area.rowCount, area.columnCount)
                .merge(false);
              mergedAreaCount++;
            });
            await context.sync();
          }
        } catch {
          // Merged-area recreation isn't safely supported here (older Excel
          // host) — the copied values/formatting above are unaffected.
        }

        destRowCursor += srcRowCount;
        sourcesWithData++;
      }

      if (sourcesWithData > 0) {
        destSheet.activate();
      }
      await context.sync();
    });

    if (sourcesWithData === 0) {
      result.innerHTML = "<p style='color:red'>No data found in selected sheets.</p>";
      return;
    }

    const mergeNote =
      mergedAreaCount > 0 ? ` — ${mergedAreaCount} merged cell area(s) recreated.` : "";
    result.innerHTML = `<p style='color:green'><b>Done!</b> Data merged into "<b>${escapeHtml(destName)}</b>" — each sheet kept its own formatting.${mergeNote}</p>`;
  } catch (e) {
    result.innerHTML = `<p style='color:red'>Error: ${e.message || e}</p>`;
  }
}

// 4. MERGE WORKBOOKS AS SHEETS WITH FORMATTING
async function mergeWorkbooks() {
  const filesInput = document.getElementById("workbookFiles");
  const outputName =
    document.getElementById("mergedWorkbookName").value.trim() || "Merged Workbook";
  const result = document.getElementById("mergeWorkbooksResult");

  if (!filesInput.files || filesInput.files.length === 0) {
    result.innerHTML = "<p style='color:red'>Please select at least one Excel file.</p>";
    return;
  }

  result.innerHTML = "<p style='color:#6366f1'>Merging workbooks with formatting...</p>";

  try {
    const mergedWb = XLSX.utils.book_new();
    const usedNames = {};

    for (const file of filesInput.files) {
      const buffer = await file.arrayBuffer();

      const wb = XLSX.read(buffer, {
        type: "array",
        cellStyles: true,
        cellNF: true,
        cellDates: true,
        cellFormula: true,
        sheetStubs: true,
        bookVBA: true,
      });

      if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) continue;

      wb.SheetNames.forEach((sheetName) => {
        const originalWs = wb.Sheets[sheetName];
        if (!originalWs) return;

        const ws = cloneWorksheet(originalWs);

        const baseFromFile = cleanFileName(file.name.replace(/\.[^/.]+$/, ""));
        const proposedName = `${baseFromFile}_${sheetName}`;
        const finalName = makeUniqueSheetName(proposedName, usedNames);

        XLSX.utils.book_append_sheet(mergedWb, ws, finalName);
      });
    }

    if (mergedWb.SheetNames.length === 0) {
      result.innerHTML = "<p style='color:red'>No sheets found in selected files.</p>";
      return;
    }

    const buf = XLSX.write(mergedWb, {
      type: "array",
      bookType: "xlsx",
      cellStyles: true,
      compression: true,
      bookSST: true,
    });

    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    triggerDownload(blob, cleanFileName(outputName) + ".xlsx");

    result.innerHTML = `<p style='color:green'><b>Done!</b> ${filesInput.files.length} file(s) merged with formatting into <b>${escapeHtml(outputName)}.xlsx</b></p>`;
  } catch (e) {
    result.innerHTML = `<p style='color:red'>Error: ${e.message || e}</p>`;
  }
}

// HELPERS
function cloneWorksheet(ws) {
  const cloned = {};

  Object.keys(ws).forEach((key) => {
    const value = ws[key];

    if (value && typeof value === "object") {
      try {
        cloned[key] = JSON.parse(JSON.stringify(value));
      } catch {
        cloned[key] = value;
      }
    } else {
      cloned[key] = value;
    }
  });

  if (ws["!cols"]) cloned["!cols"] = JSON.parse(JSON.stringify(ws["!cols"]));
  if (ws["!rows"]) cloned["!rows"] = JSON.parse(JSON.stringify(ws["!rows"]));
  if (ws["!merges"]) cloned["!merges"] = JSON.parse(JSON.stringify(ws["!merges"]));
  if (ws["!autofilter"]) cloned["!autofilter"] = JSON.parse(JSON.stringify(ws["!autofilter"]));
  if (ws["!margins"]) cloned["!margins"] = JSON.parse(JSON.stringify(ws["!margins"]));
  if (ws["!protect"]) cloned["!protect"] = JSON.parse(JSON.stringify(ws["!protect"]));

  return cloned;
}

function makeUniqueSheetName(name, usedNames) {
  let clean = safeSheetName(name);
  let finalName = clean;

  if (usedNames[finalName] === undefined) {
    usedNames[finalName] = 0;
    return finalName;
  }

  usedNames[finalName]++;

  const suffix = "_" + usedNames[finalName];
  finalName = clean.slice(0, 31 - suffix.length) + suffix;

  while (usedNames[finalName] !== undefined) {
    usedNames[clean]++;
    const newSuffix = "_" + usedNames[clean];
    finalName = clean.slice(0, 31 - newSuffix.length) + newSuffix;
  }

  usedNames[finalName] = 0;
  return finalName;
}

function safeSheetName(name) {
  return (
    cleanFileName(name)
      .replace(/[\[\]]/g, "_")
      .slice(0, 31) || "Sheet"
  );
}

function cleanFileName(name) {
  return (
    String(name || "File")
      .replace(/[\\/:*?"<>|]/g, "_")
      .trim() || "File"
  );
}

function triggerDownload(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
