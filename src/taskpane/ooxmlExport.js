/* global Office */

// Exports a single worksheet as a standalone .xlsx by pruning the *real* OOXML
// package of the currently open workbook, instead of rebuilding a new sheet
// from Office.js range values (which drops fonts, fills, borders, merges,
// conditional formatting, tables, validation, images, charts, print setup...).
//
// Strategy: copy every package part the target sheet actually depends on
// (its own XML, drawings/charts/images/tables/comments it references,
// styles.xml, sharedStrings.xml, theme) byte-for-byte into a new package,
// and only rewrite the small "index" parts (workbook.xml, workbook.xml.rels,
// [Content_Types].xml, docProps/app.xml) that have to point at just one sheet.
// Because styles.xml and sharedStrings.xml are copied untouched, every style
// index (s="N") and shared-string index inside the sheet XML stays valid
// without any remapping.

const NS = {
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  pkgRel: "http://schemas.openxmlformats.org/package/2006/relationships",
};

export async function getOpenWorkbookBytes() {
  if (
    !(
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.document &&
      Office.context.document.getFileAsync
    )
  ) {
    throw new Error("Office.context.document.getFileAsync غير مدعوم في هذا السياق.");
  }

  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4194304 },
      (fileResult) => {
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(
            new Error((fileResult.error && fileResult.error.message) || "تعذّر قراءة الملف الحالي.")
          );
          return;
        }

        const file = fileResult.value;
        const sliceCount = file.sliceCount;

        if (sliceCount === 0) {
          file.closeAsync();
          resolve(new Uint8Array(0));
          return;
        }

        const slices = new Array(sliceCount);
        let received = 0;

        const getSlice = (index) => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(
                new Error(
                  (sliceResult.error && sliceResult.error.message) || "فشل قراءة جزء من الملف."
                )
              );
              return;
            }

            slices[index] = sliceResult.value.data;
            received++;

            if (received === sliceCount) {
              file.closeAsync();
              let totalLength = 0;
              for (const s of slices) totalLength += s.length;
              const merged = new Uint8Array(totalLength);
              let offset = 0;
              for (const s of slices) {
                merged.set(s, offset);
                offset += s.length;
              }
              resolve(merged);
            } else {
              getSlice(index + 1);
            }
          });
        };

        getSlice(0);
      }
    );
  });
}

function parseXml(str) {
  const doc = new DOMParser().parseFromString(str, "application/xml");
  const errorNode = doc.getElementsByTagName("parsererror")[0];
  if (errorNode) {
    throw new Error("ملف XML داخل الحزمة تالف أو غير صالح.");
  }
  return doc;
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function resolvePackagePath(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const segments = (baseDir ? baseDir.split("/") : []).concat(target.split("/"));
  const stack = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") stack.pop();
    else stack.push(seg);
  }
  return stack.join("/");
}

function dirOf(path) {
  return path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "";
}

function baseOf(path) {
  return path.includes("/") ? path.slice(path.lastIndexOf("/") + 1) : path;
}

async function readXml(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error(`الجزء "${path}" غير موجود داخل ملف .xlsx.`);
  return parseXml(await entry.async("string"));
}

// Recursively walks a part's own _rels file and pulls in every internal
// target (drawings -> charts -> images, tables, comments, printer settings...)
// so nothing the sheet depends on is left out of the pruned package.
async function collectDependencies(zip, partPath, keepSet) {
  if (keepSet.has(partPath)) return;
  const entry = zip.file(partPath);
  if (!entry) return; // referenced part missing from source file — skip defensively

  keepSet.add(partPath);

  const dir = dirOf(partPath);
  const relsPath = (dir ? dir + "/" : "") + "_rels/" + baseOf(partPath) + ".rels";
  const relsEntry = zip.file(relsPath);
  if (!relsEntry) return;

  keepSet.add(relsPath);
  const relsDoc = parseXml(await relsEntry.async("string"));
  const rels = Array.from(relsDoc.getElementsByTagNameNS(NS.pkgRel, "Relationship"));

  for (const rel of rels) {
    if (rel.getAttribute("TargetMode") === "External") continue;
    const target = resolvePackagePath(dir, rel.getAttribute("Target"));
    await collectDependencies(zip, target, keepSet);
  }
}

function escapeXmlText(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildAppXml(sheetName) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
<vt:variant><vt:i4>1</vt:i4></vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="1" baseType="lpstr">
<vt:lpstr>${escapeXmlText(sheetName)}</vt:lpstr>
</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>`;
}

/**
 * Prunes the source .xlsx package down to a single sheet, keeping every part
 * that sheet depends on byte-for-byte. Returns the new file as an array
 * buffer plus a note of anything that could not be safely carried over.
 */
export async function extractSheetAsWorkbookBuffer(sourceZip, sheetName, JSZip) {
  if (!sourceZip.file("xl/workbook.xml")) {
    throw new Error("الملف المفتوح ليس بصيغة .xlsx صالحة (لا يحتوي على xl/workbook.xml).");
  }

  const wbXml = await readXml(sourceZip, "xl/workbook.xml");
  const wbRelsXml = await readXml(sourceZip, "xl/_rels/workbook.xml.rels");

  const sheetEl = Array.from(wbXml.getElementsByTagNameNS(NS.main, "sheet")).find(
    (el) => el.getAttribute("name") === sheetName
  );
  if (!sheetEl) throw new Error(`لم يتم العثور على الشيت "${sheetName}" في workbook.xml.`);

  const rId = sheetEl.getAttributeNS(NS.r, "id");
  const relEls = Array.from(wbRelsXml.getElementsByTagNameNS(NS.pkgRel, "Relationship"));
  const sheetRel = relEls.find((el) => el.getAttribute("Id") === rId);
  if (!sheetRel) throw new Error(`تعذّر تحديد مسار ملف الشيت "${sheetName}" داخل الحزمة.`);

  const sheetPath = resolvePackagePath("xl", sheetRel.getAttribute("Target"));

  const keepSet = new Set();
  await collectDependencies(sourceZip, sheetPath, keepSet);

  const stylesRel = relEls.find((el) => (el.getAttribute("Type") || "").endsWith("/styles"));
  const sstRel = relEls.find((el) => (el.getAttribute("Type") || "").endsWith("/sharedStrings"));
  const themeRel = relEls.find((el) => (el.getAttribute("Type") || "").endsWith("/theme"));

  if (stylesRel)
    await collectDependencies(
      sourceZip,
      resolvePackagePath("xl", stylesRel.getAttribute("Target")),
      keepSet
    );
  if (sstRel)
    await collectDependencies(
      sourceZip,
      resolvePackagePath("xl", sstRel.getAttribute("Target")),
      keepSet
    );
  if (themeRel)
    await collectDependencies(
      sourceZip,
      resolvePackagePath("xl", themeRel.getAttribute("Target")),
      keepSet
    );

  keepSet.add("docProps/core.xml");
  keepSet.add("_rels/.rels");

  // --- rewrite xl/workbook.xml: keep only this one <sheet> ---
  const sheetsEl = wbXml.getElementsByTagNameNS(NS.main, "sheets")[0];
  Array.from(sheetsEl.children).forEach((ch) => sheetsEl.removeChild(ch));
  sheetsEl.appendChild(sheetEl);

  let droppedDefinedNames = 0;
  const definedNamesEl = wbXml.getElementsByTagNameNS(NS.main, "definedNames")[0];
  if (definedNamesEl) {
    droppedDefinedNames = definedNamesEl.getElementsByTagNameNS(NS.main, "definedName").length;
    definedNamesEl.parentNode.removeChild(definedNamesEl);
  }

  const workbookViewEl = wbXml.getElementsByTagNameNS(NS.main, "workbookView")[0];
  if (workbookViewEl) workbookViewEl.setAttribute("activeTab", "0");

  // --- rewrite xl/_rels/workbook.xml.rels: keep only sheet/styles/sst/theme ---
  const keptRelIds = new Set([rId]);
  if (stylesRel) keptRelIds.add(stylesRel.getAttribute("Id"));
  if (sstRel) keptRelIds.add(sstRel.getAttribute("Id"));
  if (themeRel) keptRelIds.add(themeRel.getAttribute("Id"));
  relEls.forEach((el) => {
    if (!keptRelIds.has(el.getAttribute("Id"))) el.parentNode.removeChild(el);
  });

  // --- rebuild [Content_Types].xml: keep Defaults, filter Overrides to kept parts ---
  const ctXml = await readXml(sourceZip, "[Content_Types].xml");
  const overrideEls = Array.from(ctXml.getElementsByTagNameNS(NS.ct, "Override"));
  let hasAppOverride = false;
  overrideEls.forEach((el) => {
    const partName = (el.getAttribute("PartName") || "").replace(/^\//, "");
    if (partName === "docProps/app.xml") hasAppOverride = true;
    if (partName !== "xl/workbook.xml" && !keepSet.has(partName)) {
      el.parentNode.removeChild(el);
    }
  });
  if (!hasAppOverride) {
    const ov = ctXml.createElementNS(NS.ct, "Override");
    ov.setAttribute("PartName", "/docProps/app.xml");
    ov.setAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    );
    ctXml.documentElement.appendChild(ov);
  }

  // --- assemble the pruned package ---
  const outZip = new JSZip();
  outZip.file("[Content_Types].xml", serializeXml(ctXml));
  outZip.file("xl/workbook.xml", serializeXml(wbXml));
  outZip.file("xl/_rels/workbook.xml.rels", serializeXml(wbRelsXml));
  outZip.file("docProps/app.xml", buildAppXml(sheetName));

  for (const path of keepSet) {
    const entry = sourceZip.file(path);
    if (!entry) continue; // dependency referenced but not present in source — already noted as skipped
    const bytes = await entry.async("uint8array");
    outZip.file(path, bytes);
  }

  const buffer = await outZip.generateAsync({ type: "array", compression: "DEFLATE" });

  return {
    buffer,
    droppedDefinedNames,
  };
}
