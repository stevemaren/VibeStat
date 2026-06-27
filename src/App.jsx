import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ComposedChart, Scatter, Line, LineChart, ErrorBar, ResponsiveContainer,
} from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { save as tauriSaveDialog, open as tauriOpenDialog } from "@tauri-apps/plugin-dialog";
import { readTextFile, writeTextFile, writeFile, readFile } from "@tauri-apps/plugin-fs";
import { Menu, Submenu, MenuItem, CheckMenuItem, PredefinedMenuItem } from "@tauri-apps/api/menu";
import { writeImage as tauriWriteImage } from "@tauri-apps/plugin-clipboard-manager";
import { Image as TauriImage } from "@tauri-apps/api/image";

/* =========================================================================
   DATASET I/O — JSON open/save + CSV/Excel import
   ========================================================================= */
// Build a VibeStat dataset object from a header row + array-of-arrays data rows.
function datasetFromTable(header, dataRows) {
  const cols = (header || []).map((h, i) => ({ id: "v" + i, name: String(h == null || h === "" ? "Col" + (i + 1) : h).trim(), type: "real" }));
  cols.forEach((col, i) => {
    let numeric = true, allInt = true, any = false;
    for (const r of dataRows) {
      const v = r[i];
      if (v == null || v === "") continue;
      any = true;
      const n = Number(v);
      if (!Number.isFinite(n)) { numeric = false; break; }
      if (!Number.isInteger(n)) allInt = false;
    }
    col.type = !any ? "string" : numeric ? (allInt ? "integer" : "real") : "string";
  });
  const rows = dataRows
    .filter((r) => r && r.some((v) => v != null && v !== ""))
    .map((r) => { const o = {}; cols.forEach((col, i) => { const v = r[i]; o[col.id] = v == null ? "" : String(v); }); return o; });
  return { app: "VibeStat", version: 4, columns: cols, rows, compacts: [], analyses: [], excluded: [], colW: {} };
}
function parseCSVText(text) {
  const p = Papa.parse(String(text).trim(), { skipEmptyLines: true });
  const aoa = p.data || [];
  return datasetFromTable(aoa[0] || [], aoa.slice(1));
}
function parseXLSXBytes(bytes) {
  const wb = XLSX.read(bytes, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
  return datasetFromTable(aoa[0] || [], aoa.slice(1));
}
/* ---- GraphPad Prism .pzfx (XML) importer --------------------------------
   Captures the data structure faithfully: group/category labels (Y-column &
   row titles), replicate subcolumns, X/Y layout, contingency counts, plus the
   Info-sheet notes & name/value constants. NOTE: a .pzfx file does NOT record
   which factors are repeated-measures (within-subjects) vs independent
   (between) — Prism decides that at analysis time — so values are imported in
   a layout the user can assign either way (tick columns + Compact for a
   repeated factor). All of this is summarized in a "Prism Import" notes panel. */
const PZFX_TT = {
  OneWay: "Column table (one grouping factor; groups in columns)",
  TwoWay: "Grouped table (two factors: rows \u00d7 columns; replicates in subcolumns)",
  XY: "XY table (X column + Y columns)",
  Contingency: "Contingency table (counts; rows \u00d7 columns)",
  Survival: "Survival table", PartsOfWhole: "Parts-of-whole table",
  Nested: "Nested table", MultipleVariables: "Multiple-variables table",
};
function _pzLocal(node) { const tn = (node && (node.tagName || node.nodeName)) || ""; return tn.indexOf(":") >= 0 ? tn.split(":").pop() : tn; }
function _pzText(node) { return node ? (node.textContent == null ? "" : String(node.textContent)) : ""; }
function _pzKids(node, name) { const out = []; if (!node) return out; for (let c = node.firstChild; c; c = c.nextSibling) { if (c.nodeType === 1 && _pzLocal(c) === name) out.push(c); } return out; }
function _pzFirst(node, name) { const k = _pzKids(node, name); return k.length ? k[0] : null; }
function _pzAttr(node, name) { return node && node.getAttribute ? (node.getAttribute(name) || "") : ""; }
function _pzTitle(node) { return _pzText(_pzFirst(node, "Title")).trim(); }
function _pzCell(d, strikeExclude) { if (!d) return ""; const ex = _pzAttr(d, "Excluded") || _pzAttr(d, "excluded"); if (strikeExclude && (ex === "1" || ex === "true" || ex === "yes")) return ""; return _pzText(d).trim(); }
function _pzNum(s) { if (s == null) return s; let t = String(s).trim(); if (t === "") return t; if (t.indexOf(",") >= 0 && t.indexOf(".") < 0) { const p = t.split(","); if (p.length === 2 && /^[+-]?\d*$/.test(p[0]) && /^\d+$/.test(p[1])) t = p[0] + "." + p[1]; } return t; }
function _pzInferNumType(values) { let any = false, allInt = true, numeric = true; for (const v of values) { if (v == null || v === "") continue; any = true; const n = Number(v); if (!Number.isFinite(n)) { numeric = false; break; } if (!Number.isInteger(n)) allInt = false; } return !any ? "real" : numeric ? (allInt ? "integer" : "real") : "string"; }
function _pzIsSDN(evf) { const e = (evf || "").toUpperCase().replace(/[^A-Z]/g, ""); return /SDN|SEN|SEM|MEANSD|MEANSEM|UPPERLOWER|CI/.test(e) || /SD$|SE$/.test(e); }
function _pzReadColumn(colNode, strikeExclude) { return { title: _pzTitle(colNode), subs: _pzKids(colNode, "Subcolumn").map((sc) => _pzKids(sc, "d").map((d) => _pzCell(d, strikeExclude))) }; }
// Map one Prism table node to flat column descriptors {name, kind:"cat"|"num", values:[]}.
function _pzMapTable(tbl, strikeExclude) {
  const ttype = (_pzAttr(tbl, "TableType") || "").trim();
  const evf = (_pzAttr(tbl, "EVFormat") || "").trim();
  const title = _pzTitle(tbl) || _pzAttr(tbl, "ID") || "Table";
  const cols = []; let groups = 0, reps = 0, hasRowTitles = false;
  const rt = _pzFirst(tbl, "RowTitlesColumn");
  if (rt) { const rc = _pzReadColumn(rt, strikeExclude); const vals = rc.subs[0] || []; if (vals.some((v) => v !== "")) { hasRowTitles = true; cols.push({ name: rc.title || "Row Title", kind: "cat", values: vals.slice() }); } }
  _pzKids(tbl, "XColumn").forEach((xc) => { const c = _pzReadColumn(xc, strikeExclude); cols.push({ name: c.title || "X", kind: "num", values: (c.subs[0] || []).map(_pzNum) }); });
  const yc = _pzKids(tbl, "YColumn"); groups = yc.length; const sdn = _pzIsSDN(evf);
  yc.forEach((y) => {
    const c = _pzReadColumn(y, strikeExclude); const n = c.subs.length; if (n > reps) reps = n;
    if (n <= 1) { cols.push({ name: c.title || "Y", kind: "num", values: (c.subs[0] || []).map(_pzNum) }); }
    else c.subs.forEach((sv, si) => { const suf = sdn ? ("_" + (["MEAN", "SD", "N"][si] || ("V" + (si + 1)))) : ("_" + (si + 1)); cols.push({ name: (c.title || "Y") + suf, kind: "num", values: sv.map(_pzNum) }); });
  });
  return { title, ttype, evf, cols, groups, reps, hasRowTitles };
}
function _pzUniqueNames(names) { const seen = {}, out = []; for (let nm of names) { nm = String(nm == null || nm === "" ? "Col" : nm).trim() || "Col"; if (seen[nm] == null) { seen[nm] = 1; out.push(nm); } else { seen[nm] += 1; out.push(nm + " (" + seen[nm] + ")"); } } return out; }
const PZFX_TDESC = (tt) => PZFX_TT[tt] || (tt ? ("\u201c" + tt + "\u201d table") : "table");
// Assemble columns/rows from a _pzMapTable result. Returns null if no data columns.
function _pzBuild(map) {
  if (!map.cols.length) return null;
  const names = _pzUniqueNames(map.cols.map((c) => c.name));
  const columns = map.cols.map((c, i) => ({ id: "v" + i, name: names[i], type: c.kind === "cat" ? "category" : _pzInferNumType(c.values) }));
  const maxLen = map.cols.reduce((m, c) => Math.max(m, c.values.length), 0);
  const rows = [];
  for (let r = 0; r < maxLen; r++) { const o = {}; columns.forEach((col, i) => { const v = map.cols[i].values[r]; o[col.id] = v == null ? "" : String(v); }); rows.push(o); }
  const rows2 = rows.filter((o) => columns.some((col) => o[col.id] !== "" && o[col.id] != null));
  if (!rows2.length) return null;
  const shape = `${map.groups} column${map.groups === 1 ? "" : "s"}${map.reps > 1 ? ", " + map.reps + " subcolumns" : ""}${map.hasRowTitles ? ", row titles" : ""}, ${rows2.length} rows`;
  return { columns, rows: rows2, shape };
}
// Detect a grouped/replicates (TwoWay) Prism table whose subcolumn replicates are subjects
// and whose rows are repeated measures, and read its raw group/replicate/row structure.
function _pzfxGroupedRM(tbl, strikeExclude) {
  if ((_pzAttr(tbl, "TableType") || "") !== "TwoWay") return null;
  const yc = _pzKids(tbl, "YColumn"); if (yc.length < 1) return null;
  const groups = yc.map((y) => ({ title: _pzTitle(y) || "Group", subs: _pzKids(y, "Subcolumn").map((sc) => _pzKids(sc, "d").map((d) => _pzCell(d, strikeExclude))) }));
  const rep = groups.reduce((m, g) => Math.max(m, g.subs.length), 0); if (rep < 1) return null;
  const rt = _pzFirst(tbl, "RowTitlesColumn");
  let rowTitles = rt ? ((_pzReadColumn(rt, strikeExclude).subs[0]) || []).slice() : [];
  const nRows = groups.reduce((m, g) => Math.max(m, g.subs.reduce((mm, sub) => Math.max(mm, sub.length), 0)), 0);
  if (rowTitles.length < nRows) for (let r = rowTitles.length; r < nRows; r++) rowTitles.push("Level" + (r + 1));
  if (nRows < 2 || rep < 1) return null; // need >1 row level for "repeated measures"
  const subTitles = {}; const sct = _pzFirst(tbl, "SubColumnTitles");
  if (sct) _pzKids(sct, "Subcolumn").forEach((sc, sIdx) => { _pzKids(sc, "d").map((d) => _pzText(d).trim()).forEach((nm, gIdx) => { (subTitles[gIdx] = subTitles[gIdx] || {})[sIdx] = nm; }); });
  return { groups, rep, rowTitles, subTitles, nRows };
}
// Transpose a grouped-RM structure to VibeStat cases x within-levels (+ compacted within variable).
function _pzfxBuildGroupedRM(grm) {
  const { groups, rep, rowTitles } = grm, subTitles = grm.subTitles, nGroups = groups.length;
  const fn = { colFactor: "Group", rowFactor: "Trial", matchedFactor: "Subject" };
  const columns = [{ id: "subj", name: fn.matchedFactor, type: "string" }, { id: "grp", name: fn.colFactor, type: "category" }];
  rowTitles.forEach((rt, k) => columns.push({ id: "m" + k, name: rt || ("Level" + (k + 1)), type: "real" }));
  const rows = [];
  for (let g = 0; g < nGroups; g++) for (let s = 0; s < rep; s++) {
    const sub = groups[g].subs[s] || [];
    const vals = rowTitles.map((_, k) => { const t = _pzNum(sub[k]); return t === "" || t == null ? "" : Number(t); });
    if (vals.every((v) => v === "")) continue;
    const row = { subj: (subTitles[g] && subTitles[g][s]) || (groups[g].title + "-" + (s + 1)), grp: groups[g].title };
    rowTitles.forEach((_, k) => (row["m" + k] = vals[k]));
    rows.push(row);
  }
  const compact = { id: "cpRM", name: "Measure", factors: [{ name: fn.rowFactor, levels: rowTitles.map((rt, k) => rt || ("Level" + (k + 1))) }], leaves: rowTitles.map((_, k) => "m" + k) };
  return { columns, rows, compact, shape: nGroups + " groups \u00d7 " + rep + " subjects \u00d7 " + grm.nRows + " levels", typeDesc: "Grouped, repeated measures (transposed to cases \u00d7 within levels)" };
}
// Parse a .pzfx into one VibeStat project per data table (for the table picker).
function parsePZFXTables(xmlText, strikeExclude) {
  if (strikeExclude == null) strikeExclude = true;
  if (typeof DOMParser === "undefined") throw new Error("XML parsing is unavailable in this environment.");
  const doc = new DOMParser().parseFromString(String(xmlText), "application/xml");
  const perr = doc.getElementsByTagName("parsererror");
  if (perr && perr.length) throw new Error("This file is not valid XML (.pzfx).");
  const root = doc.documentElement;
  if (!root || _pzLocal(root) !== "GraphPadPrismFile") throw new Error("Not a GraphPad Prism .pzfx file.");
  const infoSheets = _pzKids(root, "Info").map((inf) => ({
    title: _pzTitle(inf) || _pzAttr(inf, "ID") || "Info",
    notes: _pzText(_pzFirst(inf, "Notes")).replace(/\r/g, "").trim(),
    constants: _pzKids(inf, "Constant").map((cn) => ({ name: _pzText(_pzFirst(cn, "Name")).trim(), value: _pzText(_pzFirst(cn, "Value")).trim() })).filter((c) => c.name || c.value),
  }));
  const prismVer = _pzAttr(root, "PrismXMLVersion");
  let createdBy = ""; const created = _pzFirst(root, "Created"); const ov = created ? _pzFirst(created, "OriginalVersion") : null;
  if (ov) createdBy = (_pzAttr(ov, "CreatedByProgram") + " " + _pzAttr(ov, "CreatedByVersion")).trim();
  const tableNodes = []; for (let c = root.firstChild; c; c = c.nextSibling) { if (c.nodeType === 1) { const ln = _pzLocal(c); if (ln === "Table" || ln === "HugeTable") tableNodes.push(c); } }
  if (!tableNodes.length) throw new Error("No data tables found in this .pzfx file.");
  let selIdx = 0; const tseq = _pzFirst(root, "TableSequence");
  if (tseq) { const refs = _pzKids(tseq, "Ref"); const sel = refs.find((r) => _pzAttr(r, "Selected") === "1") || refs[0]; if (sel) { const id = _pzAttr(sel, "ID"); const k = tableNodes.findIndex((t) => _pzAttr(t, "ID") === id); if (k >= 0) selIdx = k; } }
  const built = tableNodes.map((tn, i) => {
    const grm = _pzfxGroupedRM(tn, strikeExclude);
    if (grm) { const gb = _pzfxBuildGroupedRM(grm); if (gb && gb.rows.length) return { i, title: _pzTitle(tn) || _pzAttr(tn, "ID") || "Table", ttype: "TwoWay", typeDesc: gb.typeDesc, columns: gb.columns, rows: gb.rows, shape: gb.shape, compact: gb.compact }; }
    const map = _pzMapTable(tn, strikeExclude); const b = _pzBuild(map);
    return b ? { i, title: map.title, ttype: map.ttype, typeDesc: PZFX_TDESC(map.ttype), columns: b.columns, rows: b.rows, shape: b.shape } : null;
  }).filter(Boolean);
  if (!built.length) throw new Error("This .pzfx file has no tables with data columns.");
  let activeIdx = built.findIndex((b) => b.i === selIdx); if (activeIdx < 0) activeIdx = 0;
  const projects = built.map((b, k) => {
    const otherTables = built.filter((_, j) => j !== k).map((o) => ({ title: o.title, typeDesc: o.typeDesc, shape: o.shape }));
    const sheets = b.compact ? [{ title: "Repeated-measures transpose applied", notes: 'This grouped Prism table stored a within-subjects factor down the rows and the groups across the columns (subjects in subcolumns). VibeStat transposed it so each subject is one case: a "Group" between-subjects column plus a compacted repeated-measures variable over "Trial" (' + b.compact.factors[0].levels.join(", ") + '). NOTE: a .pzfx file does not record which factor is repeated, so VibeStat assumed the subcolumn replicates are the SAME subjects across rows. If they are independent replicates instead, select the compact and click Expand to undo the repeated-measures grouping.' }].concat(infoSheets) : infoSheets;
    const info = { source: "GraphPad Prism (.pzfx)", prismVer, createdBy, infoSheets: sheets, activeTable: { title: b.title, typeDesc: b.typeDesc, shape: b.shape }, otherTables, tableCount: built.length };
    return { app: "VibeStat", version: 4, columns: b.columns, rows: b.rows, compacts: b.compact ? [b.compact] : [], analyses: [{ id: "pzfx_notes", type: "importnotes", roles: {}, info }], selAnalysis: "pzfx_notes", excluded: [], colW: {} };
  });
  return { projects, titles: built.map((b) => b.title), typeDescs: built.map((b) => b.typeDesc), shapes: built.map((b) => b.shape), activeIdx, count: built.length };
}
function parsePZFX(xmlText, strikeExclude) { const r = parsePZFXTables(xmlText, strikeExclude); return r.projects[r.activeIdx]; }

/* ---- GraphPad Prism .prism (modern zip container) importer: grouped/RM tables are
   transposed to VibeStat cases × within-levels and compacted. ------------------- */
async function _prismInflate(u8) {
  if (typeof DecompressionStream === "undefined") throw new Error("This environment cannot read .prism files (no DecompressionStream support).");
  const ds = new DecompressionStream("deflate-raw");
  const ab = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
  return new Uint8Array(ab);
}
// Minimal ZIP reader: walk the central directory, inflate only the entries we read.
async function _prismUnzip(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1; const minEocd = Math.max(0, bytes.length - 22 - 65536);
  for (let i = bytes.length - 22; i >= minEocd; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error("Not a valid .prism file (no ZIP directory found).");
  const count = dv.getUint16(eocd + 10, true), cdOff = dv.getUint32(eocd + 16, true);
  const dec = new TextDecoder(); const out = {}; let p = cdOff;
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    p += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith("/")) continue;
    if (!(name === "document.json" || name.indexOf("data/") === 0 || name.indexOf("analyses/") === 0)) continue;
    const lNameLen = dv.getUint16(localOff + 26, true), lExtraLen = dv.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const comp = bytes.subarray(dataStart, dataStart + compSize);
    out[name] = method === 0 ? comp : await _prismInflate(comp);
  }
  return out;
}
function _prismDsTitle(rj, uid) { try { const d = rj("data/sets/" + uid + ".json"); return typeof d.title === "string" ? d.title : (d.title && d.title.string) || uid; } catch (e) { return uid; } }
async function parsePrismBytes(bytes) {
  const strFromU8 = (u) => new TextDecoder().decode(u);
  const zip = await _prismUnzip(bytes);
  const rdText = (p) => { const u = zip[p]; if (!u) throw new Error("Prism file is missing " + p); return strFromU8(u); };
  const rj = (p) => JSON.parse(rdText(p));
  const doc = rj("document.json");
  const dataUids = (doc.sheets && doc.sheets.data) || [];
  if (!dataUids.length) throw new Error("No data sheets found in this .prism file.");
  let fn = { colFactor: "Group", rowFactor: "Trial", matchedFactor: "Subject" };
  for (const aUid of (doc.sheets.analyses || [])) { try { const pp = rj("analyses/" + aUid + "/parameters.json"); if (pp.content && pp.content.factorNames) { fn = Object.assign({}, fn, pp.content.factorNames); break; } } catch (e) {} }
  const createdBy = doc.createdBy ? (doc.createdBy.name + (doc.createdBy.version ? " " + doc.createdBy.version : "")) : "";
  const otherFor = (uid) => dataUids.filter((u) => u !== uid).map((u) => { try { const sh = rj("data/sheets/" + u + "/sheet.json"); return { title: sh.title || u, typeDesc: (sh.table && sh.table.format) || "data" }; } catch (e) { return { title: u, typeDesc: "data" }; } });
  const projects = [], titles = [], typeDescs = [], shapes = [];
  for (const uid of dataUids) {
    let sheet; try { sheet = rj("data/sheets/" + uid + "/sheet.json"); } catch (e) { continue; }
    const t = sheet.table || {};
    const title = sheet.title || (doc.sheetAttributesMap && doc.sheetAttributesMap[uid] && doc.sheetAttributesMap[uid].title) || "Data";
    const csvU = zip["data/tables/" + t.uid + "/data.csv"]; if (!csvU) continue;
    const matrix = strFromU8(csvU).replace(/\r/g, "").split("\n").filter((l) => l.length).map((l) => l.split(","));
    if (!matrix.length) continue;
    const groups = t.dataSets || [];
    const rep = t.replicatesCount || 0;
    const grouped = t.format === "grouped" && groups.length >= 1 && rep >= 1;
    if (grouped) {
      const nGroups = groups.length, nRows = matrix.length;
      const rowTitles = matrix.map((r, k) => (r[0] != null && r[0] !== "" ? r[0] : "Level" + (k + 1)));
      const groupTitles = groups.map((g) => _prismDsTitle(rj, g));
      const subTitles = {};
      if (t.subcolumnTitlesDataSet) { try { const sc = rj("data/sets/" + t.subcolumnTitlesDataSet + ".json"); (sc.titles || []).forEach((col) => { subTitles[col.column] = {}; (col.replicates || []).forEach((r) => (subTitles[col.column][r.replicate] = r.name)); }); } catch (e) {} }
      const columns = [{ id: "subj", name: fn.matchedFactor || "Subject", type: "string" }, { id: "grp", name: fn.colFactor || "Group", type: "category" }];
      rowTitles.forEach((rt, k) => columns.push({ id: "m" + k, name: rt, type: "real" }));
      const rows = [];
      for (let g = 0; g < nGroups; g++) for (let j = 0; j < rep; j++) {
        const ci = 1 + g * rep + j;
        const vals = rowTitles.map((_, k) => { const v = matrix[k] && matrix[k][ci]; return v === "" || v == null ? "" : Number(v); });
        if (vals.every((v) => v === "")) continue;
        const row = { subj: (subTitles[g] && subTitles[g][j]) || (groupTitles[g] + "-" + (j + 1)), grp: groupTitles[g] };
        rowTitles.forEach((_, k) => (row["m" + k] = vals[k]));
        rows.push(row);
      }
      const compact = { id: "cpRM", name: "Measure", factors: [{ name: fn.rowFactor || "Trial", levels: rowTitles.slice() }], leaves: rowTitles.map((_, k) => "m" + k) };
      const info = { source: "GraphPad Prism (.prism)", createdBy, prismVer: doc.formatVersion, activeTable: { title, typeDesc: "Grouped, repeated measures (transposed to cases × within levels)", shape: nGroups + " groups × " + rep + " subjects × " + nRows + " " + (fn.rowFactor || "levels") }, infoSheets: [{ title: "Repeated-measures transpose applied", notes: 'Prism stored the within-subjects factor "' + (fn.rowFactor || "Trial") + '" down the rows and the ' + nGroups + ' "' + (fn.colFactor || "Group") + '" groups across the columns (' + rep + ' subjects per group). VibeStat transposed this so each subject is one case: a "' + (fn.colFactor || "Group") + '" between-subjects column plus a compacted repeated-measures variable over "' + (fn.rowFactor || "Trial") + '" (' + rowTitles.join(", ") + ').' }], otherTables: otherFor(uid), tableCount: dataUids.length };
      projects.push({ app: "VibeStat", version: 4, columns, rows, compacts: [compact], analyses: [{ id: "prism_notes", type: "importnotes", roles: {}, info }], selAnalysis: "prism_notes", excluded: [], colW: {} });
      titles.push(title); typeDescs.push("Grouped (RM → transposed)"); shapes.push(nGroups + "×" + rep + "×" + nRows);
    } else {
      const ncol = Math.max.apply(null, matrix.map((r) => r.length));
      const hasRowLabels = matrix.some((r) => r[0] && isNaN(Number(r[0])));
      const columns = []; if (hasRowLabels) columns.push({ id: "lbl", name: "Row", type: "string" });
      const groupTitles = groups.map((g) => _prismDsTitle(rj, g));
      const dataCols = ncol - (hasRowLabels ? 1 : 0);
      for (let c = 0; c < dataCols; c++) columns.push({ id: "c" + c, name: groupTitles[c] || ("Col" + (c + 1)), type: "real" });
      const rows = matrix.map((r) => { const o = {}; let off = 0; if (hasRowLabels) { o.lbl = r[0]; off = 1; } for (let c = 0; c < dataCols; c++) { const v = r[c + off]; o["c" + c] = v === "" || v == null ? "" : Number(v); } return o; });
      const info = { source: "GraphPad Prism (.prism)", createdBy, prismVer: doc.formatVersion, activeTable: { title, typeDesc: "Imported as-is (" + (t.format || "table") + ")", shape: rows.length + " rows × " + dataCols + " cols" }, infoSheets: [], otherTables: otherFor(uid), tableCount: dataUids.length };
      projects.push({ app: "VibeStat", version: 4, columns, rows, compacts: [], analyses: [{ id: "prism_notes", type: "importnotes", roles: {}, info }], selAnalysis: "prism_notes", excluded: [], colW: {} });
      titles.push(title); typeDescs.push(t.format || "table"); shapes.push(rows.length + "×" + dataCols);
    }
  }
  if (!projects.length) throw new Error("This .prism file has no importable data tables.");
  return { count: projects.length, projects, titles, typeDescs, shapes, activeIdx: 0 };
}

// ---- environment I/O layer (Tauri build: native macOS/Windows dialogs + filesystem). ----
function _vsBasename(p) { const s = String(p); let i = s.lastIndexOf("/"); const b = s.lastIndexOf("\\"); if (b > i) i = b; return i >= 0 ? s.slice(i + 1) : s; }
async function ioOpenJSON() {
  const p = await tauriOpenDialog({ multiple: false, filters: [{ name: "VibeStat dataset", extensions: ["json", "vibestat"] }] });
  if (!p) return null;
  const d = JSON.parse(await readTextFile(p)); if (d && typeof d === "object") d.__name = _vsBasename(p); return d;
}
async function ioSaveJSON(project, name) {
  const p = await tauriSaveDialog({ defaultPath: name || "study.vibestat.json", filters: [{ name: "VibeStat dataset", extensions: ["json", "vibestat"] }] });
  if (!p) return false;
  await writeTextFile(p, JSON.stringify(project, null, 2));
  return _vsBasename(p);
}
async function ioImportTable() {
  const p = await tauriOpenDialog({ multiple: false, filters: [{ name: "Data files", extensions: ["csv", "tsv", "txt", "xlsx", "xls", "pzfx", "prism"] }] });
  if (!p) return null;
  const lower = p.toLowerCase(); let r;
  if (lower.endsWith(".pzfx")) { const rr = parsePZFXTables(await readTextFile(p)); r = rr.count > 1 ? { __pzfxMulti: true, projects: rr.projects, titles: rr.titles, typeDescs: rr.typeDescs, shapes: rr.shapes, activeIdx: rr.activeIdx } : rr.projects[rr.activeIdx]; }
  else if (lower.endsWith(".prism")) { const rr = await parsePrismBytes(await readFile(p)); r = rr.count > 1 ? { __pzfxMulti: true, projects: rr.projects, titles: rr.titles, typeDescs: rr.typeDescs, shapes: rr.shapes, activeIdx: rr.activeIdx } : rr.projects[rr.activeIdx]; }
  else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) r = parseXLSXBytes(await readFile(p));
  else r = parseCSVText(await readTextFile(p));
  if (r && typeof r === "object") r.__name = _vsBasename(p);
  return r;
}

/* =========================================================================
   STATISTICS ENGINE — genuine distributions via regularized incomplete beta
   ========================================================================= */
function gammln(xx) {
  const cof = [76.18009172947146, -86.50532032941677, 24.01409824083091,
    -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let x = xx, y = xx, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) { y += 1; ser += cof[j] / y; }
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}
function betacf(a, b, x) {
  const MAXIT = 400, EPS = 3e-14, FPMIN = 1e-300;
  let qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d; let h = d;
  for (let m = 1; m <= MAXIT; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; const del = d * c; h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}
function betai(a, b, x) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(gammln(a + b) - gammln(a) - gammln(b)
    + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return bt * betacf(a, b, x) / a;
  return 1 - bt * betacf(b, a, 1 - x) / b;
}
function tTwoTailedP(t, df) { if (df <= 0) return NaN; return betai(df / 2, 0.5, df / (df + t * t)); }
function fP(F, df1, df2) { if (!(F > 0) || df1 <= 0 || df2 <= 0) return NaN; return betai(df2 / 2, df1 / 2, df2 / (df2 + df1 * F)); }

const num = (v) => {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
};

/* Evaluate a per-row formula. Refs are [ColumnNameOrId]; a ref matching a COMPACT
   name/id expands to all that compact's leaf values in the row (so AVERAGE([RT (ms)])
   is the row mean over every leaf). Functions: AVERAGE/MEAN, SUM, MIN, MAX, COUNT,
   MEDIAN, STDEV/SD, VAR, PRODUCT, RANGE (aggregates); ABS, SQRT, LN, LOG/LOG10, EXP,
   ROUND, FLOOR, CEIL, SIGN (scalar). Operators + - * / ^ and parentheses. */
function evalFormula(formula, row, columns, compacts) {
  if (formula == null || !String(formula).trim()) return "";
  let s = String(formula).trim(); if (s[0] === "=") s = s.slice(1);
  const colKey = {}; columns.forEach((c) => { colKey[c.name.toLowerCase()] = c.id; colKey[c.id.toLowerCase()] = c.id; });
  const compKey = {}; (compacts || []).forEach((cp) => { compKey[cp.name.toLowerCase()] = cp.leaves; compKey[cp.id.toLowerCase()] = cp.leaves; });
  const refVals = (name) => {
    const k = name.trim().toLowerCase();
    if (compKey[k]) return compKey[k].map((id) => num(row[id])).filter((v) => v != null);
    if (colKey[k] != null) { const v = num(row[colKey[k]]); return [v == null ? NaN : v]; }
    throw new Error("unknown ref");
  };
  const toks = []; let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === " " || ch === "\t") { i++; continue; }
    if (ch === "[") { const j = s.indexOf("]", i); if (j < 0) throw new Error("unclosed"); toks.push({ t: "ref", v: s.slice(i + 1, j) }); i = j + 1; continue; }
    if ("+-*/^(),".includes(ch)) { toks.push({ t: ch }); i++; continue; }
    if (/[0-9.]/.test(ch)) { let j = i; while (j < s.length && /[0-9.eE]/.test(s[j])) j++; toks.push({ t: "num", v: parseFloat(s.slice(i, j)) }); i = j; continue; }
    if (/[A-Za-z_]/.test(ch)) { let j = i; while (j < s.length && /[A-Za-z0-9_]/.test(s[j])) j++; toks.push({ t: "id", v: s.slice(i, j) }); i = j; continue; }
    throw new Error("bad char");
  }
  let p = 0; const peek = () => toks[p], adv = () => toks[p++];
  const AGG = {
    average: (a) => a.reduce((x, y) => x + y, 0) / a.length, mean: (a) => AGG.average(a),
    sum: (a) => a.reduce((x, y) => x + y, 0), min: (a) => Math.min(...a), max: (a) => Math.max(...a),
    count: (a) => a.length, product: (a) => a.reduce((x, y) => x * y, 1), range: (a) => Math.max(...a) - Math.min(...a),
    median: (a) => { const b = [...a].sort((x, y) => x - y), n = b.length; return n % 2 ? b[(n - 1) / 2] : (b[n / 2 - 1] + b[n / 2]) / 2; },
    stdev: (a) => { const m = AGG.average(a); return Math.sqrt(a.reduce((x, y) => x + (y - m) ** 2, 0) / (a.length - 1)); },
    sd: (a) => AGG.stdev(a), var: (a) => { const m = AGG.average(a); return a.reduce((x, y) => x + (y - m) ** 2, 0) / (a.length - 1); },
  };
  const SCALAR = { abs: Math.abs, sqrt: Math.sqrt, ln: Math.log, log: (x) => Math.log10(x), log10: (x) => Math.log10(x), exp: Math.exp, round: Math.round, floor: Math.floor, ceil: Math.ceil, sign: Math.sign };
  function prim() {
    const tk = peek(); if (!tk) throw new Error("end");
    if (tk.t === "num") { adv(); return [tk.v]; }
    if (tk.t === "ref") { adv(); return refVals(tk.v); }
    if (tk.t === "(") { adv(); const v = expr(); if (!peek() || peek().t !== ")") throw new Error(")"); adv(); return v; }
    if (tk.t === "-") { adv(); return prim().map((x) => -x); }
    if (tk.t === "id") {
      adv(); const fn = tk.v.toLowerCase(); if (!peek() || peek().t !== "(") throw new Error("("); adv();
      const args = []; if (peek() && peek().t !== ")") { args.push(expr()); while (peek() && peek().t === ",") { adv(); args.push(expr()); } }
      if (!peek() || peek().t !== ")") throw new Error(")"); adv();
      if (AGG[fn]) { const all = args.flat().filter((v) => v != null && !isNaN(v)); return all.length ? [AGG[fn](all)] : [NaN]; }
      if (SCALAR[fn]) { const a = args.flat(); if (a.length !== 1) throw new Error("arity"); return [SCALAR[fn](a[0])]; }
      throw new Error("fn");
    }
    throw new Error("token");
  }
  const sc = (v) => { if (v.length !== 1) throw new Error("scalar"); return v[0]; };
  function pow() { let l = prim(); while (peek() && peek().t === "^") { adv(); l = [Math.pow(sc(l), sc(prim()))]; } return l; }
  function term() { let l = pow(); while (peek() && (peek().t === "*" || peek().t === "/")) { const op = adv().t; const r = pow(); l = [op === "*" ? sc(l) * sc(r) : sc(l) / sc(r)]; } return l; }
  function expr() { let l = term(); while (peek() && (peek().t === "+" || peek().t === "-")) { const op = adv().t; const r = term(); l = [op === "+" ? sc(l) + sc(r) : sc(l) - sc(r)]; } return l; }
  try { const v = expr(); if (p !== toks.length) throw new Error("trailing"); const r = v.length === 1 ? v[0] : NaN; return isNaN(r) ? "#ERR" : Math.round(r * 1e6) / 1e6; }
  catch (e) { return "#ERR"; }
}
const isFormula = (t) => t === "formula" || t === "formula_static";

function describe(values) {
  const xs = values.map(num).filter((v) => v !== null);
  const n = xs.length;
  if (n === 0) return { n: 0 };
  const sum = xs.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const ss = xs.reduce((a, b) => a + (b - mean) ** 2, 0);
  const variance = n > 1 ? ss / (n - 1) : 0;
  const sd = Math.sqrt(variance);
  const sorted = [...xs].sort((a, b) => a - b);
  const median = n % 2 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
  return { n, sum, mean, variance, sd, sem: sd / Math.sqrt(n), min: sorted[0], max: sorted[n - 1], range: sorted[n - 1] - sorted[0], median };
}

function regression(xv, yv) {
  const pairs = xv.map((x, i) => [num(x), num(yv[i])]).filter(([a, b]) => a !== null && b !== null);
  const n = pairs.length;
  if (n < 3) return { n };
  const mx = pairs.reduce((a, [x]) => a + x, 0) / n;
  const my = pairs.reduce((a, [, y]) => a + y, 0) / n;
  let sxx = 0, syy = 0, sxy = 0;
  for (const [x, y] of pairs) { sxx += (x - mx) ** 2; syy += (y - my) ** 2; sxy += (x - mx) * (y - my); }
  const slope = sxy / sxx, intercept = my - slope * mx;
  const r = sxy / Math.sqrt(sxx * syy), r2 = r * r;
  const ssReg = slope * sxy, ssRes = syy - ssReg, dfReg = 1, dfRes = n - 2;
  const msReg = ssReg / dfReg, msRes = ssRes / dfRes, F = msReg / msRes;
  const seSlope = Math.sqrt(msRes / sxx), tSlope = slope / seSlope;
  return { n, slope, intercept, r, r2, adjR2: 1 - (1 - r2) * (n - 1) / (n - 2), ssReg, ssRes, ssTotal: syy, dfReg, dfRes, msReg, msRes, F, pF: fP(F, dfReg, dfRes), seSlope, tSlope, pSlope: tTwoTailedP(tSlope, dfRes), rmsResidual: Math.sqrt(msRes), pairs, mx, my };
}

/* Invert a small square matrix via Gauss–Jordan with partial pivoting; null if singular. */
function matInv(M) {
  const n = M.length, A = M.map((r, i) => r.concat(Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
  for (let c = 0; c < n; c++) {
    let p = c; for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
    if (Math.abs(A[p][c]) < 1e-12) return null;
    const t = A[c]; A[c] = A[p]; A[p] = t; const d = A[c][c];
    for (let j = 0; j < 2 * n; j++) A[c][j] /= d;
    for (let r = 0; r < n; r++) { if (r === c) continue; const f = A[r][c]; for (let j = 0; j < 2 * n; j++) A[r][j] -= f * A[c][j]; }
  }
  return A.map((r) => r.slice(n));
}

/* Multiple linear regression (OLS). yRaw + XcolsRaw are row-aligned raw values;
   listwise deletion across Y and all predictors. Validated vs statsmodels OLS. */
function multipleRegression(yRaw, XcolsRaw, xnames, opts) {
  const ci = (opts && opts.ci) || 0.95, k = XcolsRaw.length;
  if (k === 0) return { error: "Assign at least one predictor (X)." };
  const yv = [], Xv = [];
  for (let i = 0; i < yRaw.length; i++) {
    const yy = num(yRaw[i]); if (yy === null) continue;
    const xs = XcolsRaw.map((c) => num(c[i])); if (xs.some((v) => v === null)) continue;
    yv.push(yy); Xv.push(xs);
  }
  const n = yv.length, p = k + 1;
  if (n <= p) return { error: `Only ${n} complete case(s) for ${k} predictor(s) + intercept; need n > ${p}.` };
  const X = Xv.map((r) => [1, ...r]);
  const A = Array.from({ length: p }, () => new Array(p).fill(0)), Xty = new Array(p).fill(0);
  for (let i = 0; i < n; i++) { const xi = X[i]; for (let a = 0; a < p; a++) { Xty[a] += xi[a] * yv[i]; for (let b = a; b < p; b++) A[a][b] += xi[a] * xi[b]; } }
  for (let a = 0; a < p; a++) for (let b = 0; b < a; b++) A[a][b] = A[b][a];
  const Ai = matInv(A); if (!Ai) return { error: "Predictors are collinear (X′X is singular) — drop a redundant predictor." };
  const beta = Ai.map((r) => r.reduce((s, v, j) => s + v * Xty[j], 0));
  const fitted = new Array(n), resid = new Array(n); let sse = 0;
  for (let i = 0; i < n; i++) { let f = 0; for (let a = 0; a < p; a++) f += X[i][a] * beta[a]; fitted[i] = f; const e = yv[i] - f; resid[i] = e; sse += e * e; }
  const ybar = yv.reduce((s, v) => s + v, 0) / n; let sst = 0; for (const v of yv) sst += (v - ybar) * (v - ybar);
  const ssr = sst - sse, dfM = k, dfR = n - p, dfT = n - 1, mse = sse / dfR, msr = ssr / dfM;
  const r2 = 1 - sse / sst, adjR2 = 1 - (sse / dfR) / (sst / dfT), F = msr / mse, pF = fP(F, dfM, dfR);
  const tc = tCrit(1 - ci, dfR), sdy = Math.sqrt(sst / dfT);
  const sdx = XcolsRaw.map((_, j) => { let m = 0; for (const r of Xv) m += r[j]; m /= n; let s2 = 0; for (const r of Xv) s2 += (r[j] - m) * (r[j] - m); return Math.sqrt(s2 / (n - 1)); });
  const coefs = beta.map((b, j) => {
    const se = Math.sqrt(mse * Ai[j][j]), t = b / se;
    return { name: j === 0 ? "Intercept" : (xnames[j - 1] || "X" + j), b, se, t, p: tP(t, dfR, "two"), lo: b - tc * se, hi: b + tc * se, std: j === 0 ? null : b * sdx[j - 1] / sdy };
  });
  return { n, k, coefs, r2, adjR2, F, pF, dfM, dfR, dfT, sse, ssr, sst, mse, msr, rmse: Math.sqrt(mse), ci, resid, fitted };
}

/* Least-squares fit via normal equations; returns SSE, coefficients, and (X′X)⁻¹. */
function glmFit(cols, y) {
  const p = cols.length, n = y.length;
  const A = Array.from({ length: p }, () => new Array(p).fill(0)), b = new Array(p).fill(0);
  for (let i = 0; i < p; i++) { for (let j = i; j < p; j++) { let s = 0; for (let k = 0; k < n; k++) s += cols[i][k] * cols[j][k]; A[i][j] = s; A[j][i] = s; } let sb = 0; for (let k = 0; k < n; k++) sb += cols[i][k] * y[k]; b[i] = sb; }
  const Ai = matInv(A); if (!Ai) return null;
  const beta = Ai.map((r) => r.reduce((s, v, j) => s + v * b[j], 0));
  let yy = 0; for (let k = 0; k < n; k++) yy += y[k] * y[k]; let bz = 0; for (let i = 0; i < p; i++) bz += beta[i] * b[i];
  return { sse: yy - bz, beta, Ai };
}

/* General linear model / ANCOVA. Categorical factors (full-factorial or main-effects) +
   continuous covariates (main effects). Type III tests use sum-to-zero coding; parameter
   estimates use treatment coding (first level = reference). Validated vs statsmodels. */
function glmAnalyze(yRaw, factorCols, covCols, factorNames, covNames, opts) {
  const factorial = opts.factorial !== false, ci = opts.ci || 0.95;
  if (!factorNames.length && !covNames.length) return { error: "Assign at least one factor or covariate." };
  const rows = [];
  for (let i = 0; i < yRaw.length; i++) {
    const yy = num(yRaw[i]); if (yy === null) continue;
    const fv = factorCols.map((c) => c[i]); if (fv.some((v) => v === "" || v == null)) continue;
    const cv = covCols.map((c) => num(c[i])); if (cv.some((v) => v === null)) continue;
    rows.push({ y: yy, f: fv.map(String), c: cv });
  }
  const n = rows.length;
  const levels = factorNames.map((_, fi) => { const s = [...new Set(rows.map((r) => r.f[fi]))]; s.sort((a, b) => (isFinite(+a) && isFinite(+b)) ? (+a - +b) : (a < b ? -1 : a > b ? 1 : 0)); return s; });
  let effects = subsets(factorNames.map((_, i) => i)).filter((E) => E.length > 0);
  if (!factorial) effects = effects.filter((E) => E.length === 1);
  effects.sort((a, b) => a.length - b.length || a.join().localeCompare(b.join()));
  function build(coding) {
    const cols = [new Array(n).fill(1)], labels = ["Intercept"], terms = [];
    for (const E of effects) {
      let combos = [[]];
      E.forEach((fi) => { const k = levels[fi].length, nxt = []; for (const c of combos) for (let j = 0; j < k - 1; j++) nxt.push([...c, j]); combos = nxt; });
      const idx = [];
      for (const combo of combos) {
        const col = new Array(n);
        for (let i = 0; i < n; i++) { let v = 1; for (let t = 0; t < E.length; t++) { const fi = E[t], L = levels[fi], li = L.indexOf(rows[i].f[fi]); v *= coding === "sum" ? (li === L.length - 1 ? -1 : (li === combo[t] ? 1 : 0)) : (li === combo[t] + 1 ? 1 : 0); } col[i] = v; }
        idx.push(cols.length); cols.push(col);
        labels.push(E.map((fi, t) => `${factorNames[fi]}[${levels[fi][coding === "sum" ? combo[t] : combo[t] + 1]}]`).join(":"));
      }
      terms.push({ disp: E.map((fi) => factorNames[fi]).join(" \u00d7 "), idx, df: idx.length });
    }
    covNames.forEach((nm, k) => { cols.push(rows.map((r) => r.c[k])); terms.push({ disp: nm, idx: [cols.length - 1], df: 1 }); labels.push(nm); });
    if (opts.slopes) {
      for (let fi = 0; fi < factorNames.length; fi++) for (let ci = 0; ci < covNames.length; ci++) {
        const L = levels[fi], idx = [];
        for (let j = 0; j < L.length - 1; j++) {
          const col = new Array(n);
          for (let i = 0; i < n; i++) { const li = L.indexOf(rows[i].f[fi]), cc = coding === "sum" ? (li === L.length - 1 ? -1 : (li === j ? 1 : 0)) : (li === j + 1 ? 1 : 0); col[i] = cc * rows[i].c[ci]; }
          idx.push(cols.length); cols.push(col); labels.push(`${factorNames[fi]}[${L[coding === "sum" ? j : j + 1]}]:${covNames[ci]}`);
        }
        terms.push({ disp: `${factorNames[fi]} \u00d7 ${covNames[ci]}`, idx, df: L.length - 1 });
      }
    }
    return { cols, terms, labels };
  }
  const y = rows.map((r) => r.y);
  let yy = 0, sy = 0; for (const v of y) { yy += v * v; sy += v; } const sst = yy - sy * sy / n;
  const S = build("sum"), full = glmFit(S.cols, y);
  if (!full) return { error: "Model is singular — collinear predictors or an empty factor cell." };
  const p = S.cols.length, dfR = n - p;
  if (dfR <= 0) return { error: `Only ${n} complete case(s) for ${p} parameters; need n > ${p}.` };
  const mse = full.sse / dfR;
  const effOut = S.terms.map((tm) => { const keep = S.cols.filter((_, c) => !tm.idx.includes(c)); const red = glmFit(keep, y); const ss = red ? red.sse - full.sse : NaN; const ms = ss / tm.df, F = ms / mse; return { name: tm.disp, df: tm.df, ss, ms, F, p: fP(F, tm.df, dfR) }; });
  const T = build("treat"), tf = glmFit(T.cols, y), tc = tCrit(1 - ci, dfR);
  const params = tf.beta.map((b, j) => { const se = Math.sqrt(mse * tf.Ai[j][j]), t = b / se; return { name: T.labels[j], b, se, t, p: tP(t, dfR, "two"), lo: b - tc * se, hi: b + tc * se }; });
  const ssModel = sst - full.sse, dfModel = p - 1, Fmodel = (ssModel / dfModel) / mse;
  const fitted = y.map((_, i) => S.cols.reduce((s, col, j) => s + full.beta[j] * col[i], 0));
  const resid = y.map((v, i) => v - fitted[i]);
  return { n, levels, factorNames, effects: effOut, residual: { df: dfR, ss: full.sse, ms: mse }, total: { df: n - 1, ss: sst }, r2: 1 - full.sse / sst, adjR2: 1 - (full.sse / dfR) / (sst / (n - 1)), Fmodel, pModel: fP(Fmodel, dfModel, dfR), dfModel, ssModel, params, ci, factorial, resid, fitted };
}

function unpairedT(yv, groupv) {
  const groups = {};
  yv.forEach((y, i) => {
    const g = groupv[i], v = num(y);
    if (g === "" || g === null || g === undefined || v === null) return;
    (groups[g] = groups[g] || []).push(v);
  });
  const keys = Object.keys(groups);
  if (keys.length !== 2) return { error: keys.length, keys };
  const [a, b] = keys.map((k) => describe(groups[k]));
  const [ka, kb] = keys;
  if (a.n < 2 || b.n < 2) return { error: "size", keys };
  const dfp = a.n + b.n - 2;
  const sp2 = ((a.n - 1) * a.variance + (b.n - 1) * b.variance) / dfp;
  const se = Math.sqrt(sp2 * (1 / a.n + 1 / b.n));
  const t = (a.mean - b.mean) / se;
  return { groups: [{ key: ka, ...a }, { key: kb, ...b }], meanDiff: a.mean - b.mean, df: dfp, t, p: tTwoTailedP(t, dfp), se };
}

/* ---- tailed p-values + comparison engines (paired / unpaired) ---- */
// Standard normal CDF via a full-precision erfc built on the regularized upper
// incomplete gamma Q(1/2, x^2) (Numerical Recipes gser/gcf). Accurate deep into tails.
function gser(a, x) {
  const ITMAX = 400, EPS = 3e-16; let ap = a, sum = 1 / a, del = sum;
  for (let n = 1; n <= ITMAX; n++) { ap++; del *= x / ap; sum += del; if (Math.abs(del) < Math.abs(sum) * EPS) break; }
  return sum * Math.exp(-x + a * Math.log(x) - gammln(a));
}
function gcf(a, x) {
  const ITMAX = 400, EPS = 3e-16, FPMIN = 1e-300;
  let b = x + 1 - a, c = 1 / FPMIN, d = 1 / b, h = d;
  for (let i = 1; i <= ITMAX; i++) { const an = -i * (i - a); b += 2; d = an * d + b; if (Math.abs(d) < FPMIN) d = FPMIN; c = b + an / c; if (Math.abs(c) < FPMIN) c = FPMIN; d = 1 / d; const del = d * c; h *= del; if (Math.abs(del - 1) < EPS) break; }
  return Math.exp(-x + a * Math.log(x) - gammln(a)) * h;
}
function gammq(a, x) { if (x < 0 || a <= 0) return NaN; if (x === 0) return 1; return x < a + 1 ? 1 - gser(a, x) : gcf(a, x); }
function erfc(x) { return x >= 0 ? gammq(0.5, x * x) : 2 - gammq(0.5, x * x); }
const normCdf = (z) => 0.5 * erfc(-z / Math.SQRT2);

/* Inverse standard-normal CDF (Acklam's rational approximation, ~1e-9). */
function invNorm(p) {
  if (p <= 0) return -Infinity; if (p >= 1) return Infinity;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.383577518672690e2, -3.066479806614716e1, 2.506628277459239e0];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838e0, -2.549732539343734e0, 4.374664141464968e0, 2.938163982698783e0];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996e0, 3.754408661907416e0];
  const pl = 0.02425; let q;
  if (p < pl) { q = Math.sqrt(-2 * Math.log(p)); return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  if (p <= 1 - pl) { q = p - 0.5; const r = q * q; return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
  q = Math.sqrt(-2 * Math.log(1 - p)); return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
}
const _poly = (c, t) => { let r = 0, p = 1; for (let i = 0; i < c.length; i++) { r += c[i] * p; p *= t; } return r; };

/* Shapiro–Wilk normality test (Royston's AS R94). Validated vs scipy.stats.shapiro. */
function shapiroWilk(xIn) {
  const x = xIn.map(Number).sort((a, b) => a - b), n = x.length;
  if (n < 3) return { error: "Shapiro–Wilk needs n ≥ 3." };
  const xbar = x.reduce((s, v) => s + v, 0) / n; let ssd = 0; for (const v of x) ssd += (v - xbar) ** 2;
  if (ssd <= 0) return { error: "Zero variance." };
  const c1 = [0, 0.221157, -0.147981, -2.071190, 4.434685, -2.706056], c2 = [0, 0.042981, -0.293762, -1.752461, 5.682633, -3.582633];
  const nn2 = Math.floor(n / 2), m = new Array(nn2);
  for (let i = 1; i <= nn2; i++) m[i - 1] = invNorm((i - 0.375) / (n + 0.25));
  let summ2 = 0; for (let i = 0; i < nn2; i++) summ2 += m[i] * m[i]; summ2 *= 2;
  const ssumm2 = Math.sqrt(summ2), rsn = 1 / Math.sqrt(n), a = new Array(nn2); let fac, i1;
  const a1 = _poly(c1, rsn) - m[0] / ssumm2;
  if (n > 5) { const a2 = _poly(c2, rsn) - m[1] / ssumm2; fac = Math.sqrt((summ2 - 2 * m[0] * m[0] - 2 * m[1] * m[1]) / (1 - 2 * a1 * a1 - 2 * a2 * a2)); a[0] = a1; a[1] = a2; i1 = 2; }
  else { fac = Math.sqrt((summ2 - 2 * m[0] * m[0]) / (1 - 2 * a1 * a1)); a[0] = a1; i1 = 1; }
  for (let i = i1; i < nn2; i++) a[i] = -m[i] / fac;
  if (n === 3) a[0] = Math.SQRT1_2;
  let num = 0; for (let i = 0; i < nn2; i++) num += a[i] * (x[n - 1 - i] - x[i]);
  const W = num * num / ssd; let p;
  if (n === 3) { p = 6 / Math.PI * (Math.asin(Math.sqrt(W)) - Math.asin(Math.sqrt(0.75))); p = Math.max(0, Math.min(1, p)); }
  else { const w1 = 1 - W; let z;
    if (n <= 11) { const g = _poly([-2.273, 0.459], n), mu = _poly([0.5440, -0.39978, 0.025054, -6.714e-4], n), sig = Math.exp(_poly([1.3822, -0.77857, 0.062767, -0.0020322], n)); z = (-Math.log(g - Math.log(w1)) - mu) / sig; }
    else { const ln = Math.log(n), mu = _poly([-1.5861, -0.31082, -0.083751, 0.0038915], ln), sig = Math.exp(_poly([-0.4803, -0.082676, 0.0030302], ln)); z = (Math.log(w1) - mu) / sig; }
    p = 1 - normCdf(z); }
  return { W, p, n };
}

/* Levene's test (Brown–Forsythe, median-centered by default). Validated vs scipy.stats.levene. */
function leveneTest(groups, center) {
  const k = groups.length, N = groups.reduce((s, g) => s + g.length, 0);
  if (k < 2 || N - k <= 0) return { error: "Need ≥ 2 groups." };
  const Z = groups.map((g) => { const c = center === "mean" ? g.reduce((s, v) => s + v, 0) / g.length : quantileSorted([...g].sort((x, y) => x - y), 0.5); return g.map((v) => Math.abs(v - c)); });
  const Zbar = Z.map((z) => z.reduce((s, v) => s + v, 0) / z.length), Zg = Z.flat().reduce((s, v) => s + v, 0) / N;
  let nu = 0; Z.forEach((z, i) => nu += z.length * (Zbar[i] - Zg) ** 2);
  let de = 0; Z.forEach((z, i) => z.forEach((v) => de += (v - Zbar[i]) ** 2));
  const W = ((N - k) / (k - 1)) * (nu / de);
  return { W, df1: k - 1, df2: N - k, p: fP(W, k - 1, N - k) };
}

/* Bartlett's test of equal variances (χ²). Validated vs scipy.stats.bartlett. */
function bartlettTest(groups) {
  const k = groups.length, N = groups.reduce((s, g) => s + g.length, 0);
  if (k < 2 || N - k <= 0) return { error: "Need ≥ 2 groups." };
  const vars = groups.map((g) => { const m = g.reduce((s, v) => s + v, 0) / g.length; return g.reduce((s, v) => s + (v - m) ** 2, 0) / (g.length - 1); });
  const ni = groups.map((g) => g.length), sp2 = ni.reduce((s, n, i) => s + (n - 1) * vars[i], 0) / (N - k);
  const numr = (N - k) * Math.log(sp2) - ni.reduce((s, n, i) => s + (n - 1) * Math.log(vars[i]), 0);
  const C = 1 + (1 / (3 * (k - 1))) * (ni.reduce((s, n) => s + 1 / (n - 1), 0) - 1 / (N - k));
  const chi2 = numr / C;
  return { chi2, df: k - 1, p: chiSqSf(chi2, k - 1) };
}

/* One-way decomposition + planned-contrast / orthogonal-polynomial trend machinery.
   Contrast test L=Σcᵢȳᵢ, SE=√(MSE·Σcᵢ²/nᵢ), SS=L²/Σ(cᵢ²/nᵢ); validated vs statsmodels. */
function oneWayGroups(groups) {
  const k = groups.length, ni = groups.map((g) => g.length), N = ni.reduce((a, b) => a + b, 0);
  const means = groups.map((g) => g.length ? g.reduce((s, v) => s + v, 0) / g.length : NaN);
  const grand = groups.reduce((s, g) => s + g.reduce((a, b) => a + b, 0), 0) / N;
  let ssb = 0; groups.forEach((g, i) => ssb += ni[i] * (means[i] - grand) ** 2);
  let ssw = 0; groups.forEach((g, i) => g.forEach((v) => ssw += (v - means[i]) ** 2));
  const dfb = k - 1, dfw = N - k, mse = ssw / dfw, F = (ssb / dfb) / mse;
  return { k, ni, N, means, grand, ssb, ssw, dfb, dfw, mse, F, p: fP(F, dfb, dfw) };
}
function contrastStat(c, ow) {
  const L = c.reduce((s, ci, i) => s + ci * ow.means[i], 0);
  const denom = c.reduce((s, ci, i) => s + ci * ci / ow.ni[i], 0);
  const se = Math.sqrt(ow.mse * denom), ss = L * L / denom, t = L / se;
  return { L, se, t, ss, F: ss / ow.mse, df1: 1, df2: ow.dfw, p: tP(t, ow.dfw, "two") };
}
function orthoPoly(scores) {
  const basis = [new Array(scores.length).fill(1)], out = [];
  for (let deg = 1; deg <= scores.length - 1; deg++) {
    let v = scores.map((s) => Math.pow(s, deg));
    for (const b of basis) { const d = b.reduce((s, x, i) => s + x * v[i], 0) / b.reduce((s, x) => s + x * x, 0); v = v.map((x, i) => x - d * b[i]); }
    basis.push(v); out.push(v);
  }
  return out;
}

/* Within-subjects (repeated-measures) contrasts: per-subject marginal means at each level
   of the target within factor (collapsing over other within factors), then a one-sample t
   on the contrast scores — each contrast tested with its OWN error (df = n−1), the
   sphericity-robust convention used by SPSS. Validated vs scipy ttest_1samp. */
function withinContrastScores(long, targetName, targetLevels) {
  const bySubj = new Map(), subj = long.subject, tv = long.within[targetName], y = long.y;
  for (let i = 0; i < y.length; i++) { const s = subj[i]; if (!bySubj.has(s)) bySubj.set(s, new Map()); const mm = bySubj.get(s); const L = String(tv[i]); if (!mm.has(L)) mm.set(L, [0, 0]); const e = mm.get(L); e[0] += y[i]; e[1] += 1; }
  const subjects = [], cells = [];
  for (const [s, mm] of bySubj) { const vec = targetLevels.map((L) => { const e = mm.get(String(L)); return e ? e[0] / e[1] : null; }); if (vec.every((v) => v != null)) { subjects.push(s); cells.push(vec); } }
  return { subjects, cells };
}
function withinContrastTest(c, cells) {
  const norm = Math.sqrt(c.reduce((s, x) => s + x * x, 0)) || 1, cn = c.map((x) => x / norm);
  const scores = cells.map((vec) => c.reduce((s, ci, i) => s + ci * vec[i], 0));
  const scoresN = cells.map((vec) => cn.reduce((s, ci, i) => s + ci * vec[i], 0));
  const n = cells.length, mean = scores.reduce((s, v) => s + v, 0) / n;
  const variance = scores.reduce((s, x) => s + (x - mean) ** 2, 0) / (n - 1), se = Math.sqrt(variance / n), t = mean / se;
  const mN = scoresN.reduce((s, x) => s + x, 0) / n, ss = n * mN * mN;
  return { L: mean, se, t, df: n - 1, n, p: tP(t, n - 1, "two"), ss, F: t * t };
}

/* Noncentral F CDF: Poisson(λ/2)-weighted sum of regularized incomplete betas. Validated vs scipy.stats.ncf. */
function ncFcdf(x, df1, df2, lam) {
  if (x <= 0) return 0;
  const y = (df1 * x) / (df1 * x + df2);
  if (lam <= 0) return betai(df1 / 2, df2 / 2, y);
  const half = lam / 2, maxJ = Math.max(1000, Math.ceil(half * 4) + 200);
  let sum = 0, w = Math.exp(-half);
  for (let j = 0; j <= maxJ; j++) {
    if (w > 1e-300) sum += w * betai(df1 / 2 + j, df2 / 2, y);
    w *= half / (j + 1);
    if (j > half && w < 1e-14) break;
  }
  return sum;
}
/* Confidence interval for the F noncentrality λ by inverting the noncentral-F CDF (monotone in λ). */
function ncFlambdaCI(F, df1, df2, conf) {
  const aL = (1 - conf) / 2, aU = 1 - aL;
  const solve = (target) => {
    if (ncFcdf(F, df1, df2, 0) <= target) return 0;
    let lo = 0, hi = 1;
    while (ncFcdf(F, df1, df2, hi) > target && hi < 1e7) hi *= 2;
    for (let it = 0; it < 200; it++) { const mid = (lo + hi) / 2; (ncFcdf(F, df1, df2, mid) > target) ? (lo = mid) : (hi = mid); }
    return (lo + hi) / 2;
  };
  return [solve(aU), solve(aL)];
}
/* Partial η² point estimate (from F) and CI (Smithson noncentral-F method). */
function etaSqPartialCI(F, df1, df2, conf) {
  const [lamL, lamU] = ncFlambdaCI(F, df1, df2, conf), N1 = df1 + df2 + 1;
  return [lamL / (lamL + N1), lamU / (lamU + N1)];
}

/* Noncentral t CDF (series via regularized incomplete beta, AS 243 style). Validated vs scipy.stats.nct. */
function nctCDF(t, nu, delta) {
  if (t < 0) return 1 - nctCDF(-t, nu, -delta);
  const x = t * t / (t * t + nu);
  if (x <= 0) return normCdf(-delta);
  const half = delta * delta / 2, sgn = delta < 0 ? -1 : 1, ldel = Math.log(Math.abs(delta) || 1e-300);
  let sum = 0;
  for (let j = 0; j < 3000; j++) {
    const Ip = betai(j + 0.5, nu / 2, x), Iq = betai(j + 1, nu / 2, x);
    const pj = (half === 0 && j > 0) ? 0 : Math.exp(-half + j * Math.log(half || 1e-300) - gammln(j + 1));
    const qj = (delta === 0) ? 0 : sgn * Math.exp(-half + j * Math.log(half || 1e-300) + ldel - 0.5 * Math.log(2) - gammln(j + 1.5));
    sum += pj * Ip + qj * Iq;
    if (j > half && pj < 1e-17 && Math.abs(qj) < 1e-17) break;
  }
  return normCdf(-delta) + 0.5 * sum;
}
/* CI for the t noncentrality δ by inverting the noncentral-t CDF (monotone decreasing in δ). */
function nctDeltaCI(t, df, conf) {
  const aL = (1 - conf) / 2, aU = 1 - aL;
  const solve = (target) => {
    let lo = -1, hi = 1;
    while (nctCDF(t, df, lo) < target && lo > -1e7) lo *= 2;
    while (nctCDF(t, df, hi) > target && hi < 1e7) hi *= 2;
    for (let it = 0; it < 200; it++) { const mid = (lo + hi) / 2; (nctCDF(t, df, mid) > target) ? (lo = mid) : (hi = mid); }
    return (lo + hi) / 2;
  };
  return [solve(aU), solve(aL)];
}
/* Two-group standardized mean difference: Cohen's d (pooled), Hedges' g (bias-corrected),
   Glass's Δ (each SD), with a noncentral-t CI for d. Validated vs pingouin. */
function cohenTwoGroup(m1, s1, n1, m2, s2, n2, conf) {
  const df = n1 + n2 - 2, sp = Math.sqrt(((n1 - 1) * s1 * s1 + (n2 - 1) * s2 * s2) / df);
  const d = (m1 - m2) / sp, J = 1 - 3 / (4 * (n1 + n2) - 9), a = Math.sqrt(n1 * n2 / (n1 + n2));
  const tp = (m1 - m2) / (sp * Math.sqrt(1 / n1 + 1 / n2)), [dL, dU] = nctDeltaCI(tp, df, conf);
  return { d, g: J * d, J, glass1: (m1 - m2) / s1, glass2: (m1 - m2) / s2, dCI: [dL / a, dU / a], gCI: [J * dL / a, J * dU / a] };
}
/* Paired / one-sample standardized mean difference (Cohen's d_z) with noncentral-t CI. */
function cohenPaired(t, n, conf) {
  const dz = t / Math.sqrt(n), J = 1 - 3 / (4 * (n - 1) - 1), a = Math.sqrt(n), [dL, dU] = nctDeltaCI(t, n - 1, conf);
  return { dz, gz: J * dz, J, dCI: [dL / a, dU / a], gCI: [J * dL / a, J * dU / a] };
}
/* =========================================================================
   POWER ANALYSIS — a priori / post-hoc power & sample size.
   Uses the same validated noncentral distributions as the effect-size CIs:
   noncentral t (nctCDF) for mean tests, noncentral F (ncFcdf) for ANOVA,
   noncentral chi-square (Poisson-weighted central chi-square via gammq) for
   chi-square/GoF, and the Fisher-z / normal approximation (pwr convention,
   with bias correction) for correlation and two-proportion z-tests.
   Validated vs statsmodels.stats.power and pingouin / pwr. ======================= */
function _tqU2(alpha, df) { return tCrit(alpha, df); }       // two-sided critical t = t_{1-alpha/2}
function _tqU1(alpha, df) { return tCrit(2 * alpha, df); }    // one-sided critical t = t_{1-alpha}
function _chi2CDF(x, df) { return x <= 0 ? 0 : 1 - gammq(df / 2, x / 2); }
function _chi2CritU(alpha, df) { let lo = 0, hi = Math.max(20, df * 2 + 20); while (_chi2CDF(hi, df) < 1 - alpha && hi < 1e7) hi *= 2; for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; (_chi2CDF(m, df) < 1 - alpha) ? (lo = m) : (hi = m); } return (lo + hi) / 2; }
function _ncChi2CDF(x, df, lam) { if (lam <= 0) return _chi2CDF(x, df); const half = lam / 2, maxJ = Math.max(1000, Math.ceil(half * 4) + 200); let s = 0, w = Math.exp(-half); for (let j = 0; j <= maxJ; j++) { if (w > 1e-300) s += w * _chi2CDF(x, df + 2 * j); w *= half / (j + 1); if (j > half && w < 1e-14) break; } return s; }
function _fCritU(alpha, df1, df2) { let lo = 0, hi = 10; while (ncFcdf(hi, df1, df2, 0) < 1 - alpha && hi < 1e7) hi *= 2; for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; (ncFcdf(m, df1, df2, 0) < 1 - alpha) ? (lo = m) : (hi = m); } return (lo + hi) / 2; }
function powerT2(d, n1, n2, alpha, tail) { const df = n1 + n2 - 2; if (df <= 0) return NaN; const nc = Math.abs(d) * Math.sqrt(n1 * n2 / (n1 + n2)); if (tail === "one") { const tc = _tqU1(alpha, df); return 1 - nctCDF(tc, df, nc); } const tc = _tqU2(alpha, df); return (1 - nctCDF(tc, df, nc)) + nctCDF(-tc, df, nc); }
function powerT1(d, n, alpha, tail) { const df = n - 1; if (df <= 0) return NaN; const nc = Math.abs(d) * Math.sqrt(n); if (tail === "one") { const tc = _tqU1(alpha, df); return 1 - nctCDF(tc, df, nc); } const tc = _tqU2(alpha, df); return (1 - nctCDF(tc, df, nc)) + nctCDF(-tc, df, nc); }
function powerAnova(f, k, n, alpha) { const df1 = k - 1, df2 = k * (n - 1); if (df1 < 1 || df2 < 1) return NaN; const lam = f * f * k * n, fc = _fCritU(alpha, df1, df2); return 1 - ncFcdf(fc, df1, df2, lam); }
function powerCorr(r, n, alpha, tail) {
  if (n <= 3) return NaN;
  const dof = n - 2, rr = Math.max(-0.999999, Math.min(0.999999, r));
  const tc = tail === "one" ? _tqU1(alpha, dof) : _tqU2(alpha, dof);
  const rc = Math.sqrt(tc * tc / (tc * tc + dof));
  const zr = Math.atanh(rr) + rr / (2 * (n - 1)), zrc = Math.atanh(rc), s = Math.sqrt(n - 3);
  return normCdf((zr - zrc) * s) + normCdf((-zr - zrc) * s);
}
function powerProp(h, n1, n2, alpha, tail) { const nc = Math.abs(h) * Math.sqrt(n1 * n2 / (n1 + n2)); if (tail === "one") { const za = invNorm(1 - alpha); return normCdf(nc - za); } const za = invNorm(1 - alpha / 2); return normCdf(nc - za) + normCdf(-nc - za); }
function powerChi2(w, N, df, alpha) { if (df < 1 || N < 1) return NaN; const lam = w * w * N, xc = _chi2CritU(alpha, df); return 1 - _ncChi2CDF(xc, df, lam); }
function powerValue(fam, p) {
  switch (fam) {
    case "t2": return powerT2(p.es, p.n, p.n, p.alpha, p.tail);
    case "t1": return powerT1(p.es, p.n, p.alpha, p.tail);
    case "anova": return powerAnova(p.es, p.k, p.n, p.alpha);
    case "corr": return powerCorr(p.es, p.n, p.alpha, p.tail);
    case "prop": return powerProp(p.es, p.n, p.n, p.alpha, p.tail);
    case "chi2": return powerChi2(p.es, p.n, p.df, p.alpha);
    default: return NaN;
  }
}
const POWER_MINN = { t2: 2, t1: 2, anova: 2, corr: 4, prop: 2, chi2: 1 };
// Solve a power scenario. Returns the solved quantity plus a power-vs-n curve.
function solvePower(s) {
  const pn = (v, d) => { const x = parseFloat(v); return isFinite(x) ? x : d; };
  const cl = (v, d, lo, hi) => { const x = pn(v, d); return Math.min(hi, Math.max(lo, x)); };
  const fam = s.family || "t2", sf = s.solveFor || "n", tail = s.tail || "two";
  const alpha = cl(s.alpha, 0.05, 1e-9, 0.5), targetPow = cl(s.power, 0.8, 1e-6, 0.999999);
  const k = Math.max(2, Math.round(pn(s.k, 3))), df = Math.max(1, Math.round(pn(s.df, 1)));
  const minN = POWER_MINN[fam], totalFam = (fam === "corr" || fam === "chi2");
  const pv = (es, n) => powerValue(fam, { es, n: Math.max(minN, Math.round(n)), k, df, alpha, tail });
  const out = { fam, sf, tail, alpha, k, df, totalFam };
  if (sf === "power") {
    out.es = pn(s.es, 0.5); out.n = Math.max(minN, Math.round(pn(s.n, 30)));
    out.power = pv(out.es, out.n); out.result = out.power;
  } else if (sf === "n") {
    out.es = pn(s.es, 0.5); out.targetPow = targetPow;
    let hi = minN, g = 0; while (pv(out.es, hi) < targetPow && hi < 2e6) { hi = Math.ceil(hi * 1.7) + 1; if (++g > 90) break; }
    if (pv(out.es, hi) < targetPow) { out.n = Infinity; out.power = NaN; }
    else { let a = minN, b = hi; while (b - a > 1) { const m = Math.floor((a + b) / 2); (pv(out.es, m) >= targetPow) ? (b = m) : (a = m); } out.n = Math.max(minN, b); out.power = pv(out.es, out.n); }
    out.result = out.n;
  } else { // es
    out.n = Math.max(minN, Math.round(pn(s.n, 30))); out.targetPow = targetPow;
    let hi = 0.001, g = 0; while (pv(hi, out.n) < targetPow && hi < 1e4) { hi *= 1.8; if (++g > 90) break; }
    if (pv(hi, out.n) < targetPow) { out.es = Infinity; out.power = NaN; }
    else { let a = 1e-6, b = hi; for (let i = 0; i < 200; i++) { const m = (a + b) / 2; (pv(m, out.n) >= targetPow) ? (b = m) : (a = m); } out.es = b; out.power = pv(out.es, out.n); }
    out.result = out.es;
  }
  const esUse = out.es, nOp = isFinite(out.n) ? out.n : Math.max(minN, Math.round(pn(s.n, 30)));
  const loN = minN, hiN = Math.max(nOp * 2, loN + 12), pts = [];
  if (isFinite(esUse)) for (let i = 0; i <= 40; i++) { const nn = Math.round(loN + (hiN - loN) * i / 40); pts.push({ n: nn, power: pv(esUse, nn) }); }
  out.curve = { xLabel: totalFam ? "sample size (N)" : "n per group", points: pts, opN: isFinite(nOp) ? nOp : null, opPow: out.power, target: sf !== "power" ? targetPow : null };
  return out;
}
// One/two-tailed p from a t statistic. tail: "two" | "upper" | "lower".
function tP(t, df, tail) {
  if (df <= 0 || !isFinite(t)) return NaN;
  const two = tTwoTailedP(t, df); // P(|T| >= |t|), uses t^2 so sign-independent
  if (tail === "upper") return t >= 0 ? two / 2 : 1 - two / 2;
  if (tail === "lower") return t <= 0 ? two / 2 : 1 - two / 2;
  return two;
}
// One/two-tailed p from a z statistic.
function zP(z, tail) {
  if (!isFinite(z)) return NaN;
  if (tail === "upper") return normCdf(-z);
  if (tail === "lower") return normCdf(z);
  return 2 * normCdf(-Math.abs(z));
}
// Two-sided/one-sided F-test for the ratio of variances s1^2 / s2^2.
function varRatioF(v1, n1, v2, n2, tail) {
  const df1 = n1 - 1, df2 = n2 - 1, F = v1 / v2;
  const up = fP(F, df1, df2); // P(F >= f)
  let p = tail === "upper" ? up : tail === "lower" ? 1 - up : 2 * Math.min(up, 1 - up);
  if (p > 1) p = 1;
  return { F, df1, df2, p };
}
// Paired comparison: paired t-test on (x1 - x2) vs mu0, plus Pearson r for the pair.
function pairedT(x1v, x2v, mu0, tail) {
  const pairs = [];
  const n = Math.min(x1v.length, x2v.length);
  for (let i = 0; i < n; i++) { const a = num(x1v[i]), b = num(x2v[i]); if (a !== null && b !== null) pairs.push([a, b]); }
  const m = pairs.length;
  if (m < 2) return { n: m };
  const d = describe(pairs.map((p) => p[0] - p[1]));
  const se = d.sd / Math.sqrt(m);
  const df = m - 1;
  const t = (d.mean - (mu0 || 0)) / se;
  const d1 = describe(pairs.map((p) => p[0])), d2 = describe(pairs.map((p) => p[1]));
  let sxx = 0, syy = 0, sxy = 0;
  for (const [a, b] of pairs) { sxx += (a - d1.mean) ** 2; syy += (b - d2.mean) ** 2; sxy += (a - d1.mean) * (b - d2.mean); }
  const r = sxx > 0 && syy > 0 ? sxy / Math.sqrt(sxx * syy) : NaN;
  return { n: m, df, meanDiff: d.mean, sdDiff: d.sd, se, t, p: tP(t, df, tail), d1, d2, r };
}
// Fisher z-transform test of a correlation against a hypothesized rho0.
function corrZTest(r, n, rho0, tail) {
  if (!(n > 3) || !isFinite(r) || Math.abs(r) >= 1) return { n, r, invalid: true };
  const se = 1 / Math.sqrt(n - 3);
  const Z = (Math.atanh(r) - Math.atanh(rho0 || 0)) / se;
  return { n, r, rho0: rho0 || 0, Z, se, p: zP(Z, tail) };
}
// Core two-sample t (pooled or Welch) vs mu0 + variance-ratio F, from two describe()s.
function twoSample(a, b, ka, kb, mu0, tail, welch) {
  if (a.n < 2 || b.n < 2) return { error: "size", keys: [ka, kb] };
  const diff = a.mean - b.mean;
  let df, se;
  if (welch) { const va = a.variance / a.n, vb = b.variance / b.n; se = Math.sqrt(va + vb); df = (va + vb) ** 2 / ((va * va) / (a.n - 1) + (vb * vb) / (b.n - 1)); }
  else { df = a.n + b.n - 2; const sp2 = ((a.n - 1) * a.variance + (b.n - 1) * b.variance) / df; se = Math.sqrt(sp2 * (1 / a.n + 1 / b.n)); }
  const t = (diff - (mu0 || 0)) / se;
  return { groups: [{ key: ka, ...a }, { key: kb, ...b }], meanDiff: diff, df, t, se, p: tP(t, df, tail), welch, F: varRatioF(a.variance, a.n, b.variance, b.n, tail) };
}
// Unpaired comparison from one variable split by a 2-level grouping variable.
function unpairedComp(yv, groupv, mu0, tail, welch) {
  const groups = {};
  yv.forEach((y, i) => { const g = groupv[i], v = num(y); if (g === "" || g === null || g === undefined || v === null) return; (groups[g] = groups[g] || []).push(v); });
  const keys = Object.keys(groups);
  if (keys.length !== 2) return { error: keys.length, keys };
  return twoSample(describe(groups[keys[0]]), describe(groups[keys[1]]), keys[0], keys[1], mu0, tail, welch);
}
// Unpaired comparison from two separate continuous variables (independent samples).
function unpairedCompTwo(a1v, a2v, lab1, lab2, mu0, tail, welch) {
  return twoSample(describe(a1v), describe(a2v), lab1, lab2, mu0, tail, welch);
}

/* ---- nonparametric suite (validated vs SciPy; tie corrections applied) ---- */
const numArr = (v) => v.map(num).filter((x) => x !== null);
function groupArrays(yv, groupv) {
  const map = {}, keys = [];
  yv.forEach((y, i) => { const g = groupv[i], val = num(y); if (g === "" || g === null || g === undefined || val === null) return; if (!(g in map)) { map[g] = []; keys.push(g); } map[g].push(val); });
  return { keys, map };
}
// Average ranks (ties get mean rank); returns ranks aligned to input order + tie-group sizes.
function rankAvg(a) {
  const idx = a.map((v, i) => [v, i]).sort((x, y) => x[0] - y[0]);
  const r = new Array(a.length), ties = []; let i = 0;
  while (i < idx.length) { let j = i; while (j + 1 < idx.length && idx[j + 1][0] === idx[i][0]) j++; const avg = (i + j) / 2 + 1; for (let k = i; k <= j; k++) r[idx[k][1]] = avg; if (j > i) ties.push(j - i + 1); i = j + 1; }
  return { ranks: r, ties };
}
const chiSqSf = (x, df) => (x <= 0 ? 1 : gammq(df / 2, x / 2));
function mannWhitney(a, b, tail, cc) {
  const n1 = a.length, n2 = b.length, N = n1 + n2;
  const o = rankAvg(a.concat(b)), ranks = o.ranks, ties = o.ties;
  const R1 = ranks.slice(0, n1).reduce((s, v) => s + v, 0);
  const U1 = R1 - n1 * (n1 + 1) / 2, U2 = n1 * n2 - U1, mu = n1 * n2 / 2;
  const tieTerm = ties.reduce((s, t) => s + (t * t * t - t), 0);
  const sigma = Math.sqrt(n1 * n2 / 12 * ((N + 1) - tieTerm / (N * (N - 1))));
  let z, p, nm = U1 - mu;
  if (tail === "upper") { z = (nm - (cc ? 0.5 : 0)) / sigma; p = normCdf(-z); }
  else if (tail === "lower") { z = (nm + (cc ? 0.5 : 0)) / sigma; p = normCdf(z); }
  else { const c = cc ? Math.sign(nm) * 0.5 : 0; z = (nm - c) / sigma; p = 2 * normCdf(-Math.abs(z)); }
  return { n1, n2, U1, U2, R1, R2: N * (N + 1) / 2 - R1, z, p, sigma };
}
function wilcoxonSR(d0, tail, cc) {
  const d = d0.filter((v) => v !== 0), n = d.length;
  if (n < 1) return { n: 0 };
  const o = rankAvg(d.map(Math.abs)), ranks = o.ranks, ties = o.ties;
  let rPlus = 0, rMinus = 0; d.forEach((v, i) => { if (v > 0) rPlus += ranks[i]; else rMinus += ranks[i]; });
  const mn = n * (n + 1) / 4, tieTerm = ties.reduce((s, t) => s + (t * t * t - t), 0);
  const se = Math.sqrt(n * (n + 1) * (2 * n + 1) / 24 - tieTerm / 48);
  let T, z, p;
  if (tail === "two") { T = Math.min(rPlus, rMinus); const dcc = cc ? 0.5 * Math.sign(T - mn) : 0; z = (T - mn - dcc) / se; p = 2 * normCdf(-Math.abs(z)); }
  else { T = rPlus; const dcc = cc ? 0.5 * Math.sign(T - mn) : 0; z = (T - mn - dcc) / se; p = tail === "upper" ? normCdf(-z) : normCdf(z); }
  return { n, rPlus, rMinus, T, z, p };
}
const lbinom = (n, k) => gammln(n + 1) - gammln(k + 1) - gammln(n - k + 1);
function binomCdf(k, n, p) { let s = 0; for (let i = 0; i <= k; i++) s += Math.exp(lbinom(n, i) + i * Math.log(p) + (n - i) * Math.log(1 - p)); return Math.min(1, s); }
function signTest(diffs, tail) {
  let np = 0, nmn = 0; diffs.forEach((d) => { if (d > 0) np++; else if (d < 0) nmn++; });
  const n = np + nmn; if (n < 1) return { n: 0 };
  let p;
  if (tail === "upper") p = 1 - binomCdf(np - 1, n, 0.5);
  else if (tail === "lower") p = binomCdf(np, n, 0.5);
  else p = Math.min(1, 2 * binomCdf(Math.min(np, nmn), n, 0.5));
  return { n, nplus: np, nminus: nmn, p };
}
function kruskal(keys, map) {
  const k = keys.length, all = []; keys.forEach((key) => map[key].forEach((v) => all.push(v)));
  const N = all.length, o = rankAvg(all), ranks = o.ranks, ties = o.ties;
  const Rsum = {}; keys.forEach((key) => (Rsum[key] = 0));
  let off = 0; keys.forEach((key) => { for (let i = 0; i < map[key].length; i++) Rsum[key] += ranks[off + i]; off += map[key].length; });
  let H = 12 / (N * (N + 1)) * keys.reduce((s, key) => s + Rsum[key] * Rsum[key] / map[key].length, 0) - 3 * (N + 1);
  const tieTerm = ties.reduce((s, t) => s + (t * t * t - t), 0), corr = 1 - tieTerm / (N * N * N - N);
  if (corr > 0) H /= corr;
  const df = k - 1;
  return { k, N, H, df, p: chiSqSf(H, df), groups: keys.map((key) => ({ key, n: map[key].length, meanRank: Rsum[key] / map[key].length })) };
}
function friedman(blocks) {
  const n = blocks.length, k = blocks[0].length, Rj = new Array(k).fill(0); let tieAdj = 0;
  blocks.forEach((b) => { const o = rankAvg(b); o.ranks.forEach((r, j) => (Rj[j] += r)); tieAdj += o.ties.reduce((s, t) => s + (t * t * t - t), 0); });
  let Q = 12 / (n * k * (k + 1)) * Rj.reduce((s, r) => s + r * r, 0) - 3 * n * (k + 1);
  const corr = 1 - tieAdj / (n * k * (k * k - 1)); if (corr > 0) Q /= corr;
  const df = k - 1;
  return { n, k, Q, df, p: chiSqSf(Q, df), Rj };
}
function spearman(p1, p2, tail) {
  const n = p1.length;
  const r1 = rankAvg(p1).ranks, r2 = rankAvg(p2).ranks;
  const m1 = r1.reduce((s, v) => s + v, 0) / n, m2 = r2.reduce((s, v) => s + v, 0) / n;
  let sxx = 0, syy = 0, sxy = 0;
  for (let i = 0; i < n; i++) { sxx += (r1[i] - m1) ** 2; syy += (r2[i] - m2) ** 2; sxy += (r1[i] - m1) * (r2[i] - m2); }
  const rho = sxy / Math.sqrt(sxx * syy);
  const t = rho * Math.sqrt((n - 2) / ((1 - rho) * (1 + rho))), df = n - 2;
  return { n, rho, t, df, p: tP(t, df, tail) };
}

/* Pearson correlation over a set of complete pairs (numeric arrays), with two-tailed t-test p. */
function pearsonPair(x, y) {
  const n = x.length; if (n < 3) return { n, r: NaN, p: NaN };
  const mx = x.reduce((s, v) => s + v, 0) / n, my = y.reduce((s, v) => s + v, 0) / n;
  let sxx = 0, syy = 0, sxy = 0;
  for (let i = 0; i < n; i++) { sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; sxy += (x[i] - mx) * (y[i] - my); }
  if (sxx <= 0 || syy <= 0) return { n, r: NaN, p: NaN };
  const r = sxy / Math.sqrt(sxx * syy), t = r * Math.sqrt((n - 2) / (1 - r * r));
  return { n, r, p: tP(t, n - 2, "two") };
}
const _variance = (a) => { const m = a.reduce((s, v) => s + v, 0) / a.length; return a.reduce((s, v) => s + (v - m) ** 2, 0) / (a.length - 1); };
/* Cronbach's α on a set of item columns (numeric arrays, listwise-complete). Validated vs pingouin. */
function cronbachAlpha(items) {
  const k = items.length, n = items[0].length;
  const sv = items.reduce((s, col) => s + _variance(col), 0);
  const tot = []; for (let i = 0; i < n; i++) { let s = 0; for (let j = 0; j < k; j++) s += items[j][i]; tot.push(s); }
  const vt = _variance(tot);
  return vt > 0 ? (k / (k - 1)) * (1 - sv / vt) : NaN;
}

/* Fisher's exact test (2×2), two-tailed: sum of hypergeometric probabilities ≤ p(observed). Validated vs scipy. */
function fisher2x2(a, b, c, d) {
  const n = a + b + c + d, r1 = a + b, c1 = a + c, lf = (x) => gammln(x + 1);
  const logp = (x) => lf(r1) + lf(c + d) + lf(c1) + lf(b + d) - lf(n) - lf(x) - lf(r1 - x) - lf(c1 - x) - lf(n - r1 - c1 + x);
  const p0 = Math.exp(logp(a)), lo = Math.max(0, c1 - (c + d)), hi = Math.min(r1, c1);
  let p = 0; for (let x = lo; x <= hi; x++) { const px = Math.exp(logp(x)); if (px <= p0 * (1 + 1e-7)) p += px; }
  return Math.min(1, p);
}
/* Two-way contingency table + Pearson χ², likelihood-ratio G², Yates correction, Cramér's V, φ.
   Validated vs scipy.stats.chi2_contingency / fisher_exact. */
function contingency(rowV, colV) {
  const uniq = (arr) => { const s = [...new Set(arr)], allNum = s.length > 0 && s.every((v) => isFinite(Number(v))); return s.sort(allNum ? (a, b) => Number(a) - Number(b) : (a, b) => String(a).localeCompare(String(b))); };
  const rl = uniq(rowV), cl = uniq(colV), R = rl.length, C = cl.length;
  const ri = Object.fromEntries(rl.map((v, i) => [v, i])), ci = Object.fromEntries(cl.map((v, i) => [v, i]));
  const O = rl.map(() => cl.map(() => 0));
  for (let i = 0; i < rowV.length; i++) O[ri[rowV[i]]][ci[colV[i]]]++;
  const rt = O.map((r) => r.reduce((a, b) => a + b, 0)), ct = cl.map((_, j) => O.reduce((s, r) => s + r[j], 0)), N = rt.reduce((a, b) => a + b, 0);
  const E = rl.map((_, i) => cl.map((_, j) => rt[i] * ct[j] / N));
  let chi2 = 0, g2 = 0, chi2y = 0, minE = Infinity, lowE = 0;
  for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { const o = O[i][j], e = E[i][j]; minE = Math.min(minE, e); if (e < 5) lowE++; if (e > 0) { chi2 += (o - e) ** 2 / e; if (o > 0) g2 += 2 * o * Math.log(o / e); const dd = Math.max(0, Math.abs(o - e) - 0.5); chi2y += dd * dd / e; } }
  const df = (R - 1) * (C - 1), mind = Math.min(R - 1, C - 1);
  const out = { rl, cl, R, C, O, E, rt, ct, N, chi2, g2, chi2y, df, minE, lowE, cells: R * C, p: chiSqSf(chi2, df), pG: chiSqSf(g2, df), pY: chiSqSf(chi2y, df), V: mind > 0 ? Math.sqrt(chi2 / (N * mind)) : NaN, phi: Math.sqrt(chi2 / N) };
  if (R === 2 && C === 2) out.fisher = fisher2x2(O[0][0], O[0][1], O[1][0], O[1][1]);
  return out;
}
// Paired difference vector (x1 - x2 - shift) over complete cases, for paired nonparametric tests.
function pairedDiffs(x1v, x2v, shift) {
  const d = []; const n = Math.min(x1v.length, x2v.length);
  for (let i = 0; i < n; i++) { const a = num(x1v[i]), b = num(x2v[i]); if (a !== null && b !== null) d.push(a - b - (shift || 0)); }
  return d;
}

/* ---- distribution-plot statistics (box / violin) ---- */
// Linear-interpolation quantile (R type 7 / NumPy default) on a sorted array.
function quantileSorted(s, p) {
  const n = s.length; if (n === 0) return NaN; if (n === 1) return s[0];
  const h = (n - 1) * p, lo = Math.floor(h), f = h - lo;
  return s[lo] + (lo + 1 < n ? f * (s[lo + 1] - s[lo]) : 0);
}
function boxStats(vals) {
  const s = [...vals].sort((a, b) => a - b), n = s.length;
  const q1 = quantileSorted(s, 0.25), med = quantileSorted(s, 0.5), q3 = quantileSorted(s, 0.75), iqr = q3 - q1;
  const loF = q1 - 1.5 * iqr, hiF = q3 + 1.5 * iqr;
  let wlo = s[0], whi = s[n - 1];
  for (const v of s) { if (v >= loF) { wlo = v; break; } }
  for (let i = n - 1; i >= 0; i--) { if (s[i] <= hiF) { whi = s[i]; break; } }
  const outliers = s.filter((v) => v < loF || v > hiF);
  const mean = s.reduce((a, b) => a + b, 0) / n;
  return { n, min: s[0], max: s[n - 1], q1, med, q3, iqr, wlo, whi, outliers, mean, sorted: s };
}
const normPdf = (x) => Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI);
// Silverman bandwidth: 0.9 * min(sd, IQR/1.349) * n^(-1/5).
function silverman(vals) {
  const n = vals.length, m = vals.reduce((a, b) => a + b, 0) / n;
  const sd = Math.sqrt(vals.reduce((a, b) => a + (b - m) ** 2, 0) / ((n - 1) || 1));
  const s = [...vals].sort((a, b) => a - b), iqr = quantileSorted(s, 0.75) - quantileSorted(s, 0.25);
  const cand = [sd, iqr > 0 ? iqr / 1.349 : Infinity].filter((x) => x > 0 && isFinite(x));
  const sigma = cand.length ? Math.min(...cand) : 1;
  return 0.9 * sigma * Math.pow(n, -0.2);
}
function kdeDensity(vals, grid, bwMult) {
  let h = silverman(vals) * (bwMult || 1); if (!(h > 0)) h = 1;
  const n = vals.length;
  return grid.map((x) => vals.reduce((acc, xi) => acc + normPdf((x - xi) / h), 0) / (n * h));
}

function histogram(values) {
  const xs = values.map(num).filter((v) => v !== null);
  if (xs.length === 0) return { bins: [] };
  const min = Math.min(...xs), max = Math.max(...xs);
  const k = Math.max(1, Math.ceil(Math.log2(xs.length) + 1));
  const width = (max - min) / k || 1;
  const bins = Array.from({ length: k }, (_, i) => ({ lo: min + i * width, hi: min + (i + 1) * width, count: 0, label: (min + i * width).toFixed(1) }));
  for (const x of xs) { let idx = Math.floor((x - min) / width); if (idx >= k) idx = k - 1; if (idx < 0) idx = 0; bins[idx].count++; }
  return { bins, n: xs.length };
}

/* =========================================================================
   ANOVA ENGINE — balanced mixed factorial / repeated-measures
   (validated to 4+ decimals vs. pingouin / statsmodels)
   ========================================================================= */
const SUBJ = "__SUBJ__";
function subsets(arr) { const out = [[]]; for (const x of arr) { const cur = out.length; for (let i = 0; i < cur; i++) out.push([...out[i], x]); } return out; }
const ekey = (a) => a.join("\u0001");
function effName(E) { return E.length ? E.join(" \u00d7 ") : "(grand)"; }

/* Least-squares GLM for UNBALANCED between-subjects designs: Type III (default) / Type II
   sums of squares via effect (sum-to-zero) coding. Validated vs statsmodels (Sum coding). */
function glmSolve(A, b) {
  const p = b.length, M = A.map((r, i) => [...r, b[i]]);
  for (let col = 0; col < p; col++) {
    let piv = col; for (let r = col + 1; r < p; r++) if (Math.abs(M[r][col]) > Math.abs(M[piv][col])) piv = r;
    if (Math.abs(M[piv][col]) < 1e-9) return null;
    [M[col], M[piv]] = [M[piv], M[col]]; const d = M[col][col];
    for (let r = 0; r < p; r++) { if (r === col) continue; const f = M[r][col] / d; for (let c = col; c <= p; c++) M[r][c] -= f * M[col][c]; }
  }
  return M.map((row, i) => row[p] / row[i]);
}
function glmSSE(cols, y) {
  const p = cols.length, N = y.length;
  const A = Array.from({ length: p }, () => new Array(p).fill(0)), b = new Array(p).fill(0);
  for (let i = 0; i < p; i++) { for (let j = i; j < p; j++) { let s = 0; for (let k = 0; k < N; k++) s += cols[i][k] * cols[j][k]; A[i][j] = s; A[j][i] = s; } let sb = 0; for (let k = 0; k < N; k++) sb += cols[i][k] * y[k]; b[i] = sb; }
  const beta = glmSolve(A, b); if (!beta) return null;
  let yy = 0; for (let k = 0; k < N; k++) yy += y[k] * y[k];
  let bz = 0; for (let i = 0; i < p; i++) bz += beta[i] * b[i];
  return yy - bz;
}
function glmContrast(L, vals, N) { const k = L.length, last = L[k - 1], cols = []; for (let j = 0; j < k - 1; j++) { const Lj = L[j], c = new Array(N); for (let i = 0; i < N; i++) { const v = vals[i]; c[i] = v === Lj ? 1 : v === last ? -1 : 0; } cols.push(c); } return cols; }
function glmEffectCols(E, between, levels, N) { let cols = [new Array(N).fill(1)]; for (const f of E) { const cc = glmContrast(levels[f], between[f], N), nx = []; for (const a of cols) for (const c of cc) { const pr = new Array(N); for (let i = 0; i < N; i++) pr[i] = a[i] * c[i]; nx.push(pr); } cols = nx; } return cols; }
function glmDesign(active, between, levels, N) { const cols = [new Array(N).fill(1)]; for (const E of active) for (const c of glmEffectCols(E, between, levels, N)) cols.push(c); return cols; }
function glmBetween(y, between, levels, bNames, ssType) {
  const N = y.length, allE = subsets(bNames).filter((E) => E.length > 0);
  const sseFull = glmSSE(glmDesign(allE, between, levels, N), y);
  if (sseFull == null) return null; // singular (e.g. empty cells)
  const contains = (U, E) => E.every((f) => U.includes(f)), out = {};
  for (const E of allE) {
    const key = ekey([...E].sort());
    let ss;
    if (ssType === "II") { const M0 = allE.filter((U) => !contains(U, E)); const s0 = glmSSE(glmDesign(M0, between, levels, N), y), s1 = glmSSE(glmDesign([...M0, E], between, levels, N), y); ss = s0 != null && s1 != null ? s0 - s1 : NaN; }
    else { const reduced = allE.filter((U) => U !== E); const sR = glmSSE(glmDesign(reduced, between, levels, N), y); ss = sR != null ? sR - sseFull : NaN; }
    out[key] = { ss };
  }
  return { out, sseFull };
}

/* Sphericity for repeated-measures effects: per-effect orthonormal within-subject
   contrasts (SPSS/afex convention), Mauchly's W test + Greenhouse–Geisser and
   Huynh–Feldt epsilon. Validated vs pingouin (main effects) and independent numpy. */
function sphDot(a, b) { let s = 0; for (let i = 0; i < a.length; i++) s += a[i] * b[i]; return s; }
function sphOrtho(k) { const B = []; for (let j = 1; j < k; j++) { const v = new Array(k).fill(-1 / k); v[j] = 1 - 1 / k; for (const b of B) { const d = sphDot(v, b); for (let i = 0; i < k; i++) v[i] -= d * b[i]; } const n = Math.sqrt(sphDot(v, v)); for (let i = 0; i < k; i++) v[i] /= n; B.push(v); } return B; }
function sphKron(a, b) { const out = []; for (const x of a) for (const y of b) out.push(x * y); return out; }
function sphEffectBasis(activeIdx, wlev) { // orthonormal basis vectors (length C) for the effect's contrast space
  const perF = wlev.map((k, f) => activeIdx.includes(f) ? sphOrtho(k) : [new Array(k).fill(1 / Math.sqrt(k))]);
  let combos = [[]]; perF.forEach((vs) => { const nx = []; for (const c of combos) for (const v of vs) nx.push([...c, v]); combos = nx; });
  return combos.map((vecs) => vecs.reduce((acc, v) => sphKron(acc, v), [1]));
}
function sphDet(M) { const n = M.length, A = M.map((r) => r.slice()); let det = 1; for (let c = 0; c < n; c++) { let p = c; for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r; if (Math.abs(A[p][c]) < 1e-300) return 0; if (p !== c) { const t = A[p]; A[p] = A[c]; A[c] = t; det = -det; } det *= A[c][c]; for (let r = c + 1; r < n; r++) { const f = A[r][c] / A[c][c]; for (let cc = c; cc < n; cc++) A[r][cc] -= f * A[c][cc]; } } return det; }
// V: array of per-subject within-cell vectors (length C); groups: between-cell key per subject; wlev: within level counts; activeIdx: within factor indices in the effect
function sphericityOf(V, groups, wlev, activeIdx) {
  const p = activeIdx.reduce((a, f) => a * (wlev[f] - 1), 1);
  if (p < 2) return { df: p, trivial: true, epsGG: 1, epsHF: 1 };
  const B = sphEffectBasis(activeIdx, wlev);
  const Z = V.map((v) => B.map((b) => sphDot(b, v)));
  const n = Z.length, gset = [...new Set(groups)], ng = gset.length, nu = n - ng;
  if (nu < p) return { df: p, trivial: false, epsGG: NaN, epsHF: NaN, unavailable: true };
  const C = Array.from({ length: p }, () => new Array(p).fill(0));
  for (const gr of gset) { const idx = []; for (let i = 0; i < n; i++) if (groups[i] === gr) idx.push(i); const m = new Array(p).fill(0); idx.forEach((i) => Z[i].forEach((z, j) => m[j] += z)); for (let j = 0; j < p; j++) m[j] /= idx.length; idx.forEach((i) => { for (let a = 0; a < p; a++) for (let b = 0; b < p; b++) C[a][b] += (Z[i][a] - m[a]) * (Z[i][b] - m[b]); }); }
  for (let a = 0; a < p; a++) for (let b = 0; b < p; b++) C[a][b] /= nu;
  let trC = 0; for (let i = 0; i < p; i++) trC += C[i][i];
  let trC2 = 0; for (let a = 0; a < p; a++) for (let b = 0; b < p; b++) trC2 += C[a][b] * C[a][b];
  const epsGG = (trC * trC) / (p * trC2);
  let epsHF = (n * p * epsGG - 2) / (p * (nu - p * epsGG)); epsHF = Math.min(epsHF, 1);
  const det = sphDet(C); let W = NaN, chi2 = NaN, dfChi = p * (p + 1) / 2 - 1, pval = NaN;
  if (det > 0) { W = det / Math.pow(trC / p, p); const corr = 1 - (2 * p * p + p + 2) / (6 * p * nu); chi2 = -nu * corr * Math.log(W); pval = chiSqSf(chi2, dfChi); }
  return { df: p, trivial: false, epsGG, epsHF, W, chi2, dfChi, pMauchly: pval, nu };
}

/* Type III decomposition of a single contrast-score variable on the between-subjects
   design (sum coding): returns the intercept SS (= within-effect contribution), each
   between effect's SS (= effect×between interaction contribution), and the residual
   (= effect×Subject error). Core of the exact unbalanced-mixed split-plot engine. */
function splitGLM(z, between, levels, bNames) {
  const N = z.length, allE = subsets(bNames).filter((E) => E.length > 0);
  const buildCols = (active, withInt) => { const cols = withInt ? [new Array(N).fill(1)] : []; for (const E of active) for (const c of glmEffectCols(E, between, levels, N)) cols.push(c); return cols; };
  const full = glmFit(buildCols(allE, true), z); if (!full) return null;
  const out = {};
  const noInt = glmFit(buildCols(allE, false), z); out.__intercept = (noInt ? noInt.sse : NaN) - full.sse;
  for (const E of allE) { const reduced = allE.filter((U) => U !== E); const red = glmFit(buildCols(reduced, true), z); out[ekey([...E].sort())] = (red ? red.sse : NaN) - full.sse; }
  return { out, sseFull: full.sse };
}

function anova(input) {
  const y = input.y, N = y.length;
  const between = input.between || {}, within = input.within || {};
  const bNames = Object.keys(between), wNames = Object.keys(within);
  const subj = input.subject;
  const dimVal = (dim, i) => dim === SUBJ ? subj[i] : (within[dim] ? within[dim][i] : between[dim][i]);

  const levels = {};
  for (const f of [...bNames, ...wNames]) { const s = new Set(); for (let i = 0; i < N; i++) s.add(dimVal(f, i)); levels[f] = [...s]; }
  const dfFactor = (f) => levels[f].length - 1;
  const dfEffect = (E) => E.reduce((p, f) => p * dfFactor(f), 1);

  const bracketCache = new Map();
  function bracket(dims) {
    const ck = ekey([...dims].sort());
    if (bracketCache.has(ck)) return bracketCache.get(ck);
    const g = new Map();
    for (let i = 0; i < N; i++) { const k = ekey(dims.map((d) => String(dimVal(d, i)))); const e = g.get(k) || [0, 0]; e[0] += y[i]; e[1] += 1; g.set(k, e); }
    let s = 0; for (const [, [sum, n]] of g) s += sum * sum / n;
    bracketCache.set(ck, s); return s;
  }
  function ssEffect(dims) { let s = 0; for (const T of subsets(dims)) { const sign = ((dims.length - T.length) % 2 === 0) ? 1 : -1; s += sign * bracket(T); } return s; }

  const messages = [];
  const nWithinCells = wNames.reduce((p, f) => p * levels[f].length, 1);
  const perSubj = new Map(); for (let i = 0; i < N; i++) perSubj.set(subj[i], (perSubj.get(subj[i]) || 0) + 1);
  const subjects = [...perSubj.keys()], Nsubj = subjects.length;
  let balanced = true; for (const c of perSubj.values()) if (c !== nWithinCells) balanced = false;
  const subjGroup = new Map(); for (let i = 0; i < N; i++) subjGroup.set(subj[i], ekey(bNames.map((b) => String(between[b][i]))));
  const subjPerGroup = new Map(); for (const s of subjects) { const gk = subjGroup.get(s); subjPerGroup.set(gk, (subjPerGroup.get(gk) || 0) + 1); }
  const nBetweenCells = subjPerGroup.size;
  const pg = [...subjPerGroup.values()];
  const betweenUnbalanced = pg.some((v) => v !== pg[0]);
  if (betweenUnbalanced) balanced = false;

  const CF = bracket([]);
  let raw = 0; for (let i = 0; i < N; i++) raw += y[i] * y[i];
  const ssTotal = raw - CF, dfTotal = N - 1;

  const ssSB = bracket([SUBJ]) - bracket(bNames);
  const dfSB = Nsubj - nBetweenCells;
  const msSB = dfSB > 0 ? ssSB / dfSB : NaN;

  const wErr = new Map();
  for (const Ew of subsets(wNames)) {
    if (Ew.length === 0) continue;
    const pooled = ssEffect([...Ew, SUBJ]);
    let sub = 0;
    for (const Eb of subsets(bNames)) { if (Eb.length === 0) continue; sub += ssEffect([...Ew, ...Eb]); }
    const ss = pooled - sub, df = dfEffect(Ew) * (Nsubj - nBetweenCells);
    wErr.set(ekey([...Ew].sort()), { ss, df, ms: df > 0 ? ss / df : NaN });
  }

  // ---- sphericity prep: per-subject within-cell vectors (canonical order, first within factor outermost) ----
  const wlev = wNames.map((f) => levels[f].length);
  let spherV = null, spherGroups = null, spherOK = false;
  if (wNames.length) {
    const cellIndex = (i) => { let idx = 0; for (const f of wNames) idx = idx * levels[f].length + levels[f].indexOf(dimVal(f, i)); return idx; };
    const Vmap = new Map();
    for (let i = 0; i < N; i++) { const s = subj[i]; if (!Vmap.has(s)) Vmap.set(s, new Array(nWithinCells).fill(null)); Vmap.get(s)[cellIndex(i)] = y[i]; }
    spherOK = dfSB > 0 && subjects.every((s) => { const v = Vmap.get(s); return v && v.every((x) => x !== null); });
    if (spherOK) { spherV = subjects.map((s) => Vmap.get(s)); spherGroups = subjects.map((s) => subjGroup.get(s)); }
  }
  const spherList = [];

  // ---- exact unbalanced-mixed (complete within): within-contrast scores + Type III between-GLM ----
  let useSP = false, spEffSS = null, spSubjGroups = null, spWErr = null;
  if (wNames.length && bNames.length && betweenUnbalanced && spherOK) {
    const seen = new Map(); for (let i = 0; i < N; i++) if (!seen.has(subj[i])) seen.set(subj[i], i);
    const betweenObj = {}; bNames.forEach((b) => betweenObj[b] = subjects.map((s) => String(between[b][seen.get(s)])));
    spEffSS = new Map(); spWErr = new Map(); let ok = true;
    const grandVec = sphEffectBasis([], wlev)[0];
    const gsp = splitGLM(spherV.map((v) => sphDot(grandVec, v)), betweenObj, levels, bNames);
    if (gsp) {
      spSubjGroups = { ss: gsp.sseFull, df: Nsubj - nBetweenCells };
      for (const Eb of subsets(bNames)) { if (Eb.length) spEffSS.set(ekey([...Eb].sort()), gsp.out[ekey([...Eb].sort())]); }
      for (const Ew of subsets(wNames)) {
        if (!Ew.length) continue;
        const activeIdx = Ew.map((f) => wNames.indexOf(f)), vecs = sphEffectBasis(activeIdx, wlev);
        let interceptSS = 0, errSS = 0; const interMap = new Map();
        for (const cv of vecs) {
          const sp = splitGLM(spherV.map((v) => sphDot(cv, v)), betweenObj, levels, bNames);
          if (!sp) { ok = false; break; }
          interceptSS += sp.out.__intercept; errSS += sp.sseFull;
          for (const Eb of subsets(bNames)) { if (Eb.length) { const k = ekey([...Eb].sort()); interMap.set(k, (interMap.get(k) || 0) + sp.out[k]); } }
        }
        if (!ok) break;
        const wk = ekey([...Ew].sort());
        spEffSS.set(wk, interceptSS);
        for (const [bk, ss] of interMap) spEffSS.set(ekey([...wk.split("\u0001"), ...bk.split("\u0001")].sort()), ss);
        spWErr.set(wk, { ss: errSS, df: dfEffect(Ew) * (Nsubj - nBetweenCells), ms: errSS / (dfEffect(Ew) * (Nsubj - nBetweenCells)) });
      }
      useSP = ok;
    }
  }

  const sources = [];
  const allFixed = subsets([...bNames, ...wNames]).filter((E) => E.length > 0);
  const wPart = (E) => E.filter((f) => wNames.includes(f));

  const ssType = input.ssType === "II" ? "II" : "III";
  let glmUsed = false;
  let glm = null;
  if (wNames.length === 0 && betweenUnbalanced) glm = glmBetween(y, between, levels, bNames, ssType);
  glmUsed = !!glm;
  const sbMS = useSP ? spSubjGroups.ss / spSubjGroups.df : msSB, sbDF = useSP ? spSubjGroups.df : dfSB, sbSS = useSP ? spSubjGroups.ss : ssSB;
  const pureBetween = allFixed.filter((E) => wPart(E).length === 0).sort((a, b) => a.length - b.length);
  for (const E of pureBetween) {
    const df = dfEffect(E), key = ekey([...E].sort());
    const ss = useSP ? spEffSS.get(key) : (glm ? glm.out[key].ss : ssEffect(E));
    const ms = ss / df, F = ms / sbMS, p = fP(F, df, sbDF);
    sources.push({ name: effName(E), type: "between", df, ss, ms, F, p, errMS: sbMS, errDF: sbDF, errSS: sbSS });
  }
  if (glmUsed) messages.push(`Unbalanced between-subjects design — Type ${ssType} sums of squares (least-squares GLM, marginal/unweighted-means hypotheses).`);
  else if (useSP) messages.push("Unbalanced mixed design — exact Type III via within-subject contrasts + between-subjects least-squares GLM (complete within-cell data).");
  else if (betweenUnbalanced) messages.push("Unbalanced between-subjects design in a mixed model — F-tests are approximate (Type III GLM not yet applied to mixed designs).");
  else if (!balanced) messages.push("Unequal observations per subject (design is unbalanced; F-tests approximate).");
  if (bNames.length || wNames.length) sources.push({ name: wNames.length ? "Subject (Groups)" : "Residual", type: "error", df: sbDF, ss: sbSS, ms: sbMS });

  const withinEffects = allFixed.filter((E) => wPart(E).length > 0);
  const blocks = new Map();
  for (const E of withinEffects) { const wk = ekey(wPart(E).sort()); if (!blocks.has(wk)) blocks.set(wk, []); blocks.get(wk).push(E); }
  const blockKeys = [...blocks.keys()].sort((a, b) => a.split("\u0001").length - b.split("\u0001").length || a.localeCompare(b));
  for (const wk of blockKeys) {
    const Es = blocks.get(wk).sort((a, b) => a.length - b.length);
    const err = useSP ? spWErr.get(wk) : wErr.get(wk);
    const wkFactors = wk.split("\u0001");
    let sph = null;
    if (spherOK) {
      const activeIdx = wkFactors.map((f) => wNames.indexOf(f));
      sph = sphericityOf(spherV, spherGroups, wlev, activeIdx);
      spherList.push({ block: effName(wkFactors), df: sph.df, trivial: !!sph.trivial, unavailable: !!sph.unavailable, W: sph.W, chi2: sph.chi2, dfChi: sph.dfChi, pMauchly: sph.pMauchly, epsGG: sph.epsGG, epsHF: sph.epsHF });
    }
    for (const E of Es) {
      const ss = useSP ? spEffSS.get(ekey([...E].sort())) : ssEffect(E), df = dfEffect(E), ms = ss / df, F = ms / err.ms, p = fP(F, df, err.df);
      const s = { name: effName(E), type: "within", df, ss, ms, F, p, errMS: err.ms, errDF: err.df, errSS: err.ss };
      if (sph && !sph.unavailable) {
        s.epsGG = sph.epsGG; s.epsHF = sph.epsHF;
        if (!sph.trivial && isFinite(sph.epsGG)) { s.pGG = fP(F, df * sph.epsGG, err.df * sph.epsGG); s.pHF = fP(F, df * sph.epsHF, err.df * sph.epsHF); }
      }
      sources.push(s);
    }
    sources.push({ name: effName(wkFactors) + " \u00d7 Subj", type: "error", df: err.df, ss: err.ss, ms: err.ms });
  }

  return { sources, total: { df: dfTotal, ss: ssTotal }, balanced, messages, Nsubj, nWithinCells, nBetweenCells, betweenUnbalanced, glmUsed, splitPlot: useSP, ssType, sphericity: spherList, spherOK };
}

/* =========================================================================
   COMPACT-VARIABLE MODEL
   compact = { id, name, factors:[{name,levels:[]}], leaves:[colId,...] }
   factors ordered OUTERMOST -> innermost; leaves in odometer order with the
   INNERMOST factor varying fastest (VibeStat's column-compaction scheme).
   ========================================================================= */
function leafLevels(compact, idx) {
  const out = {};
  let rem = idx;
  for (let j = compact.factors.length - 1; j >= 0; j--) {
    const f = compact.factors[j], L = f.levels.length;
    out[f.name] = f.levels[rem % L];
    rem = Math.floor(rem / L);
  }
  return out;
}
const compactCells = (c) => c.factors.reduce((p, f) => p * Math.max(1, f.levels.length), 1);

function compactToLong(compact, betweenSpecs, rows) {
  const subject = [], y = [];
  const within = {}; compact.factors.forEach((f) => (within[f.name] = []));
  const between = {}; betweenSpecs.forEach((b) => (between[b.name] = []));
  rows.forEach((r, ri) => {
    compact.leaves.forEach((colId, li) => {
      const v = num(r[colId]);
      if (v === null) return;
      subject.push("S" + ri); y.push(v);
      const lv = leafLevels(compact, li);
      compact.factors.forEach((f) => within[f.name].push(lv[f.name]));
      betweenSpecs.forEach((b) => between[b.name].push(r[b.id]));
    });
  });
  return { subject, between, within, y };
}

/* ---- t critical value for a two-tailed alpha (bisection on the t cdf) ---- */
function tCrit(alpha, df) {
  if (df <= 0) return NaN;
  let lo = 0, hi = 1000;
  for (let k = 0; k < 80; k++) { const mid = (lo + hi) / 2; (tTwoTailedP(mid, df) > alpha) ? (lo = mid) : (hi = mid); }
  return (lo + hi) / 2;
}
const cartesian = (arrs) => arrs.length ? arrs.reduce((a, b) => a.flatMap((x) => b.map((y) => [...x, y])), [[]]) : [[]];

/* All non-empty subsets of `factors`, ordered by size then position. Each is an "effect". */
function effectList(factors, mode) {
  if (factors.length === 0) return [];
  if (mode === "highest") return [[...factors]];
  const out = [];
  const rec = (start, acc) => { for (let i = start; i < factors.length; i++) { const next = [...acc, factors[i]]; out.push(next); rec(i + 1, next); } };
  rec(0, []);
  return out.sort((a, b) => a.length - b.length);
}

/* Cell means for an effect: aggregate to one value per subject per cell, then summarize across subjects. */
function cellMeans(long, effectFactors, alpha, errType) {
  const N = long.y.length;
  const dimVal = (name, i) => (long.within[name] ? long.within[name][i] : long.between[name][i]);
  const subjCell = new Map();
  for (let i = 0; i < N; i++) {
    const lvl = effectFactors.map((f) => String(dimVal(f, i))).join("\u0001");
    const sk = long.subject[i] + "\u0002" + lvl;
    const e = subjCell.get(sk) || [0, 0]; e[0] += long.y[i]; e[1] += 1; subjCell.set(sk, e);
  }
  const byCell = new Map();
  for (const [sk, [sum, n]] of subjCell) {
    const lvl = sk.split("\u0002")[1];
    if (!byCell.has(lvl)) byCell.set(lvl, []);
    byCell.get(lvl).push(sum / n);
  }
  const rows = [];
  for (const [lvl, arr] of byCell) {
    const d = describe(arr);
    const se = d.sem;
    let err = 0;
    if (errType === "se") err = se; else if (errType === "sd") err = d.sd; else if (errType === "ci") err = tCrit(alpha, d.n - 1) * se;
    rows.push({ levels: lvl.split("\u0001"), mean: d.mean, sd: d.sd, se, n: d.n, err });
  }
  return rows;
}

/* ---- studentized range distribution (for Tukey HSD); validated vs. published q tables ---- */
function _phiN(x) { return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI); }
function _PhiN(x) { const s = x < 0 ? -1 : 1, ax = Math.abs(x) / Math.SQRT2, p = 0.3275911, a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429; const t = 1 / (1 + p * ax); const er = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-ax * ax); return 0.5 * (1 + s * er); }
function _prange(w, k) { if (w <= 0) return 0; const N = 120, lo = -9, hi = 9, h = (hi - lo) / N; let s = 0; for (let i = 0; i <= N; i++) { const u = lo + i * h; const f = _phiN(u) * Math.pow(_PhiN(u) - _PhiN(u - w), k - 1); s += (i === 0 || i === N ? 1 : (i % 2 ? 4 : 2)) * f; } return Math.min(1, k * s * h / 3); }
function ptukey(q, k, nu) { if (q <= 0) return 0; if (!isFinite(nu) || nu > 2000) return _prange(q, k); const half = nu / 2, logZ = half * Math.log(0.5) - gammln(half); const fchi = (u) => u <= 0 ? 0 : Math.exp(logZ + (half - 1) * Math.log(u) - 0.5 * u); const N = 160, lo = 1e-7, hi = nu + 14 * Math.sqrt(2 * nu) + 50, h = (hi - lo) / N; let s = 0; for (let i = 0; i <= N; i++) { const u = lo + i * h; const f = _prange(q * Math.sqrt(u / nu), k) * fchi(u); s += (i === 0 || i === N ? 1 : (i % 2 ? 4 : 2)) * f; } return Math.min(1, s * h / 3); }

const PH_METHODS = { lsd: "Fisher's LSD", bonf: "Bonferroni", holm: "Holm", sidak: "Šidák", tukey: "Tukey HSD", scheffe: "Scheffé", bh: "Benjamini–Hochberg (FDR)" };

/* Pairwise post-hoc comparisons among the CELLS of an effect (a main effect or any
   interaction). All-within or all-between effects use a single pooled error term
   (df matching the ANOVA stratum) so Tukey HSD & Scheffé are exact. Mixed within×between
   effects have no single valid pooled error, so each pair is tested directly — paired t
   when the two cells share the same between level(s), two-sample t otherwise — and
   Tukey/Scheffé are flagged not-applicable. Then the chosen multiplicity adjustment. */
function posthoc(long, effectFactors, withinNames, betweenNames, levelOrder, method) {
  const N = long.y.length;
  const dv = (name, i) => (long.within[name] ? long.within[name][i] : long.between[name][i]);
  const betweenInE = effectFactors.filter((f) => betweenNames.includes(f));
  const withinInE = effectFactors.filter((f) => withinNames.includes(f));
  const mixed = betweenInE.length > 0 && withinInE.length > 0;
  const allBetween = withinInE.length === 0;
  const cells = cartesian(effectFactors.map((f) => levelOrder[f] || []));
  const K = cells.length, cn = (arr) => arr.join("\u0001");
  const subj = new Map();
  for (let i = 0; i < N; i++) {
    const s = long.subject[i];
    if (!subj.has(s)) subj.set(s, { bfull: betweenNames.map((b) => String(long.between[b][i])).join("|"), bInE: betweenInE.map((b) => String(long.between[b][i])).join("\u0001"), cv: new Map(), all: [0, 0] });
    const o = subj.get(s), c = effectFactors.map((f) => String(dv(f, i))).join("\u0001");
    const e = o.cv.get(c) || [0, 0]; e[0] += long.y[i]; e[1] += 1; o.cv.set(c, e);
    o.all[0] += long.y[i]; o.all[1] += 1;
  }
  const subjects = [...subj.keys()], nSubj = subjects.length;
  const bFulls = [...new Set(subjects.map((s) => subj.get(s).bfull))], nB = bFulls.length;
  const cellMean = {}, nPer = {}, pairs = [];
  let commonDf = null, msErr = null, family;
  if (!mixed) {
    family = "pooled";
    if (!allBetween) { // all-within composite factor
      const X = {}; subjects.forEach((s) => { X[s] = {}; cells.forEach((c) => { const e = subj.get(s).cv.get(cn(c)) || [0, 0]; X[s][cn(c)] = e[1] ? e[0] / e[1] : NaN; }); });
      cells.forEach((c) => { cellMean[cn(c)] = subjects.reduce((a, s) => a + X[s][cn(c)], 0) / nSubj; nPer[cn(c)] = nSubj; });
      const grand = cells.reduce((a, c) => a + cellMean[cn(c)], 0) / K;
      const sm = {}; subjects.forEach((s) => sm[s] = cells.reduce((a, c) => a + X[s][cn(c)], 0) / K);
      let ssFS = 0; subjects.forEach((s) => cells.forEach((c) => { const d = X[s][cn(c)] - sm[s] - cellMean[cn(c)] + grand; ssFS += d * d; }));
      let ssFB = 0; bFulls.forEach((g) => { const gs = subjects.filter((s) => subj.get(s).bfull === g), ng = gs.length; const gM = {}; cells.forEach((c) => gM[cn(c)] = gs.reduce((a, s) => a + X[s][cn(c)], 0) / ng); const gG = cells.reduce((a, c) => a + gM[cn(c)], 0) / K; cells.forEach((c) => { const d = gM[cn(c)] - gG - cellMean[cn(c)] + grand; ssFB += ng * d * d; }); });
      commonDf = (K - 1) * (nSubj - nB); msErr = commonDf > 0 ? (ssFS - ssFB) / commonDf : NaN;
    } else { // all-between composite factor
      const xv = {}; subjects.forEach((s) => { const al = subj.get(s).all; xv[s] = al[0] / al[1]; });
      const byCell = {}; subjects.forEach((s) => { (byCell[subj.get(s).bInE] = byCell[subj.get(s).bInE] || []).push(s); });
      cells.forEach((c) => { const arr = byCell[cn(c)] || []; cellMean[cn(c)] = arr.length ? arr.reduce((a, s) => a + xv[s], 0) / arr.length : NaN; nPer[cn(c)] = arr.length; });
      const cf = {}; subjects.forEach((s) => { (cf[subj.get(s).bfull] = cf[subj.get(s).bfull] || []).push(s); });
      let ssW = 0; Object.values(cf).forEach((arr) => { const m = arr.reduce((a, s) => a + xv[s], 0) / arr.length; arr.forEach((s) => ssW += (xv[s] - m) ** 2); });
      commonDf = nSubj - nB; msErr = commonDf > 0 ? ssW / commonDf : NaN;
    }
    for (let i = 0; i < K; i++) for (let j = i + 1; j < K; j++) { const a = cn(cells[i]), b = cn(cells[j]); const se = Math.sqrt(msErr * (1 / nPer[a] + 1 / nPer[b])); const diff = cellMean[a] - cellMean[b]; pairs.push({ a: cells[i], b: cells[j], diff, se, t: se > 0 ? diff / se : 0, df: commonDf }); }
  } else { // mixed: per-comparison
    family = "percontrast";
    const X = {}; subjects.forEach((s) => { X[s] = {}; subj.get(s).cv.forEach((e, c) => X[s][c] = e[0] / e[1]); });
    cells.forEach((c) => { const who = subjects.filter((s) => X[s][cn(c)] != null); cellMean[cn(c)] = who.length ? who.reduce((a, s) => a + X[s][cn(c)], 0) / who.length : NaN; nPer[cn(c)] = who.length; });
    const bproj = (c) => betweenInE.map((f) => String(c[effectFactors.indexOf(f)])).join("\u0001");
    for (let i = 0; i < K; i++) for (let j = i + 1; j < K; j++) {
      const ci = cells[i], cj = cells[j], a = cn(ci), b = cn(cj); let se, t, df;
      if (bproj(ci) === bproj(cj)) {
        const who = subjects.filter((s) => subj.get(s).bInE === bproj(ci) && X[s][a] != null && X[s][b] != null);
        const d = who.map((s) => X[s][a] - X[s][b]), n = d.length, md = d.reduce((x, y) => x + y, 0) / n;
        const sd = Math.sqrt(d.reduce((x, y) => x + (y - md) ** 2, 0) / (n - 1)); se = sd / Math.sqrt(n); df = n - 1; t = se > 0 ? md / se : 0;
      } else {
        const ga = subjects.filter((s) => subj.get(s).bInE === bproj(ci) && X[s][a] != null), gb = subjects.filter((s) => subj.get(s).bInE === bproj(cj) && X[s][b] != null);
        const va = ga.map((s) => X[s][a]), vb = gb.map((s) => X[s][b]); const ma = va.reduce((x, y) => x + y, 0) / va.length, mb = vb.reduce((x, y) => x + y, 0) / vb.length;
        const ssa = va.reduce((x, y) => x + (y - ma) ** 2, 0), ssb = vb.reduce((x, y) => x + (y - mb) ** 2, 0); df = va.length + vb.length - 2; const sp2 = (ssa + ssb) / df; se = Math.sqrt(sp2 * (1 / va.length + 1 / vb.length)); t = se > 0 ? (ma - mb) / se : 0;
      }
      pairs.push({ a: ci, b: cj, diff: cellMean[a] - cellMean[b], se, t, df });
    }
  }
  pairs.forEach((p) => p.rawp = tTwoTailedP(p.t, p.df));
  const m = pairs.length, ord = pairs.map((_, i) => i).sort((x, y) => pairs[x].rawp - pairs[y].rawp);
  const pooled = family === "pooled";
  pairs.forEach((p) => {
    if (method === "lsd") p.padj = p.rawp;
    else if (method === "bonf") p.padj = Math.min(1, m * p.rawp);
    else if (method === "sidak") p.padj = 1 - Math.pow(1 - p.rawp, m);
    else if (method === "scheffe") p.padj = pooled ? fP(p.t * p.t / (K - 1), K - 1, commonDf) : p.rawp;
    else if (method === "tukey") p.padj = pooled ? 1 - ptukey(Math.SQRT2 * Math.abs(p.t), K, commonDf) : p.rawp;
  });
  if (method === "holm") { let mx = 0; ord.forEach((ii, r) => { const v = Math.min(1, (m - r) * pairs[ii].rawp); mx = Math.max(mx, v); pairs[ii].padj = mx; }); }
  if (method === "bh") { let mn = 1; for (let r = m - 1; r >= 0; r--) { const ii = ord[r]; const v = Math.min(1, (m / (r + 1)) * pairs[ii].rawp); mn = Math.min(mn, v); pairs[ii].padj = mn; } }
  return { pairs, K, m, commonDf, msErr, mixed, family, fwerNA: (mixed && (method === "tukey" || method === "scheffe")) };
}

/* =========================================================================
   ANALYSIS REGISTRY
   ========================================================================= */
const CONTINUOUS = ["real", "integer", "formula", "formula_static"];
const isContinuous = (t) => CONTINUOUS.includes(t);

const ANALYSES = {
  descriptive: { name: "Descriptive Statistics", category: "Descriptive Statistics", roles: [{ key: "y", label: "Variable", multiple: true, accept: "continuous" }] },
  histogram: { name: "Histogram", category: "Frequency Distribution", roles: [{ key: "x", label: "X Variable", accept: "continuous" }] },
  boxplot: { name: "Box Plot", category: "Distribution", roles: [{ key: "y", label: "Variable", accept: "continuous" }, { key: "group", label: "Grouping (optional)", accept: "categorical", optional: true }] },
  violin: { name: "Violin Plot", category: "Distribution", roles: [{ key: "y", label: "Variable", accept: "continuous" }, { key: "group", label: "Grouping (optional)", accept: "categorical", optional: true }] },
  regression: { name: "Simple Regression", category: "Regression", roles: [{ key: "x", label: "X (independent)", accept: "continuous" }, { key: "y", label: "Y (dependent)", accept: "continuous" }] },
  mreg: { name: "Multiple Regression", category: "Regression", roles: [{ key: "y", label: "Dependent (Y)", accept: "continuous" }, { key: "x", label: "Predictors (X)", multiple: true, accept: "continuous" }] },
  glm: { name: "General Linear Model", category: "Regression", roles: [{ key: "y", label: "Dependent (Y)", accept: "continuous" }, { key: "factors", label: "Factors (categorical)", multiple: true, accept: "categorical", optional: true }, { key: "covs", label: "Covariates (continuous)", multiple: true, accept: "continuous", optional: true }] },
  normtest: { name: "Normality (Shapiro–Wilk)", category: "Diagnostics", roles: [{ key: "y", label: "Variable", accept: "continuous" }, { key: "group", label: "Group (optional)", accept: "categorical", optional: true }] },
  homovar: { name: "Homogeneity of Variance", category: "Diagnostics", roles: [{ key: "y", label: "Variable", accept: "continuous" }, { key: "group", label: "Group", accept: "categorical" }] },
  contrast: { name: "Planned Contrasts / Trend", category: "Comparison", roles: [{ key: "y", label: "Dependent (Y)", accept: "continuous" }, { key: "factor", label: "Factor (ordered)", accept: "categorical" }] },
  wcontrast: { name: "Within-Subjects Contrasts / Trend", category: "Comparison", roles: [{ key: "dep", label: "Repeated variable (R)", accept: "depOrCompact" }] },
  corrmatrix: { name: "Correlation Matrix", category: "Correlation", roles: [{ key: "vars", label: "Variables", multiple: true, accept: "continuous" }] },
  reliability: { name: "Reliability (Cronbach's α)", category: "Correlation", roles: [{ key: "items", label: "Scale items", multiple: true, accept: "continuous" }] },
  corrviz: { name: "Correlation Heatmap / Scatter Matrix", category: "Correlation", roles: [{ key: "vars", label: "Variables", multiple: true, accept: "continuous" }] },
  profile: { name: "Profile Plot (repeated measures)", category: "Comparison", roles: [{ key: "dep", label: "Repeated variable (R)", accept: "depOrCompact" }, { key: "between", label: "Grouping factor (optional)", multiple: true, optional: true, accept: "categorical" }] },
  crosstab: { name: "Contingency Table / Chi-Square", category: "Tables", roles: [{ key: "row", label: "Row variable", accept: "categorical" }, { key: "col", label: "Column variable", accept: "categorical" }] },
  piechart: { name: "Pie Chart", category: "Frequency Distribution", roles: [{ key: "var", label: "Variable", accept: "categorical" }] },
  barchart: { name: "Bar Chart", category: "Frequency Distribution", roles: [{ key: "var", label: "Variable", accept: "categorical" }, { key: "group", label: "Grouping (optional)", accept: "categorical", optional: true }] },
  forest: { name: "Forest Plot (effect sizes)", category: "Comparison", roles: [{ key: "dep", label: "Dependent", accept: "continuous" }, { key: "factor", label: "Factor", accept: "categorical" }] },
  scattergram: { name: "Scattergram", category: "Regression", roles: [{ key: "x", label: "X", accept: "continuous" }, { key: "y", label: "Y", accept: "continuous" }] },
  ttest: { name: "Unpaired Comparison", category: "Comparison", roles: [{ key: "y", label: "Test Variable", accept: "continuous" }, { key: "group", label: "Grouping (2 levels)", accept: "categorical" }] },
  paired: { name: "Paired Comparison", category: "Comparison", roles: [{ key: "x1", label: "Variable 1", accept: "continuous" }, { key: "x2", label: "Variable 2", accept: "continuous" }] },
  mannwhitney: { name: "Mann–Whitney U", category: "Nonparametric", roles: [{ key: "y", label: "Test Variable", accept: "continuous" }, { key: "group", label: "Grouping (2 levels)", accept: "categorical" }] },
  wilcoxon: { name: "Wilcoxon Signed-Rank", category: "Nonparametric", roles: [{ key: "x1", label: "Variable 1", accept: "continuous" }, { key: "x2", label: "Variable 2", accept: "continuous" }] },
  signtest: { name: "Sign Test", category: "Nonparametric", roles: [{ key: "x1", label: "Variable 1", accept: "continuous" }, { key: "x2", label: "Variable 2", accept: "continuous" }] },
  kruskal: { name: "Kruskal–Wallis", category: "Nonparametric", roles: [{ key: "y", label: "Test Variable", accept: "continuous" }, { key: "group", label: "Grouping (≥2 levels)", accept: "categorical" }] },
  friedman: { name: "Friedman", category: "Nonparametric", roles: [{ key: "dep", label: "Repeated-measures variable", accept: "compactOnly" }] },
  spearman: { name: "Spearman Rank Correlation", category: "Nonparametric", roles: [{ key: "x1", label: "Variable 1", accept: "continuous" }, { key: "x2", label: "Variable 2", accept: "continuous" }] },
  anova: { name: "ANOVA", category: "ANOVA", roles: [{ key: "dep", label: "Dependent / Repeated", accept: "depOrCompact" }, { key: "between", label: "Between Factor(s)", multiple: true, accept: "categorical" }] },
  importnotes: { name: "Prism Import", category: "Descriptive Statistics", roles: [] },
  power: { name: "Power Analysis", category: "Power Analysis", roles: [] },
};
const CATEGORY_ORDER = ["Descriptive Statistics", "Frequency Distribution", "Distribution", "Diagnostics", "Regression", "Correlation", "Tables", "Comparison", "Nonparametric", "ANOVA", "Power Analysis"];

const DEFAULT_ANOVA_CFG = { design: "repeated", alpha: 0.05, effects: "highest", errorBars: "se" };
const DEFAULT_CMP_OPTS = { tail: "two", mu0: 0, rho0: 0, showCorr: true, varAssume: "pooled", showF: true, srcMode: "grouped", cc: true };
const tailLabel = (tail) => (tail === "upper" ? "P (1-tail, >)" : tail === "lower" ? "P (1-tail, <)" : "P (2-tail)");
// Which option controls each comparison/nonparametric test exposes.
const CMP_CAP = {
  unpaired: { tail: 1, mu0: 1, srcMode: 1, variance: 1 },
  paired: { tail: 1, mu0: 1, corr: 1 },
  mannwhitney: { tail: 1, srcMode: 1, cc: 1 },
  wilcoxon: { tail: 1, mu0: 1, cc: 1 },
  signtest: { tail: 1, mu0: 1 },
  spearman: { tail: 1 },
};
// Comparison-style types that render their own options strip + role hints (so they
// bypass the generic "assign roles" gate in renderResult).
const CMP_TYPES = new Set(["ttest", "paired", "mannwhitney", "wilcoxon", "signtest", "spearman"]);
const DEFAULT_DIST_OPTS = { showPoints: false, showMean: true, showBox: true, bw: 1 };
const DIST_TYPES = new Set(["boxplot", "violin"]);
// Effective roles for an analysis. Unpaired Comparison and Mann–Whitney can take either a
// test variable + grouping variable ("grouped") or two separate continuous variables ("twoCol").
function rolesOf(a) {
  if ((a.type === "ttest" || a.type === "mannwhitney") && a.opts && a.opts.srcMode === "twoCol")
    return [{ key: "x1", label: "Variable 1", accept: "continuous" }, { key: "x2", label: "Variable 2", accept: "continuous" }];
  return ANALYSES[a.type].roles;
}
const ANOVA_OUTPUTS = { table: "ANOVA Table", means: "Means Table", bar: "Interaction Bar Graph", line: "Interaction Line Graph", posthoc: "Post-hoc Comparisons" };

// Items shown in the Analysis Browser (several ANOVA outputs map to one analysis type).
const BROWSER_ITEMS = [
  { cat: "Descriptive Statistics", label: "Descriptive Statistics", type: "descriptive" },
  { cat: "Frequency Distribution", label: "Histogram", type: "histogram" },
  { cat: "Distribution", label: "Box Plot", type: "boxplot" },
  { cat: "Distribution", label: "Violin Plot", type: "violin" },
  { cat: "Regression", label: "Simple Regression", type: "regression" },
  { cat: "Regression", label: "Multiple Regression", type: "mreg" },
  { cat: "Regression", label: "General Linear Model / ANCOVA", type: "glm" },
  { cat: "Diagnostics", label: "Normality (Shapiro–Wilk)", type: "normtest" },
  { cat: "Diagnostics", label: "Homogeneity of Variance", type: "homovar" },
  { cat: "Comparison", label: "Planned Contrasts / Trend", type: "contrast" },
  { cat: "Comparison", label: "Within-Subjects Contrasts / Trend", type: "wcontrast" },
  { cat: "Correlation", label: "Correlation Matrix", type: "corrmatrix" },
  { cat: "Correlation", label: "Reliability (Cronbach's α)", type: "reliability" },
  { cat: "Correlation", label: "Correlation Heatmap / Scatter Matrix", type: "corrviz" },
  { cat: "Comparison", label: "Profile Plot (repeated measures)", type: "profile" },
  { cat: "Tables", label: "Contingency Table / Chi-Square", type: "crosstab" },
  { cat: "Frequency Distribution", label: "Pie Chart", type: "piechart" },
  { cat: "Frequency Distribution", label: "Bar Chart", type: "barchart" },
  { cat: "Comparison", label: "Forest Plot (effect sizes)", type: "forest" },
  { cat: "Regression", label: "Scattergram", type: "scattergram" },
  { cat: "Comparison", label: "Paired Comparison", type: "paired" },
  { cat: "Comparison", label: "Unpaired Comparison", type: "ttest" },
  { cat: "Nonparametric", label: "Mann–Whitney U", type: "mannwhitney" },
  { cat: "Nonparametric", label: "Wilcoxon Signed-Rank", type: "wilcoxon" },
  { cat: "Nonparametric", label: "Sign Test", type: "signtest" },
  { cat: "Nonparametric", label: "Kruskal–Wallis", type: "kruskal" },
  { cat: "Nonparametric", label: "Friedman", type: "friedman" },
  { cat: "Nonparametric", label: "Spearman Rank Correlation", type: "spearman" },
  { cat: "ANOVA", label: "ANOVA Table", type: "anova", output: "table" },
  { cat: "ANOVA", label: "Means Table", type: "anova", output: "means" },
  { cat: "ANOVA", label: "Interaction Bar Graph", type: "anova", output: "bar" },
  { cat: "ANOVA", label: "Interaction Line Graph", type: "anova", output: "line" },
  { cat: "ANOVA", label: "Post-hoc Comparisons", nested: Object.entries(PH_METHODS).map(([k, v]) => ({ label: v, type: "anova", output: "posthoc", method: k })) },
  { cat: "Power Analysis", label: "Power Analysis", type: "power" },
];
const PALETTE = ["#4a6fa5", "#c0653a", "#5a9367", "#9a5ba6", "#b8993a", "#5aa0a8"];

/* ---- editable-plot configuration ---- */
const DEFAULT_PLOT = {
  title: "", xLabel: "", yLabel: "",
  yMin: "", yMax: "", yTickInterval: "", // "" = auto
  xInterval: 0,                          // 0 = show every x label; n = skip n between
  xAngle: 0,                             // x-axis label angle: 0 (horizontal) | 45 | 90
  axisWidth: 1, tickWidth: 1,            // axis line / hash-mark line thickness
  grid: true, gridH: true, gridV: false, gridDashed: true,
  frame: false,                          // box around the plot area
  width: "", height: "",                 // explicit size in px ("" = auto-fit); also set by dragging
  legend: "top",                         // top | tl | tr | bl | br | right | none
  legendXY: null,                        // {x,y} in SVG units — free-dragged legend position (null = anchored by `legend`)
  labelSize: 10, labelBold: false, labelItalic: false,   // axis titles (X/Y label)
  tickSize: 9,                                           // axis tick numbers
  legendSize: 9, legendBold: false, legendItalic: false, // legend text
  series: {},                            // { [seriesKey]: { color, symbol, symbolColor, lineWidth } }
};
const SYMBOLS = ["circle", "square", "triangle", "diamond", "none"];
function defaultSeriesStyle(i) { const c = PALETTE[i % PALETTE.length]; return { color: c, symbol: "circle", symbolColor: c, symbolSize: 3.5, lineWidth: 2 }; }
function seriesStyleOf(plot, key, i) { return { ...defaultSeriesStyle(i), ...((plot.series && plot.series[key]) || {}) }; }
// custom dot renderer so symbol shape is editable (recharts dot accepts a render fn)
function plotDot(shape, fill, r) {
  if (shape === "none") return false;
  const rad = r || 3;
  return (props) => {
    const { cx, cy } = props; const key = props.key || `${props.dataKey}-${props.index}`;
    if (cx == null || cy == null) return null;
    if (shape === "square") return <rect key={key} x={cx - rad} y={cy - rad} width={2 * rad} height={2 * rad} fill={fill} />;
    if (shape === "diamond") return <path key={key} d={`M${cx} ${cy - rad}L${cx + rad} ${cy}L${cx} ${cy + rad}L${cx - rad} ${cy}Z`} fill={fill} />;
    if (shape === "triangle") return <path key={key} d={`M${cx} ${cy - rad}L${cx + rad} ${cy + rad}L${cx - rad} ${cy + rad}Z`} fill={fill} />;
    return <circle key={key} cx={cx} cy={cy} r={rad} fill={fill} />;
  };
}
// resolve Y domain + explicit ticks from the plot config (and data when an interval needs bounds)
function yAxisProps(data, keys, showErr, plot) {
  const num = (v) => (v !== "" && v != null && Number.isFinite(+v) ? +v : null);
  const yMin = num(plot.yMin), yMax = num(plot.yMax), step = num(plot.yTickInterval) > 0 ? +plot.yTickInterval : null;
  let lo = yMin, hi = yMax;
  if (step && (lo == null || hi == null)) {
    let mn = Infinity, mx = -Infinity;
    data.forEach((p) => keys.forEach((k) => { const v = p[k]; if (v == null || Number.isNaN(v)) return; const e = showErr ? (p[k + "__e"] || 0) : 0; mn = Math.min(mn, v - e); mx = Math.max(mx, v + e); }));
    if (Number.isFinite(mn)) { if (lo == null) lo = Math.floor(mn / step) * step; if (hi == null) hi = Math.ceil(mx / step) * step; }
  }
  const domain = [lo == null ? "auto" : lo, hi == null ? "auto" : hi];
  if (step && typeof lo === "number" && typeof hi === "number") {
    const ticks = []; for (let t = lo; t <= hi + 1e-9; t += step) ticks.push(+t.toFixed(6));
    return { domain, ticks };
  }
  return { domain, ticks: undefined };
}


/* =========================================================================
   SAMPLE DATA — a 1-between (Sex) x 3-within (Drug x Time x Trial[11]) mixed RM
   study: 2 x 2 x 11 = 44 repeated measurements per subject.
   Generated deterministically so the dataset is stable across reloads.
   ========================================================================= */
const RT_FACTORS = [
  { name: "Drug", levels: ["placebo", "active"] },
  { name: "Time", levels: ["pre", "post"] },
  { name: "Trial", levels: Array.from({ length: 11 }, (_, i) => String(i + 1)) },
];
const _RTls = RT_FACTORS.map((f) => f.levels.length);
const _RTn = _RTls.reduce((a, b) => a * b, 1); // 44
const _rtLevels = (idx) => { let rem = idx; const o = {}; for (let j = RT_FACTORS.length - 1; j >= 0; j--) { const L = _RTls[j]; o[RT_FACTORS[j].name] = RT_FACTORS[j].levels[rem % L]; rem = Math.floor(rem / L); } return o; };
const RT_LEAVES = Array.from({ length: _RTn }, (_, i) => "y" + i); // odometer order, Trial varies fastest
const _RT_LEAF_COLS = RT_LEAVES.map((id, i) => { const lv = _rtLevels(i); return { id, name: `${lv.Drug === "placebo" ? "pl" : "ac"}/${lv.Time}/T${lv.Trial}`, type: "real" }; });

const SAMPLE_COLS = [
  { id: "subj", name: "Subject", type: "string" },
  { id: "sex", name: "Sex", type: "category" },
  ..._RT_LEAF_COLS,
];

const _mulberry = (seed) => () => { seed |= 0; seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
const _rand = _mulberry(20260603);
const _gauss = () => { let u = 0, v = 0; while (!u) u = _rand(); while (!v) v = _rand(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); };
const _SUBJECTS = [
  { subj: "S01", sex: "F" }, { subj: "S02", sex: "F" }, { subj: "S03", sex: "F" }, { subj: "S04", sex: "F" },
  { subj: "S05", sex: "M" }, { subj: "S06", sex: "M" }, { subj: "S07", sex: "M" }, { subj: "S08", sex: "M" },
];
const SAMPLE_ROWS = _SUBJECTS.map((s) => {
  const row = { subj: s.subj, sex: s.sex };
  const subjIntercept = _gauss() * 16 + (s.sex === "M" ? 42 : 0);   // between-subject spread + Sex effect
  for (let i = 0; i < _RTn; i++) {
    const lv = _rtLevels(i);
    const drugEff = lv.Drug === "active" ? -32 : 0;
    const timeEff = lv.Time === "post" ? -14 : 0;
    const trial = Number(lv.Trial);
    const trialEff = -3.6 * (trial - 1);                            // learning: RT falls across trials
    const drugTime = (lv.Drug === "active" && lv.Time === "post") ? -11 : 0;
    const mu = 480 + subjIntercept + drugEff + timeEff + trialEff + drugTime + _gauss() * 6;
    row[RT_LEAVES[i]] = Math.round(mu * 10) / 10;
  }
  return row;
});

const SAMPLE_COMPACTS = [
  { id: "rt", name: "RT (ms)", factors: RT_FACTORS, leaves: RT_LEAVES },
];
const SAMPLE_ANALYSES = [
  { id: "a1", type: "anova", output: "table", cfg: { ...DEFAULT_ANOVA_CFG }, roles: { dep: { kind: "compact", id: "rt" }, between: ["sex"] } },
  { id: "a2", type: "anova", output: "line", swap: true, cfg: { ...DEFAULT_ANOVA_CFG, errorBars: "se" }, roles: { dep: { kind: "compact", id: "rt" }, between: ["sex"] } },
];
// Full sample project (the study we have been developing against). The app launches
// blank; this loads on demand via File > Open Sample Dataset, and is also written to
// the kit as samples/mixed-RM-study.vibestat.json for opening through the file dialog.
const SAMPLE_PROJECT = { app: "VibeStat", version: 4, columns: SAMPLE_COLS, rows: SAMPLE_ROWS, compacts: SAMPLE_COMPACTS, analyses: SAMPLE_ANALYSES, selAnalysis: "a1", excluded: [], colW: {} };
// Blank slate the app opens with: a few empty columns ready for typing or import.
const BLANK_COLS = [{ id: "c1", name: "Column 1", type: "real" }];
const BLANK_ROWS = Array.from({ length: 16 }, () => ({}));

/* =========================================================================
   STYLING — Mac OS 9 "Platinum"
   ========================================================================= */
const FONT = '"Geneva","Helvetica Neue",Verdana,sans-serif';
const PLAT = { face: "#dcdcdc", faceLite: "#ececec", light: "#ffffff", dark: "#888888", darker: "#555555", text: "#000", sel: "#b8c8e8", selBorder: "#3a5a9a" };
const CANVAS_DEFAULT = { color: "#7c8a99", pattern: "pinstripe" };
function canvasStyle(bg) {
  const color = (bg && bg.color) || CANVAS_DEFAULT.color;
  const pat = (bg && bg.pattern) || CANVAS_DEFAULT.pattern;
  const pats = {
    none: { backgroundImage: "none" },
    pinstripe: { backgroundImage: "repeating-linear-gradient(0deg,rgba(255,255,255,0.06),rgba(255,255,255,0.06) 1px,transparent 1px,transparent 2px)" },
    dots: { backgroundImage: "radial-gradient(rgba(255,255,255,0.16) 1.2px, transparent 1.3px)", backgroundSize: "9px 9px" },
    grid: { backgroundImage: "repeating-linear-gradient(0deg,rgba(255,255,255,0.07),rgba(255,255,255,0.07) 1px,transparent 1px,transparent 9px),repeating-linear-gradient(90deg,rgba(255,255,255,0.07),rgba(255,255,255,0.07) 1px,transparent 1px,transparent 9px)" },
  };
  return { background: color, ...(pats[pat] || pats.pinstripe) };
}
function SettingsDialog({ bg, setBg, onClose }) {
  const cur = bg || CANVAS_DEFAULT;
  const colors = ["#7c8a99", "#6f7f73", "#5b7d8a", "#5a6a9a", "#7a6f8f", "#8a7a5f", "#7d7d7d", "#9aa3ad", "#506070", "#2f3a40"];
  const pats = [["none", "None"], ["pinstripe", "Pinstripe"], ["dots", "Dots"], ["grid", "Grid"]];
  return (
    <div onMouseDown={onClose} style={{ position: "absolute", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.22)", zIndex: 4000, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onMouseDown={(e) => e.stopPropagation()} style={{ width: 330, background: PLAT.face, ...bevelOut, boxShadow: "2px 2px 0 rgba(0,0,0,0.3)", fontFamily: FONT, color: PLAT.text }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 19, borderBottom: `1px solid ${PLAT.dark}`, position: "relative" }}>
          <div onMouseDown={(e) => { e.stopPropagation(); onClose(); }} title="close" style={{ position: "absolute", left: 6, top: 4, width: 11, height: 11, ...bevelOut, background: PLAT.faceLite, cursor: "pointer" }} />
          <span style={{ background: PLAT.face, padding: "0 8px", fontWeight: "bold" }}>Settings</span>
        </div>
        <div style={{ padding: 12 }}>
          <div style={{ fontWeight: "bold", marginBottom: 5 }}>Canvas color</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 11, alignItems: "center" }}>
            {colors.map((c) => <div key={c} onClick={() => setBg({ ...cur, color: c })} title={c} style={{ width: 22, height: 22, background: c, cursor: "pointer", boxSizing: "border-box", border: cur.color === c ? `2px solid ${PLAT.selBorder}` : "1px solid #555" }} />)}
            <label title="custom color" style={{ display: "inline-flex" }}><input type="color" value={cur.color} onChange={(e) => setBg({ ...cur, color: e.target.value })} style={{ width: 26, height: 24, padding: 0, border: "1px solid #555", cursor: "pointer", background: "none" }} /></label>
          </div>
          <div style={{ fontWeight: "bold", marginBottom: 5 }}>Pattern</div>
          <div style={{ display: "flex", gap: 5, marginBottom: 12 }}>
            {pats.map(([k, lbl]) => <span key={k} onClick={() => setBg({ ...cur, pattern: k })} style={{ cursor: "pointer", fontSize: 11, padding: "2px 9px", border: `1px solid ${PLAT.dark}`, ...bevelOut, background: cur.pattern === k ? PLAT.sel : PLAT.faceLite }}>{lbl}</span>)}
          </div>
          <div title="preview" style={{ height: 56, ...canvasStyle(cur), border: `1px solid ${PLAT.darker}`, marginBottom: 12 }} />
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <MacBtn onClick={() => setBg({ ...CANVAS_DEFAULT })}>Reset</MacBtn>
            <MacBtn primary onClick={onClose}>Done</MacBtn>
          </div>
        </div>
      </div>
    </div>
  );
}
const bevelOut = { borderTop: `1px solid ${PLAT.light}`, borderLeft: `1px solid ${PLAT.light}`, borderRight: `1px solid ${PLAT.darker}`, borderBottom: `1px solid ${PLAT.darker}` };
const bevelIn = { borderTop: `1px solid ${PLAT.darker}`, borderLeft: `1px solid ${PLAT.darker}`, borderRight: `1px solid ${PLAT.light}`, borderBottom: `1px solid ${PLAT.light}` };
const STRIPES = "repeating-linear-gradient(0deg,#d6d6d6,#d6d6d6 1px,#f2f2f2 1px,#f2f2f2 2px)";
const COMPACT_CLR = "#c63";

/* =========================================================================
   DRAGGABLE WINDOW
   ========================================================================= */
function MacWindow({ title, pos, setPos, z, onFocus, onClose, width, height, children, active }) {
  const drag = useRef(null);
  const rz = useRef(null);
  const [collapsed, setCollapsed] = useState(false);
  const onDown = (e) => { onFocus(); drag.current = { sx: e.clientX, sy: e.clientY, ox: pos.x, oy: pos.y }; e.preventDefault(); };
  const onRzDown = (e) => { onFocus(); rz.current = { sx: e.clientX, sy: e.clientY, ow: width, oh: height }; e.preventDefault(); e.stopPropagation(); };
  useEffect(() => {
    const move = (e) => {
      if (drag.current) setPos({ x: drag.current.ox + e.clientX - drag.current.sx, y: Math.max(24, drag.current.oy + e.clientY - drag.current.sy) });
      if (rz.current) setPos({ w: Math.max(220, rz.current.ow + e.clientX - rz.current.sx), h: Math.max(90, rz.current.oh + e.clientY - rz.current.sy) });
    };
    const up = () => { drag.current = null; rz.current = null; };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
    return () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); };
  }, [setPos]);
  return (
    <div onMouseDown={onFocus} style={{ position: "absolute", left: pos.x, top: pos.y, width, zIndex: z, background: PLAT.face, ...bevelOut, boxShadow: "2px 2px 0 rgba(0,0,0,0.28)", fontFamily: FONT, color: PLAT.text }}>
      <div onMouseDown={onDown} onDoubleClick={(e) => { e.stopPropagation(); setCollapsed((c) => !c); }} style={{ height: 18, display: "flex", alignItems: "center", padding: "0 4px", cursor: "default", borderBottom: `1px solid ${PLAT.dark}`, background: active ? STRIPES : PLAT.face }}>
        <div onMouseDown={(e) => { e.stopPropagation(); onClose && onClose(); }} title="close" style={{ width: 11, height: 11, ...bevelOut, background: PLAT.faceLite, cursor: "pointer", marginRight: 6, flexShrink: 0 }} />
        <div style={{ flex: 1, textAlign: "center", fontSize: 11, fontWeight: "bold", whiteSpace: "nowrap", overflow: "hidden", background: active ? STRIPES : "transparent" }}>
          <span style={{ background: active ? PLAT.face : "transparent", padding: "0 8px" }}>{title}</span>
        </div>
        <div onMouseDown={(e) => { e.stopPropagation(); onFocus(); setCollapsed((c) => !c); }} title={collapsed ? "expand" : "collapse (window shade)"} style={{ width: 11, height: 11, marginLeft: 6, ...bevelOut, background: PLAT.faceLite, cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ width: 6, height: 1, background: PLAT.dark }} />
        </div>
      </div>
      {!collapsed && (
        <div style={{ height, display: "flex", flexDirection: "column", overflow: "hidden", position: "relative" }}>
          {children}
          <div onMouseDown={onRzDown} title="drag to resize" style={{ position: "absolute", right: 0, bottom: 0, width: 15, height: 15, cursor: "nwse-resize", zIndex: 5, background: "linear-gradient(135deg, transparent 0 45%, " + PLAT.darker + " 45% 55%, transparent 55% 70%, " + PLAT.darker + " 70% 80%, transparent 80%)" }} />
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   MAIN APP
   ========================================================================= */
export default function VibeStat() {
  const [columns, setColumns] = useState(BLANK_COLS);
  const [rows, setRows] = useState(BLANK_ROWS);
  const [compacts, setCompacts] = useState([]);
  const [analyses, setAnalyses] = useState([]);
  const [selAnalysis, setSelAnalysis] = useState(null);
  const [sel, setSel] = useState(null); // {kind:'col'|'compact', id}
  const [selRow, setSelRow] = useState(null); // highlighted case (row index) or null
  const [colSel, setColSel] = useState(() => new Set()); // multi-selected column ids (for simultaneous resize)
  const [docName, setDocName] = useState(null); // null => window titled "Dataset"; otherwise the saved/opened file name
  const delVarRef = useRef(null);
  const pasteRef = useRef(null);
  const pasteAnchorRef = useRef(null);
  const cellRangeRef = useRef(null); // current rectangular cell selection in the grid (for copy/cut/paste)
  const [excluded, setExcluded] = useState(() => new Set()); // excluded case (row) indices
  const [colW, setColW] = useState({}); // data-grid column id -> pixel width (resizable)
  const [dlg, setDlg] = useState(null); // compact dialog state
  const [anovaDlg, setAnovaDlg] = useState(null); // ANOVA setup popup: {id, cfg}
  const [plotDlg, setPlotDlg] = useState(null); // plot settings popup: {id, plot, seriesKeys}
  const [expanded, setExpanded] = useState({});
  const [openMenu, setOpenMenu] = useState(null);
  const [aboutOpen, setAboutOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [canvasBg, setCanvasBg] = useState({ ...CANVAS_DEFAULT });
  const [pzfxPick, setPzfxPick] = useState(null); // {projects, titles, typeDescs, shapes, activeIdx}
  const [windows, setWindows] = useState({
    data: { x: 16, y: 34, w: 580, h: 300, open: true }, browser: { x: 612, y: 34, w: 230, h: 250, open: true },
    vars: { x: 612, y: 368, w: 230, h: 344, open: true }, view: { x: 16, y: 368, w: 580, h: 470, open: true },
  });
  const [zStack, setZStack] = useState(["browser", "vars", "data", "view"]);

  /* ---------- undo / redo: debounced snapshots of the document ---------- */
  const histRef = useRef(null);
  const travelingRef = useRef(false);
  const pendingRef = useRef(null);
  const lastKeyRef = useRef("");
  const [, setHistTick] = useState(0);
  const bumpHist = () => setHistTick((t) => t + 1);
  const buildDoc = () => ({ columns, rows, compacts, analyses, selAnalysis, excluded: [...excluded], colW });
  const docKey = (d) => JSON.stringify([d.columns, d.rows, d.compacts, d.analyses, d.excluded, d.colW]);
  if (histRef.current === null) { const d0 = buildDoc(); histRef.current = { stack: [d0], idx: 0 }; lastKeyRef.current = docKey(d0); }
  const resetHistory = (d) => { histRef.current = { stack: [d], idx: 0 }; lastKeyRef.current = docKey(d); if (pendingRef.current) { clearTimeout(pendingRef.current); pendingRef.current = null; } bumpHist(); };
  const applyDoc = (d) => {
    travelingRef.current = true;
    setColumns(d.columns); setRows(d.rows); setCompacts(d.compacts); setAnalyses(d.analyses); setColW(d.colW || {});
    setExcluded(new Set(d.excluded || []));
    setSelAnalysis(d.selAnalysis && d.analyses.some((a) => a.id === d.selAnalysis) ? d.selAnalysis : (d.analyses[0] ? d.analyses[0].id : null));
    setSel(null); setSelRow(null); setColSel(new Set());
    lastKeyRef.current = docKey(d);
  };
  const pushSnapshot = (d, key) => {
    const h = histRef.current; let stack = h.stack.slice(0, h.idx + 1); stack.push(d);
    if (stack.length > 100) stack = stack.slice(stack.length - 100);
    histRef.current = { stack, idx: stack.length - 1 }; lastKeyRef.current = key;
  };
  const flushHistory = () => {
    if (!pendingRef.current) return;
    clearTimeout(pendingRef.current); pendingRef.current = null;
    const d = buildDoc(), key = docKey(d);
    if (key !== lastKeyRef.current) pushSnapshot(d, key);
  };
  const undo = () => { flushHistory(); const h = histRef.current; if (h.idx <= 0) return; const ni = h.idx - 1; histRef.current = { stack: h.stack, idx: ni }; applyDoc(h.stack[ni]); bumpHist(); };
  const redo = () => { const h = histRef.current; if (h.idx >= h.stack.length - 1) return; const ni = h.idx + 1; histRef.current = { stack: h.stack, idx: ni }; applyDoc(h.stack[ni]); bumpHist(); };
  const canUndo = histRef.current.idx > 0;
  const canRedo = histRef.current.idx < histRef.current.stack.length - 1;
  const undoRef = useRef(undo); undoRef.current = undo;
  const redoRef = useRef(redo); redoRef.current = redo;
  useEffect(() => {
    if (travelingRef.current) { travelingRef.current = false; return; }
    const d = buildDoc(), key = docKey(d);
    if (key === lastKeyRef.current) return;
    if (pendingRef.current) clearTimeout(pendingRef.current);
    pendingRef.current = setTimeout(() => { pendingRef.current = null; pushSnapshot(d, key); bumpHist(); }, 450);
    // selAnalysis intentionally excluded from deps (navigation, not a document edit)
  }, [columns, rows, compacts, analyses, excluded, colW]);
  useEffect(() => {
    const onKey = (e) => {
      if (!(e.metaKey || e.ctrlKey)) return;
      const k = (e.key || "").toLowerCase();
      if (k === "z" && !e.shiftKey) { e.preventDefault(); undoRef.current(); }
      else if ((k === "z" && e.shiftKey) || k === "y") { e.preventDefault(); redoRef.current(); }
    };
    window.addEventListener("keydown", onKey);
    return () => { window.removeEventListener("keydown", onKey); if (pendingRef.current) clearTimeout(pendingRef.current); };
  }, []);
  // Delete key removes the selected variable (unless typing in a field).
  useEffect(() => {
    const onDel = (e) => {
      if (e.key !== "Delete" && e.key !== "Backspace") return;
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.isContentEditable)) return;
      if (!sel) return;
      e.preventDefault(); if (delVarRef.current) delVarRef.current();
    };
    window.addEventListener("keydown", onDel);
    return () => window.removeEventListener("keydown", onDel);
  }, [sel]);
  useEffect(() => {
    const grab = () => {
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA") && ae.selectionStart != null && ae.selectionStart !== ae.selectionEnd) return null;
      const cr = cellRangeRef.current;
      if (cr && cr.text != null && cr.cells && cr.cells.length) return { text: cr.text, clearCells: cr.cells };
      let cols = [];
      if (colSel && colSel.size) cols = columns.filter((c) => colSel.has(c.id));
      else if (sel && sel.kind === "col") { const c = columns.find((x) => x.id === sel.id); if (c) cols = [c]; }
      if (!cols.length) return null;
      const ids = cols.map((c) => c.id);
      const head = cols.map((c) => c.name).join("\t");
      const body = rows.map((r) => ids.map((id) => (r[id] == null ? "" : r[id])).join("\t")).join("\n");
      return { ids, text: head + "\n" + body };
    };
    const onCopy = (e) => { const g = grab(); if (!g) return; e.preventDefault(); (e.clipboardData || window.clipboardData).setData("text/plain", g.text); };
    const onCut = (e) => {
      const g = grab(); if (!g) return; e.preventDefault(); (e.clipboardData || window.clipboardData).setData("text/plain", g.text);
      if (g.clearCells) { const byRow = {}; g.clearCells.forEach((cc) => { (byRow[cc.ri] = byRow[cc.ri] || []).push(cc.cid); }); setRows((rs) => rs.map((r, i) => { if (!byRow[i]) return r; const n = { ...r }; byRow[i].forEach((cid) => (n[cid] = "")); return n; })); }
      else if (g.ids) { const idset = new Set(g.ids); setRows((rs) => rs.map((r) => { const n = { ...r }; idset.forEach((id) => (n[id] = "")); return n; })); }
    };
    const onPaste = (e) => {
      if (e.defaultPrevented) return;
      const cd = e.clipboardData || window.clipboardData; const t = cd ? cd.getData("text") : "";
      if (!t) return;
      const a = pasteAnchorRef.current; if (!a) return;
      const ae = document.activeElement; const editing = ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA");
      if (!/[\t\n]/.test(t) && editing) return; // let the focused cell take a single value
      e.preventDefault(); if (pasteRef.current) pasteRef.current(a.ri, a.cid, t);
    };
    window.addEventListener("copy", onCopy); window.addEventListener("cut", onCut); window.addEventListener("paste", onPaste);
    return () => { window.removeEventListener("copy", onCopy); window.removeEventListener("cut", onCut); window.removeEventListener("paste", onPaste); };
  }, [colSel, sel, columns, rows]);

  const colById = useMemo(() => Object.fromEntries(columns.map((c) => [c.id, c])), [columns]);
  const compactById = useMemo(() => Object.fromEntries(compacts.map((c) => [c.id, c])), [compacts]);
  const leafToCompact = useMemo(() => { const m = {}; compacts.forEach((c) => c.leaves.forEach((id) => (m[id] = c.id))); return m; }, [compacts]);
  const dynCols = useMemo(() => columns.filter((c) => c.type === "formula"), [columns]);
  const computeRow = useCallback((r) => {
    if (!dynCols.length) return r;
    const nr = { ...r };
    dynCols.forEach((c) => { nr[c.id] = evalFormula(c.formula, nr, columns, compacts); });
    return nr;
  }, [dynCols, columns, compacts]);
  const computedRows = useMemo(() => rows.map(computeRow), [rows, computeRow]);
  const activeRows = useMemo(() => computedRows.filter((_, i) => !excluded.has(i)), [computedRows, excluded]);
  const valuesOf = useCallback((id) => activeRows.map((r) => r[id]), [activeRows]);
  const toggleExcluded = useCallback((ri) => setExcluded((s) => { const n = new Set(s); n.has(ri) ? n.delete(ri) : n.add(ri); return n; }), []);
  const setColFormula = (cid, formula) => setColumns((c) => c.map((x) => (x.id === cid ? { ...x, formula } : x)));
  const recomputeStatic = useCallback((cid) => {
    const col = columns.find((c) => c.id === cid); if (!col) return;
    setRows((rs) => rs.map((r) => ({ ...r, [cid]: evalFormula(col.formula, computeRow(r), columns, compacts) })));
  }, [columns, compacts, computeRow]);

  const focus = (key) => setZStack((s) => [...s.filter((k) => k !== key), key]);
  const zOf = (key) => (key === "vars" ? 100 : 0) + 10 + zStack.indexOf(key); // Variables browser floats on top
  const setPos = (key) => (p) => setWindows((w) => ({ ...w, [key]: { ...w[key], ...p } }));
  const isActive = (key) => zStack[zStack.length - 1] === key;
  const toggleWin = (key) => setWindows((w) => ({ ...w, [key]: { ...w[key], open: !w[key].open } }));

  /* ---- data editing ---- */
  const setCell = (ri, cid, val) => setRows((rs) => rs.map((r, i) => (i === ri ? { ...r, [cid]: val } : r)));
  // Paste tab/newline-delimited data (e.g. copied from Excel) starting at one cell,
  // growing columns and rows as needed.
  const pasteAt = (startRi, startCid, text) => {
    const matrix = String(text).replace(/\r/g, "").split("\n").map((ln) => ln.split("\t"));
    while (matrix.length > 1 && matrix[matrix.length - 1].length === 1 && matrix[matrix.length - 1][0] === "") matrix.pop();
    if (!matrix.length) return;
    const w = matrix.reduce((m, row) => Math.max(m, row.length), 0);
    const cols = columns.slice();
    const start = cols.findIndex((c) => c.id === startCid);
    if (start < 0) return;
    while (cols.length < start + w) cols.push({ id: "v" + Math.random().toString(36).slice(2, 8), name: "Column " + (cols.length + 1), type: "real" });
    const ids = []; for (let j = 0; j < w; j++) ids.push(cols[start + j].id);
    const rws = rows.map((r) => ({ ...r }));
    while (rws.length < startRi + matrix.length) rws.push({});
    matrix.forEach((mr, i) => mr.forEach((v, j) => { if (ids[j] != null) rws[startRi + i][ids[j]] = v; }));
    setColumns(cols); setRows(rws);
  };
  pasteRef.current = pasteAt;
  const addCase = () => setRows((rs) => [...rs, Object.fromEntries(columns.map((c) => [c.id, ""]))]);
  const addVar = () => {
    const id = "v" + Math.random().toString(36).slice(2, 7);
    setColumns((c) => [...c, { id, name: "New Var", type: "real" }]);
    setRows((rs) => rs.map((r) => ({ ...r, [id]: "" })));
  };
  const setColType = (cid, type) => setColumns((c) => c.map((x) => (x.id === cid ? { ...x, type } : x)));
  const setColName = (cid, name) => setColumns((c) => c.map((x) => (x.id === cid ? { ...x, name } : x)));
  const setColDecimals = (cid, d) => setColumns((cs) => {
    const targets = colSel && colSel.has(cid) && colSel.size > 1 ? colSel : new Set([cid]);
    return cs.map((x) => (targets.has(x.id) ? { ...x, decimals: d } : x));
  });
  const setCompactDecimals = (cpId, d) => { const cp = compactById[cpId]; if (!cp) return; const leaves = new Set(cp.leaves); setColumns((cs) => cs.map((x) => (leaves.has(x.id) ? { ...x, decimals: d } : x))); };
  const addFormulaVar = () => {
    const id = "v" + Math.random().toString(36).slice(2, 7);
    setColumns((c) => [...c, { id, name: "Formula", type: "formula", formula: "" }]);
    setRows((rs) => rs.map((r) => ({ ...r, [id]: "" })));
    setTimeout(() => openFormulaBuilder(id), 0);
  };
  const scrubColFromAnalyses = (cid) => setAnalyses((arr) => arr.map((a) => {
    if (!a.roles) return a; const roles = { ...a.roles }; let ch = false;
    ["dep", "x", "y"].forEach((k) => { const r = roles[k]; if (r && (r.id === cid || r === cid)) { delete roles[k]; ch = true; } });
    if (Array.isArray(roles.between)) { const nb = roles.between.filter((b) => (b && b.id ? b.id : b) !== cid); if (nb.length !== roles.between.length) { roles.between = nb; ch = true; } }
    return ch ? { ...a, roles } : a;
  }));
  const deleteSelectedVar = () => {
    if (!sel) return;
    if (sel.kind === "col") {
      const cid = sel.id; const c = colById[cid];
      if (!window.confirm(`Delete variable “${c ? c.name : cid}”? This cannot be undone.`)) return;
      scrubColFromAnalyses(cid);
      setColumns((cs) => cs.filter((x) => x.id !== cid));
      setRows((rs) => rs.map((r) => { const n = { ...r }; delete n[cid]; return n; }));
      setColW((w) => { const n = { ...w }; delete n[cid]; return n; });
      setCompacts((cs) => cs.map((cp) => ({ ...cp, leaves: cp.leaves.filter((l) => l !== cid) })).filter((cp) => cp.leaves.length >= 1));
    } else {
      const cp = compactById[sel.id]; if (!cp) return;
      if (!window.confirm(`Delete compact variable “${cp.name}” and its ${cp.leaves.length} columns? This cannot be undone.`)) return;
      const leaves = new Set(cp.leaves); detachAnalyses(sel.id);
      setColumns((cs) => cs.filter((x) => !leaves.has(x.id)));
      setRows((rs) => rs.map((r) => { const n = { ...r }; cp.leaves.forEach((l) => delete n[l]); return n; }));
      setCompacts((cs) => cs.filter((x) => x.id !== sel.id));
    }
    setSel(null); setColSel(new Set());
  };
  delVarRef.current = deleteSelectedVar;
  const includeAllCases = () => setExcluded(new Set());
  const insertColumn = () => {
    if (!sel) return;
    const id = "v" + Math.random().toString(36).slice(2, 7);
    const newCol = { id, name: "New Var", type: "real" };
    let afterId;
    if (sel.kind === "compact") { const cp = compactById[sel.id]; afterId = cp ? cp.leaves[cp.leaves.length - 1] : null; }
    else { const cpId = leafToCompact[sel.id]; afterId = cpId && compactById[cpId] ? compactById[cpId].leaves[compactById[cpId].leaves.length - 1] : sel.id; }
    setColumns((cs) => { const i = cs.findIndex((c) => c.id === afterId); const n = [...cs]; n.splice(i < 0 ? cs.length : i + 1, 0, newCol); return n; });
    setRows((rs) => rs.map((r) => ({ ...r, [id]: "" })));
    setSel({ kind: "col", id }); setColSel(new Set([id]));
  };
  const insertCase = () => {
    if (selRow == null) return;
    const empty = Object.fromEntries(columns.map((c) => [c.id, ""]));
    setRows((rs) => { const n = [...rs]; n.splice(selRow + 1, 0, empty); return n; });
    setExcluded((s) => new Set([...s].map((i) => (i > selRow ? i + 1 : i))));
    setSelRow(selRow + 1);
  };
  const deleteExcludedCases = () => {
    if (excluded.size === 0) return;
    if (!window.confirm(`Delete ${excluded.size} excluded case${excluded.size === 1 ? "" : "s"}? This cannot be undone.`)) return;
    setRows((rs) => rs.filter((_, i) => !excluded.has(i)));
    setExcluded(new Set()); setSelRow(null);
  };

  /* ---- compaction ---- */
  // Columns eligible to compact = currently SELECTED, continuous, loose (non-leaf) columns, in column order.
  const compactableIds = () => columns.filter((c) => colSel.has(c.id) && !leafToCompact[c.id] && (c.type === "real" || c.type === "integer")).map((c) => c.id);
  const openCompact = () => {
    const picks = compactableIds();
    if (picks.length < 2) return;
    setDlg({ picks, name: "Measure", factors: [{ name: "Factor1", levels: "level1, level2" }] });
  };
  const createCompact = (built) => {
    // reorder columns so the leaf block is contiguous (in chosen order) at the
    // position of the first leaf; all non-leaf columns keep their relative order.
    const leafSet = new Set(built.leaves);
    const leafCols = built.leaves.map((id) => columns.find((c) => c.id === id)).filter(Boolean);
    const ordered = [];
    let inserted = false;
    for (const c of columns) {
      if (leafSet.has(c.id)) { if (!inserted) { ordered.push(...leafCols); inserted = true; } }
      else ordered.push(c);
    }
    if (!inserted) ordered.push(...leafCols);
    setColumns(ordered);
    setCompacts((cs) => [...cs, built]);
    setColSel(new Set());
    setDlg(null);
  };
  const detachAnalyses = (cid) => setAnalyses((arr) => arr.map((a) => {
    if (a.type === "anova" && a.roles.dep && a.roles.dep.kind === "compact" && a.roles.dep.id === cid) { const roles = { ...a.roles }; delete roles.dep; return { ...a, roles }; }
    return a;
  }));
  const expand = (cid) => {
    setCompacts((cs) => cs.filter((c) => c.id !== cid));
    detachAnalyses(cid);
    if (sel && sel.kind === "compact" && sel.id === cid) setSel(null);
  };
  // Remove ONE within factor: splits the compact into one variable per level of
  // that factor (VibeStat's layer-peeling). Removing the last factor dissolves to loose columns.
  const expandFactor = (cid, factorName) => {
    const cp = compactById[cid];
    if (!cp) return;
    const removed = cp.factors.find((f) => f.name === factorName);
    if (!removed) return;
    const remaining = cp.factors.filter((f) => f.name !== factorName);
    const newCompacts = [];
    if (remaining.length > 0) {
      removed.levels.forEach((lvl) => {
        const leaves = cp.leaves.filter((id, idx) => leafLevels(cp, idx)[factorName] === lvl);
        newCompacts.push({ id: "cp" + Math.random().toString(36).slice(2, 7), name: `${cp.name} (${lvl})`, factors: remaining.map((f) => ({ name: f.name, levels: [...f.levels] })), leaves });
      });
    }
    // reorder columns so each new compact's leaves stay contiguous (dissolved leaves remain loose, in place)
    const leafToNew = {};
    newCompacts.forEach((nc) => nc.leaves.forEach((id) => (leafToNew[id] = nc.id)));
    const emitted = new Set(); const order = [];
    for (const c of columns) {
      const ncId = leafToNew[c.id];
      if (ncId) { if (!emitted.has(ncId)) { emitted.add(ncId); order.push(...newCompacts.find((n) => n.id === ncId).leaves.map((id) => colById[id])); } }
      else order.push(c);
    }
    setColumns(order);
    setCompacts((cs) => [...cs.filter((c) => c.id !== cid), ...newCompacts]);
    detachAnalyses(cid);
    if (sel && sel.kind === "compact" && sel.id === cid) setSel(null);
  };

  /* ---- analysis management ---- */
  const addAnalysis = (item) => {
    const spec = typeof item === "string" ? { type: item } : item;
    const id = "a" + Math.random().toString(36).slice(2, 7);
    const a = { id, type: spec.type, roles: {} };
    let anovaInherited = false;
    if (spec.type === "anova") {
      a.output = spec.output || "table";
      a.cfg = { ...DEFAULT_ANOVA_CFG };
      if (spec.method) a.method = spec.method;
      // Follow-up outputs (means table, interaction graphs, post-hoc) borrow the
      // variable assignments and settings of the ANOVA currently selected in the
      // Analysis View, so they don't re-prompt. The ANOVA Setup dialog appears only
      // for a fresh ANOVA Table, or for a follow-up added with no parent selected.
      if (a.output !== "table") {
        const cur = analyses.find((x) => x.id === selAnalysis);
        if (cur && cur.type === "anova" && cur.roles && cur.roles.dep) {
          a.roles = { dep: { ...cur.roles.dep }, between: [...(cur.roles.between || [])] };
          a.cfg = { ...DEFAULT_ANOVA_CFG, ...cur.cfg };
          if (a.output === "posthoc" && cur.phEffect) a.phEffect = cur.phEffect;
          anovaInherited = true;
        }
      }
    }
    setAnalyses((arr) => [...arr, a]);
    setSelAnalysis(id);
    setWindows((w) => ({ ...w, view: { ...w.view, open: true } }));
    focus("view"); setOpenMenu(null);
    if (spec.type === "anova" && !anovaInherited) setAnovaDlg({ id, cfg: { ...a.cfg } });
  };
  const removeAnalysis = (id) => setAnalyses((a) => a.filter((x) => x.id !== id));
  const setOutput = (id, output) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, output } : a)));
  const setOpt = useCallback((id, patch) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, opts: { ...(a.opts || {}), ...patch } } : a))), []);
  const toggleSwap = (id) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, swap: !a.swap } : a)));
  const setMethod = (id, method) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, method } : a)));
  const setPhEffect = (id, phEffect) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, phEffect } : a)));
  const openAnovaSetup = (id) => { const a = analyses.find((x) => x.id === id); if (a) setAnovaDlg({ id, cfg: { ...DEFAULT_ANOVA_CFG, ...a.cfg } }); };
  const applyAnovaSetup = (cfg) => { setAnalyses((arr) => arr.map((a) => (a.id === anovaDlg.id ? { ...a, cfg } : a))); setAnovaDlg(null); };
  const openPlotSetup = (id) => {
    const a = analyses.find((x) => x.id === id); if (!a) return;
    const keys = new Set();
    try {
      const m = anovaModel(a, colById, compactById, activeRows);
      if (!m.error) {
        const c = { ...DEFAULT_ANOVA_CFG, ...a.cfg };
        effectList(m.factors, c.effects).forEach((E) => { const rws = cellMeans(m.long, E, c.alpha, c.errorBars); graphModel(E, rws, m.levelOrder, a.swap).facets.forEach((f) => f.seriesKeys.forEach((k) => keys.add(k))); });
      }
    } catch (e) { /* keys stays empty */ }
    setPlotDlg({ id, plot: { ...DEFAULT_PLOT, ...(a.plot || {}) }, seriesKeys: [...keys] });
  };
  const applyPlotSetup = (plot) => { setAnalyses((arr) => arr.map((a) => (a.id === plotDlg.id ? { ...a, plot } : a))); setPlotDlg(null); };
  const commitPlotSize = useCallback((id, w, h) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, plot: { ...DEFAULT_PLOT, ...(a.plot || {}), width: w, height: h } } : a))), []);
  const commitPlotLegend = useCallback((id, xy) => setAnalyses((arr) => arr.map((a) => (a.id === id ? { ...a, plot: { ...DEFAULT_PLOT, ...(a.plot || {}), legendXY: xy } } : a))), []);
  const growVars = useCallback(() => setWindows((w) => (w.vars.open && w.vars.h < 372 ? { ...w, vars: { ...w.vars, h: 372 } } : w)), []);

  const accepts = (accept, s) => {
    if (!s) return false;
    if (accept === "depOrCompact") return s.kind === "compact" || (s.kind === "col" && isContinuous(colById[s.id]?.type));
    if (accept === "compactOnly") return s.kind === "compact";
    if (accept === "continuous") return s.kind === "col" && isContinuous(colById[s.id]?.type);
    if (accept === "categorical") return s.kind === "col" && !isContinuous(colById[s.id]?.type);
    return false;
  };

  const assignRole = (roleKey) => {
    if (!selAnalysis || !sel) return;
    setAnalyses((arr) => arr.map((a) => {
      if (a.id !== selAnalysis) return a;
      const role = rolesOf(a).find((r) => r.key === roleKey);
      if (!accepts(role.accept, sel)) return a;
      const roles = { ...a.roles };
      if (role.accept === "depOrCompact" || role.accept === "compactOnly") { roles[roleKey] = { ...sel }; }
      else if (role.multiple) {
        const cur = roles[roleKey] || [];
        roles[roleKey] = cur.includes(sel.id) ? cur.filter((v) => v !== sel.id) : [...cur, sel.id];
      } else { roles[roleKey] = sel.id; }
      return { ...a, roles };
    }));
  };

  const selectedAnalysis = analyses.find((a) => a.id === selAnalysis);
  const selectedDef = selectedAnalysis ? { ...ANALYSES[selectedAnalysis.type], roles: rolesOf(selectedAnalysis) } : null;

  /* ---- display ordering (loose + contiguous compact leaf blocks) ---- */
  const display = useMemo(() => {
    const units = []; const consumed = new Set();
    for (const c of columns) {
      if (consumed.has(c.id)) continue;
      const cid = leafToCompact[c.id];
      if (cid && compactById[cid] && compactById[cid].leaves[0] === c.id) {
        const cp = compactById[cid];
        cp.leaves.forEach((id) => consumed.add(id));
        units.push({ kind: "compact", compact: cp });
      } else if (cid && compactById[cid]) {
        // a leaf whose block head isn't reached yet — emit whole block now
        const cp = compactById[cid];
        cp.leaves.forEach((id) => consumed.add(id));
        units.push({ kind: "compact", compact: cp });
      } else {
        units.push({ kind: "col", col: c });
      }
    }
    const orderedColIds = [];
    units.forEach((u) => { if (u.kind === "col") orderedColIds.push(u.col.id); else u.compact.leaves.forEach((id) => orderedColIds.push(id)); });
    const maxF = compacts.reduce((m, c) => Math.max(m, c.factors.length), 0);
    return { units, orderedColIds, maxF };
  }, [columns, compacts, compactById, leafToCompact]);
  // Drag-and-drop column reordering. A loose (single) column moves on its own; a
  // compact's leaf columns move together as one block. Individual leaves of a compact
  // can't be reordered. Reorder the display units, then flatten back to a column order.
  const moveColumnUnit = (dragKey, dropKey) => {
    if (!dragKey || !dropKey || dragKey === dropKey) return;
    const units = display.units;
    const keyOf = (u) => (u.kind === "col" ? "col:" + u.col.id : "cmp:" + u.compact.id);
    const fromIdx = units.findIndex((u) => keyOf(u) === dragKey);
    const toIdx = units.findIndex((u) => keyOf(u) === dropKey);
    if (fromIdx < 0 || toIdx < 0) return;
    const arr = units.slice();
    const [moved] = arr.splice(fromIdx, 1);
    let insertIdx = arr.findIndex((u) => keyOf(u) === dropKey);
    if (fromIdx < toIdx) insertIdx += 1; // dragging rightwards lands after the target
    arr.splice(insertIdx, 0, moved);
    const byId = {}; columns.forEach((c) => { byId[c.id] = c; });
    const newCols = [];
    arr.forEach((u) => {
      if (u.kind === "col") { if (byId[u.col.id]) newCols.push(byId[u.col.id]); }
      else u.compact.leaves.forEach((id) => { if (byId[id]) newCols.push(byId[id]); });
    });
    if (newCols.length === columns.length) setColumns(newCols);
  };
  // ordered numeric columns (raw data cells) for the formula builder picker
  const numericPickCols = useMemo(() => {
    const out = [];
    display.units.forEach((u) => {
      if (u.kind === "col") { if (u.col.type === "real" || u.col.type === "integer") out.push({ id: u.col.id, label: u.col.name, group: null }); }
      else u.compact.leaves.forEach((id) => out.push({ id, label: (colById[id] && colById[id].name) || id, group: u.compact.name }));
    });
    return out;
  }, [display, colById]);
  const [fxDlg, setFxDlg] = useState(null); // formula builder: { colId }
  const openFormulaBuilder = (cid) => setFxDlg({ colId: cid });

  /* ---- dataset open / save / import ---- */
  const serializeProject = () => ({ app: "VibeStat", version: 4, savedAt: new Date().toISOString(), columns, rows, compacts, analyses, selAnalysis, excluded: [...excluded], colW, canvasBg });
  const loadProject = (d, name) => {
    if (!d || typeof d !== "object" || !Array.isArray(d.columns) || !Array.isArray(d.rows) || d.columns.length === 0) { window.alert("That file has no usable columns/rows."); return; }
    setDocName(name ? String(name).replace(/\.vibestat\.json$/i, "").replace(/\.json$/i, "") : null);
    setColumns(d.columns);
    setRows(d.rows);
    setCompacts(Array.isArray(d.compacts) ? d.compacts : []);
    const an = Array.isArray(d.analyses) ? d.analyses : [];
    setAnalyses(an);
    setSelAnalysis(d.selAnalysis && an.some((a) => a.id === d.selAnalysis) ? d.selAnalysis : (an[0] ? an[0].id : null));
    setExcluded(new Set(Array.isArray(d.excluded) ? d.excluded : []));
    setColW(d.colW && typeof d.colW === "object" ? d.colW : {});
    setCanvasBg(d.canvasBg && typeof d.canvasBg === "object" ? d.canvasBg : { ...CANVAS_DEFAULT });
    setSel(null); setSelRow(null); setColSel(new Set());
    resetHistory({ columns: d.columns, rows: d.rows, compacts: Array.isArray(d.compacts) ? d.compacts : [], analyses: an, selAnalysis: d.selAnalysis && an.some((a) => a.id === d.selAnalysis) ? d.selAnalysis : (an[0] ? an[0].id : null), excluded: Array.isArray(d.excluded) ? d.excluded : [], colW: d.colW && typeof d.colW === "object" ? d.colW : {} });
  };
  const onSaveDataset = async () => {
    try {
      let base = (docName || "study").replace(/\.vibestat\.json$/i, "").replace(/\.json$/i, "");
      if (!IN_TAURI) { const inp = window.prompt("Save dataset as:", base); if (inp == null) return; base = inp.trim() || base; }
      const res = await ioSaveJSON(serializeProject(), base + ".vibestat.json");
      const fn = typeof res === "string" && res ? res : base + ".vibestat.json";
      setDocName(fn.replace(/\.vibestat\.json$/i, "").replace(/\.json$/i, ""));
    } catch (e) { window.alert("Save failed: " + (e && e.message ? e.message : e)); }
  };
  const onNewDataset = () => {
    if (!window.confirm("Start a new, empty dataset? Any unsaved data will be lost.")) return;
    loadProject({ columns: [{ id: "c1", name: "Column 1", type: "real" }], rows: Array.from({ length: 16 }, () => ({})), compacts: [], analyses: [], selAnalysis: null, excluded: [], colW: {} }, null);
  };
  const onOpenDataset = async () => { try { const d = await ioOpenJSON(); if (d) loadProject(d, d.__name); } catch (e) { window.alert("Open failed: " + (e && e.message ? e.message : e)); } };
  const onImportData = async () => { try { const d = await ioImportTable(); if (!d) return; if (d.__pzfxMulti) setPzfxPick(d); else loadProject(d, d.__name); } catch (e) { window.alert("Import failed: " + (e && e.message ? e.message : e)); } };

  const sortByColumn = (dir) => {
    if (!sel || sel.kind !== "col") return;
    const cid = sel.id, idx = computedRows.map((_, i) => i), val = (i) => computedRows[i][cid];
    const nonEmpty = idx.filter((i) => { const v = val(i); return !(v === "" || v == null); });
    const allNum = nonEmpty.length > 0 && nonEmpty.every((i) => isFinite(Number(val(i))));
    idx.sort((a, b) => {
      const va = val(a), vb = val(b), ea = va === "" || va == null, eb = vb === "" || vb == null;
      if (ea && eb) return 0; if (ea) return 1; if (eb) return -1; // empties always last
      const c = allNum ? Number(va) - Number(vb) : String(va).localeCompare(String(vb));
      return dir === "desc" ? -c : c;
    });
    setRows((rs) => idx.map((i) => rs[i]));
    setExcluded(() => new Set(idx.map((oldI, newI) => (excluded.has(oldI) ? newI : -1)).filter((x) => x >= 0)));
    setSelRow(null);
  };

  const editItems = [
    { label: "Undo", onClick: undo, disabled: !canUndo },
    { label: "Redo", onClick: redo, disabled: !canRedo },
  ];
  const manageItems = [
    { label: "Add Variable", onClick: addVar },
    { label: "Add Formula Variable", onClick: addFormulaVar },
    { label: "Insert Variable (after selected)", onClick: insertColumn, disabled: !sel },
    { sep: true },
    { label: "Add Case", onClick: addCase },
    { label: "Insert Case (below selected)", onClick: insertCase, disabled: selRow == null },
    { sep: true },
    { label: "Group into Compact Variable…", onClick: openCompact, disabled: compactableIds().length < 2 },
    { label: "Delete Variable", onClick: deleteSelectedVar, disabled: !sel },
    { sep: true },
    { label: "Sort Cases ↑ (by selected variable)", onClick: () => sortByColumn("asc"), disabled: !sel || sel.kind !== "col" },
    { label: "Sort Cases ↓ (by selected variable)", onClick: () => sortByColumn("desc"), disabled: !sel || sel.kind !== "col" },
    { sep: true },
    { label: "Include All Cases", onClick: includeAllCases, disabled: excluded.size === 0 },
    { label: "Delete Excluded Cases", onClick: deleteExcludedCases, disabled: excluded.size === 0 },
  ];

  const menuModel = buildMenuModel({ onAbout: () => setAboutOpen(true), onSettings: () => setSettingsOpen(true), onNew: onNewDataset, onOpen: onOpenDataset, onSave: onSaveDataset, onImport: onImportData, onSample: () => loadProject(SAMPLE_PROJECT, "mixed-RM-study.vibestat.json"), editItems, manageItems, addAnalysis, windows, toggleWin });
  const nativeMenuReady = useNativeMenu(menuModel);
  return (
    <div style={{ ...(IN_TAURI ? { position: "fixed", top: 0, left: 0, right: 0, bottom: 0 } : { position: "relative", height: 900 }), overflow: "hidden", fontFamily: FONT, fontSize: 11, userSelect: "none", ...canvasStyle(canvasBg), border: `1px solid ${PLAT.darker}` }}>
      {nativeMenuReady
        ? <div {...(NATIVE_MAC ? { "data-tauri-drag-region": true } : {})} style={{ position: "absolute", top: 0, left: 0, right: 0, height: MAC_BAR_H, zIndex: 1000, background: "transparent" }} />
        : <MenuBar model={menuModel} openMenu={openMenu} setOpenMenu={setOpenMenu} />}

      {windows.data.open && (
        <MacWindow title={docName || "Dataset"} width={windows.data.w} height={windows.data.h} pos={windows.data} setPos={setPos("data")} z={zOf("data")} active={isActive("data")} onFocus={() => focus("data")} onClose={() => toggleWin("data")}>
          <DataGrid display={display} rows={computedRows} setCell={setCell} setColType={setColType} setColName={setColName} setColDecimals={setColDecimals} colById={colById} sel={sel} setSel={setSel} expand={expand} expandFactor={expandFactor} leafToCompact={leafToCompact} excluded={excluded} toggleExcluded={toggleExcluded} colW={colW} setColW={setColW} setColFormula={setColFormula} recomputeStatic={recomputeStatic} openFormulaBuilder={openFormulaBuilder} selRow={selRow} setSelRow={setSelRow} colSel={colSel} setColSel={setColSel} onReorder={moveColumnUnit} onPasteCell={pasteAt} onCellFocus={(ri, cid) => { pasteAnchorRef.current = { ri, cid }; }} onCellRange={(info) => { cellRangeRef.current = info; if (info && info.top) pasteAnchorRef.current = { ri: info.top.ri, cid: info.top.cid }; }} />
          <div style={{ display: "flex", gap: 6, padding: "5px 6px", borderTop: `1px solid ${PLAT.dark}`, alignItems: "center", flexWrap: "wrap" }}>
            <MacBtn onClick={addCase}>Add Row</MacBtn>
            <MacBtn onClick={addVar}>Add Variable</MacBtn>
            <MacBtn onClick={openCompact} disabled={compactableIds().length < 2} primary={compactableIds().length >= 2}>Compact{compactableIds().length ? ` (${compactableIds().length})` : ""}…</MacBtn>
            <MacBtn onClick={() => { if (sel && sel.kind === "compact") expand(sel.id); }} disabled={!(sel && sel.kind === "compact")}>Expand</MacBtn>
            <MacBtn onClick={deleteSelectedVar} disabled={!sel}>Delete Variable</MacBtn>
            <span style={{ color: "#556", fontSize: 10 }}>select continuous columns (Shift- or ⌘/⌥-click), then Compact; select a compact, then Expand</span>
            {excluded.size > 0 && <span style={{ color: "#a40", fontSize: 10, fontWeight: "bold" }}>{excluded.size} case{excluded.size > 1 ? "s" : ""} excluded · {rows.length - excluded.size} active (double-click a case number to restore)</span>}
          </div>
        </MacWindow>
      )}

      {windows.browser.open && (
        <MacWindow title="Analysis Browser" width={windows.browser.w} height={windows.browser.h} pos={windows.browser} setPos={setPos("browser")} z={zOf("browser")} active={isActive("browser")} onFocus={() => focus("browser")} onClose={() => toggleWin("browser")}>
          <AnalysisBrowser expanded={expanded} setExpanded={setExpanded} addAnalysis={addAnalysis} />
        </MacWindow>
      )}

      {windows.vars.open && (
        <MacWindow title="Variables" width={windows.vars.w} height={windows.vars.h} pos={windows.vars} setPos={setPos("vars")} z={zOf("vars")} active={isActive("vars")} onFocus={() => focus("vars")} onClose={() => toggleWin("vars")}>
          <VariablesBrowser columns={columns} compacts={compacts} leafToCompact={leafToCompact} sel={sel} setSel={setSel} def={selectedDef} accepts={accepts} assignRole={assignRole} analysis={selectedAnalysis} colById={colById} compactById={compactById} expand={expand} expandFactor={expandFactor} colSel={colSel} setColSel={setColSel} openCompact={openCompact} setCompactDecimals={setCompactDecimals} onNeedHeight={growVars} />
        </MacWindow>
      )}

      {windows.view.open && (
        <MacWindow title="Analysis View" width={windows.view.w} height={windows.view.h} pos={windows.view} setPos={setPos("view")} z={zOf("view")} active={isActive("view")} onFocus={() => focus("view")} onClose={() => toggleWin("view")}>
          <AnalysisView analyses={analyses} selAnalysis={selAnalysis} setSelAnalysis={setSelAnalysis} colById={colById} compactById={compactById} valuesOf={valuesOf} rows={activeRows} removeAnalysis={removeAnalysis} onOutput={setOutput} onConfig={openAnovaSetup} onSwap={toggleSwap} onMethod={setMethod} onPhEffect={setPhEffect} onPlot={openPlotSetup} onPlotResize={commitPlotSize} onPlotLegend={commitPlotLegend} onOpt={setOpt} />
        </MacWindow>
      )}

      {dlg && <CompactDialog dlg={dlg} setDlg={setDlg} columns={columns} onCreate={createCompact} />}
      {fxDlg && <FormulaBuilderDialog col={colById[fxDlg.colId]} numericCols={numericPickCols} compacts={compacts}
        onApply={(formula) => { setColFormula(fxDlg.colId, formula); const cc = colById[fxDlg.colId]; if (cc && cc.type === "formula_static") setTimeout(() => recomputeStatic(fxDlg.colId), 0); setFxDlg(null); }}
        onCancel={() => setFxDlg(null)} />}
      {anovaDlg && <AnovaSetupDialog dlg={anovaDlg} onApply={applyAnovaSetup} onCancel={() => setAnovaDlg(null)} />}
      {aboutOpen && <AboutDialog onClose={() => setAboutOpen(false)} />}
      {settingsOpen && <SettingsDialog bg={canvasBg} setBg={setCanvasBg} onClose={() => setSettingsOpen(false)} />}
      {pzfxPick && <PzfxTableDialog pick={pzfxPick} onPick={(i) => { loadProject(pzfxPick.projects[i], pzfxPick.__name); setPzfxPick(null); }} onCancel={() => setPzfxPick(null)} />}
      {plotDlg && <PlotSettingsDialog dlg={plotDlg} onApply={applyPlotSetup} onCancel={() => setPlotDlg(null)} />}
    </div>
  );
}

/* =========================================================================
   MENU BAR
   ========================================================================= */
function MenuList({ items, onClose }) {
  const [openSub, setOpenSub] = useState(null);
  return (
    <div style={{ minWidth: 180, background: PLAT.faceLite, ...bevelOut, boxShadow: "2px 2px 0 rgba(0,0,0,0.3)", padding: "2px 0" }}>
      {items.map((it, i) => {
        if (it.sep) return <div key={i} style={{ height: 1, background: PLAT.dark, margin: "3px 6px", opacity: 0.5 }} />;
        const hasSub = !!(it.sub && it.sub.length);
        return (
          <div key={i} style={{ position: "relative" }} onMouseEnter={() => setOpenSub(hasSub ? i : null)}>
            <div
              onClick={() => { if (it.disabled || hasSub) return; if (it.onClick) it.onClick(); onClose(); }}
              style={{ display: "flex", justifyContent: "space-between", gap: 16, padding: "3px 16px", whiteSpace: "nowrap", cursor: it.disabled ? "default" : "pointer", color: openSub === i ? "#fff" : (it.disabled ? "#aaa" : "#000"), background: openSub === i ? PLAT.selBorder : "transparent" }}
              onMouseEnter={(e) => { if (!it.disabled) { e.currentTarget.style.background = PLAT.selBorder; e.currentTarget.style.color = "#fff"; } }}
              onMouseLeave={(e) => { const on = openSub === i; e.currentTarget.style.background = on ? PLAT.selBorder : "transparent"; e.currentTarget.style.color = on ? "#fff" : (it.disabled ? "#aaa" : "#000"); }}>
              <span>{it.label}</span>{hasSub && <span style={{ fontSize: 9, marginLeft: 8 }}>▶</span>}
            </div>
            {hasSub && openSub === i && (
              <div style={{ position: "absolute", top: -3, left: "100%" }}>
                <MenuList items={it.sub} onClose={onClose} />
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function buildAnalyzeMenu(addAnalysis) {
  const byCat = {}; BROWSER_ITEMS.forEach((it) => (byCat[it.cat] = byCat[it.cat] || []).push(it));
  const out = [];
  CATEGORY_ORDER.forEach((cat) => {
    const items = byCat[cat] || []; if (!items.length) return;
    const entries = items.map((it) => it.nested
      ? { label: it.label, sub: it.nested.map((sub) => ({ label: sub.label, onClick: () => addAnalysis(sub) })) }
      : { label: it.label, onClick: () => addAnalysis(it) });
    if (entries.length === 1 && !items[0].nested) out.push(entries[0]); // promote single-item category
    else out.push({ label: cat, sub: entries });
  });
  const cmpIdx = out.findIndex((e) => e.label === "Comparison");
  if (cmpIdx > 0) out.splice(cmpIdx, 0, { sep: true }); // divide summaries/plots from inferential tests
  return out;
}

// Frameless macOS chrome detection. True ONLY inside the Tauri macOS build; always
// false in a browser/artifact (no Tauri globals), so this is inert outside the app.
const IN_TAURI = (() => { try { return typeof window !== "undefined" && !!(window.__TAURI_INTERNALS__ || window.__TAURI__ || window.isTauri); } catch (e) { return false; } })();
const NATIVE_MAC = (() => {
  try {
    if (!IN_TAURI) return false;
    const plat = (typeof navigator !== "undefined" && ((navigator.userAgentData && navigator.userAgentData.platform) || navigator.platform || navigator.userAgent)) || "";
    return /Mac/i.test(plat);
  } catch (e) { return false; }
})();
const MAC_BAR_H = 28;      // taller top strip so the macOS traffic lights sit inside it
const MAC_BAR_PADL = 78;   // left inset clearing the traffic lights
// Assign a stable id to every actionable menu leaf so the native macOS menu can
// dispatch to (and later update) the right command without rebuilding the menu.
function assignMenuIds(model) {
  model.forEach((top) => {
    const walk = (items, prefix) => items.forEach((it, i) => {
      if (it.sep) return;
      it.__id = prefix + ":" + i;
      if (it.sub) walk(it.sub, it.__id);
    });
    walk(top.items, top.id);
  });
  return model;
}
// Single source of truth for the menu structure. Consumed by BOTH the in-canvas
// MenuBar and (in the desktop kit only) the native macOS menu bar.
function buildMenuModel({ onAbout, onSettings, onNew, onOpen, onSave, onImport, onSample, editItems, manageItems, addAnalysis, windows, toggleWin }) {
  const win = (k, label) => ({ label: (windows[k].open ? "\u2713 " : "   ") + label, checked: windows[k].open, onClick: () => toggleWin(k) });
  return assignMenuIds([
    { id: "apple", label: "\uF8FF", apple: true, items: [{ label: "About VibeStat\u2026", onClick: onAbout }, { sep: true }, { label: "Settings\u2026", onClick: onSettings }] },
    { id: "file", label: "File", items: [
      { label: "New Dataset\u2026", onClick: onNew },
      { label: "Open Dataset\u2026", onClick: onOpen },
      { label: "Save Dataset\u2026", onClick: onSave },
      { label: "Import CSV / Excel / Prism\u2026", onClick: onImport },
      { sep: true },
      { label: "Open Sample Dataset", onClick: onSample },
    ] },
    { id: "edit", label: "Edit", items: editItems },
    { id: "manage", label: "Manage", items: manageItems },
    { id: "analyze", label: "Analyze", items: buildAnalyzeMenu(addAnalysis) },
    { id: "windows", label: "Windows", items: [win("data", "Dataset"), win("browser", "Analysis Browser"), win("vars", "Variables"), win("view", "Analysis View")] },
  ]);
}
// Native macOS menu bar. The REAL implementation lives in the desktop kit's App.jsx
// (it imports @tauri-apps/api/menu). In a browser/artifact this is a no-op returning
// false, so the in-canvas MenuBar is always used. Keeps the artifact dependency-free.
// ---- Native macOS menu bar (module-level singletons so it survives React StrictMode's
// mount/unmount/mount double-invoke and so menu actions always see the live model) ----
let _vsModel = [];        // latest menu model from the live component
const _vsReg = {};        // __id -> native item instance (for in-place enable/check updates)
let _vsBuilt = false;     // native menu build has started/succeeded
let _vsFailed = false;    // native menu build failed -> fall back to the in-canvas menu
let _vsForce = null;      // forceUpdate of the live component (only needed on failure)
function _vsStripMark(s) { return String(s == null ? "" : s).replace(/^[\u2713\u2714]\s?/, "").replace(/^\s{3}/, ""); }
function _vsMenuSig(model) { let out = ""; const walk = (arr) => arr.forEach((it) => { if (it.sep) return; out += (it.__id || "") + (it.disabled ? "D" : "") + (it.checked ? "C" : "") + "|"; if (it.sub) walk(it.sub); }); model.forEach((t) => walk(t.items)); return out; }
function _vsDispatch(id) {
  const find = (items) => { for (const it of items) { if (it.__id === id) return it; if (it.sub) { const r = find(it.sub); if (r) return r; } } return null; };
  for (const top of _vsModel) { const r = find(top.items); if (r) { if (!r.disabled && r.onClick) r.onClick(); return; } }
}
async function _vsBuildNativeMenu() {
  const pre = async (item) => { try { return await PredefinedMenuItem.new({ item }); } catch (e) { return null; } };
  const buildItem = async (it) => {
    if (it.sep) return await pre("Separator");
    if (it.sub && it.sub.length) {
      const kids = []; for (const k of it.sub) { const ni = await buildItem(k); if (ni) kids.push(ni); }
      return await Submenu.new({ text: _vsStripMark(it.label), items: kids });
    }
    const id = it.__id;
    if (it.checked != null) {
      const ci = await CheckMenuItem.new({ text: _vsStripMark(it.label), checked: !!it.checked, enabled: !it.disabled, action: () => _vsDispatch(id) });
      if (id) _vsReg[id] = ci; return ci;
    }
    const mi = await MenuItem.new({ text: _vsStripMark(it.label), enabled: !it.disabled, action: () => _vsDispatch(id) });
    if (id) _vsReg[id] = mi; return mi;
  };
  const apple = _vsModel.find((t) => t.apple);
  const appItems = [];
  if (apple) for (const it of apple.items) { const ni = await buildItem(it); if (ni) appItems.push(ni); }
  for (const it of ["Separator", "Hide", "HideOthers", "ShowAll", "Separator", "Quit"]) { const pm = await pre(it); if (pm) appItems.push(pm); }
  const submenus = [await Submenu.new({ text: "VibeStat", items: appItems })];
  for (const top of _vsModel) {
    if (top.apple) continue;
    const kids = []; for (const it of top.items) { const ni = await buildItem(it); if (ni) kids.push(ni); }
    // macOS routes the standard clipboard shortcuts through these menu items; without
    // them, Cmd-X/C/V/A never reach the webview. Predefined items also fire the DOM
    // copy/cut/paste events that the grid's handlers listen for.
    if (top.id === "edit") { const sep = await pre("Separator"); if (sep) kids.push(sep); for (const it of ["Cut", "Copy", "Paste", "SelectAll"]) { const pm = await pre(it); if (pm) kids.push(pm); } }
    submenus.push(await Submenu.new({ text: top.label, items: kids }));
  }
  const menu = await Menu.new({ items: submenus });
  await menu.setAsAppMenu();
}
function useNativeMenu(model) {
  _vsModel = model;
  const [, setTick] = useState(0);
  const force = () => setTick((t) => t + 1);
  useEffect(() => { _vsForce = force; return () => { if (_vsForce === force) _vsForce = null; }; });
  useEffect(() => {
    if (!NATIVE_MAC || _vsBuilt) return;
    _vsBuilt = true;
    _vsBuildNativeMenu().catch((e) => {
      _vsFailed = true; _vsBuilt = false;
      console.error("VibeStat: native menu setup failed; keeping in-canvas menu.", e);
      if (_vsForce) _vsForce();
    });
  }, []);
  // Keep enabled/checked states in sync WITHOUT rebuilding the menu (Tauri #10121).
  useEffect(() => {
    if (!NATIVE_MAC || _vsFailed) return;
    const walk = (arr) => arr.forEach((it) => {
      if (it.sub) return walk(it.sub);
      const native = it.__id && _vsReg[it.__id]; if (!native) return;
      try { if (typeof native.setEnabled === "function") native.setEnabled(!it.disabled); } catch (e) {}
      try { if (it.checked != null && typeof native.setChecked === "function") native.setChecked(!!it.checked); } catch (e) {}
    });
    _vsModel.forEach((t) => walk(t.items));
  }, [_vsMenuSig(model)]);
  return NATIVE_MAC && !_vsFailed;
}
function MenuBar({ model, openMenu, setOpenMenu }) {
  return (
    <div {...(NATIVE_MAC ? { "data-tauri-drag-region": true } : {})} style={{ position: "absolute", top: 0, left: 0, right: 0, height: NATIVE_MAC ? MAC_BAR_H : 20, zIndex: 1000, background: PLAT.faceLite, borderBottom: `1px solid ${PLAT.darker}`, display: "flex", alignItems: "center", paddingLeft: NATIVE_MAC ? MAC_BAR_PADL : 6, paddingRight: 6, fontSize: 11 }} onMouseLeave={() => setOpenMenu(null)}>
      {model.map((top) => (
        <div key={top.id} style={{ position: "relative" }}>
          <div onClick={() => setOpenMenu(openMenu === top.id ? null : top.id)} onMouseEnter={() => openMenu && setOpenMenu(top.id)} style={{ padding: "2px 9px", cursor: "default", fontWeight: top.apple ? "normal" : "bold", background: openMenu === top.id ? PLAT.selBorder : "transparent", color: openMenu === top.id ? "#fff" : "#000" }}>{top.label}</div>
          {openMenu === top.id && (
            <div style={{ position: "absolute", top: NATIVE_MAC ? MAC_BAR_H : 20, left: 0 }}>
              <MenuList items={top.items} onClose={() => setOpenMenu(null)} />
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

/* =========================================================================
   SMALL UI BITS
   ========================================================================= */
function MacBtn({ children, onClick, primary, disabled }) {
  return (
    <button onClick={disabled ? undefined : onClick} disabled={disabled} style={{ fontFamily: FONT, fontSize: 11, padding: "2px 12px", background: disabled ? "#e4e4e4" : PLAT.faceLite, ...bevelOut, borderRadius: 8, cursor: disabled ? "default" : "pointer", boxShadow: primary ? "0 0 0 2px #000" : "none", fontWeight: primary ? "bold" : "normal", color: disabled ? "#aaa" : "#000" }}>
      {children}
    </button>
  );
}
const TYPE_GLYPH = { real: "#", integer: "#", string: "A", category: "\u25a6", formula: "\u0192", formula_static: "\u0192" };
const typeBadge = (t) => (
  <span style={{ fontFamily: "monospace", fontSize: 9, color: "#fff", background: isContinuous(t) ? "#4a6" : "#86c", borderRadius: 2, padding: "0 2px" }}>{TYPE_GLYPH[t]}</span>
);

/* =========================================================================
   DATA GRID  (hierarchical headers for compact variables)
   ========================================================================= */
function DataGrid({ display, rows, setCell, setColType, setColName, setColDecimals, colById, sel, setSel, expand, expandFactor, excluded, toggleExcluded, colW, setColW, setColFormula, recomputeStatic, openFormulaBuilder, selRow, setSelRow, colSel, setColSel, onReorder, onPasteCell, onCellFocus, onCellRange }) {
  const { units, orderedColIds, maxF } = display;
  const isColSel = (id) => colSel && colSel.has(id);
  const anchorRef = useRef(null);
  const dragKeyRef = useRef(null); // synchronous: HTML5 dnd events can fire before a state re-render flushes
  const [dropKey, setDropKey] = useState(null);
  const [focusCell, setFocusCell] = useState(null);
  const [cellSel, setCellSel] = useState(null); // { r0, r1, c0, c1 } rectangle of selected cells (indices)
  const cellAnchorRef = useRef(null);
  useEffect(() => {
    if (!onCellRange) return;
    if (!cellSel) { onCellRange(null); return; }
    const { r0, r1, c0, c1 } = cellSel;
    const cids = orderedColIds.slice(c0, c1 + 1);
    const lines = []; const cells = [];
    for (let r = r0; r <= r1; r++) {
      const rowVals = [];
      for (const cid of cids) { rowVals.push(rows[r] && rows[r][cid] != null ? rows[r][cid] : ""); cells.push({ ri: r, cid }); }
      lines.push(rowVals.join("\t"));
    }
    onCellRange({ text: lines.join("\n"), cells, top: { ri: r0, cid: cids[0] } });
  }, [cellSel, rows, orderedColIds]);
  const keyAtPoint = (x, y) => { const el = document.elementFromPoint(x, y); const cell = el && el.closest ? el.closest("[data-colkey]") : null; return cell ? cell.getAttribute("data-colkey") : null; };
  const startReorder = (key) => {
    dragKeyRef.current = key;
    const onMove = (e) => { const k = keyAtPoint(e.clientX, e.clientY); setDropKey(k && k !== dragKeyRef.current ? k : null); };
    const onUp = (e) => {
      const tk = keyAtPoint(e.clientX, e.clientY), dk = dragKeyRef.current;
      if (dk && tk && dk !== tk && onReorder) onReorder(dk, tk);
      dragKeyRef.current = null; setDropKey(null); document.body.style.userSelect = "";
      window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp);
    };
    document.body.style.userSelect = "none";
    window.addEventListener("mousemove", onMove); window.addEventListener("mouseup", onUp);
  };
  const reorderFor = (key) => ({
    dropping: dropKey === key,
    cellProps: { "data-colkey": key },
    handleProps: { title: "drag to reorder", onMouseDown: (e) => { e.preventDefault(); e.stopPropagation(); startReorder(key); } },
  });
  // Click-drag across cells to select a rectangle (begins only once the pointer leaves the start cell,
  // so a plain click still focuses the cell for editing).
  const startCellDrag = () => {
    let dragging = false;
    const onMove = (e) => {
      const el = document.elementFromPoint(e.clientX, e.clientY);
      const cell = el && el.closest ? el.closest("[data-cell]") : null;
      if (!cell) return;
      const parts = cell.getAttribute("data-cell").split(":"); const r = +parts[0], cc = +parts[1];
      const a = cellAnchorRef.current; if (!a) return;
      if (!dragging && (r !== a.ri || cc !== a.ci)) { dragging = true; document.body.style.userSelect = "none"; if (document.activeElement && document.activeElement.blur) document.activeElement.blur(); }
      if (dragging) setCellSel({ r0: Math.min(a.ri, r), r1: Math.max(a.ri, r), c0: Math.min(a.ci, cc), c1: Math.max(a.ci, cc) });
    };
    const onUp = () => { document.body.style.userSelect = ""; window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
    window.addEventListener("mousemove", onMove); window.addEventListener("mouseup", onUp);
  };
  // Display-only formatting: show N decimals for a real column unless the cell is being edited.
  const fmtCell = (val, c) => { if (!c || c.type !== "real" || c.decimals == null || val === "" || val == null) return val; const n = Number(val); return Number.isFinite(n) ? n.toFixed(c.decimals) : val; };
  // Select a set of column ids; keep `sel` (single) in sync for Delete/Insert + Variables window.
  const selectCols = (ids, anchorId, single) => {
    setColSel(new Set(ids));
    anchorRef.current = anchorId != null ? anchorId : (ids.length ? ids[ids.length - 1] : null);
    if (single) setSel(single);
  };
  const clickCol = (id, e) => { setCellSel(null);
    if (e && e.shiftKey && anchorRef.current && orderedColIds.includes(anchorRef.current)) {
      const a = orderedColIds.indexOf(anchorRef.current), b = orderedColIds.indexOf(id);
      const [lo, hi] = a < b ? [a, b] : [b, a];
      selectCols(orderedColIds.slice(lo, hi + 1), anchorRef.current, { kind: "col", id });
    } else if (e && (e.metaKey || e.ctrlKey || e.altKey)) {
      const n = new Set(colSel || []); n.has(id) ? n.delete(id) : n.add(id);
      selectCols([...n], id, { kind: "col", id });
    } else {
      selectCols([id], id, { kind: "col", id });
    }
  };
  const clickCompact = (cp) => { setCellSel(null); selectCols(cp.leaves, cp.leaves[0], { kind: "compact", id: cp.id }); };
  const clickBand = (cp, leafIds) => { setCellSel(null); selectCols(leafIds, leafIds[0], leafIds.length === cp.leaves.length ? { kind: "compact", id: cp.id } : { kind: "col", id: leafIds[0] }); };
  // header band rows count = maxF (factor bands). plus a compact-name row when maxF>0. plus base row.
  const hasCompacts = maxF > 0;
  const COLW_GUTTER = 30, COLW_DEF = 70;
  const widthOf = (id) => colW[id] || COLW_DEF;
  const startResize = (e, id) => {
    e.preventDefault(); e.stopPropagation();
    const th = e.currentTarget.closest("th");
    const startX = e.clientX, startW = th ? th.offsetWidth : widthOf(id);
    // If the dragged column is part of a multi-selection, resize every selected column to the same width.
    const targets = colSel && colSel.has(id) && colSel.size > 1 ? [...colSel] : [id];
    const move = (ev) => { const w = Math.max(28, startW + ev.clientX - startX); setColW((m) => { const n = { ...m }; targets.forEach((t) => { n[t] = w; }); return n; }); };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  };
  const ColResizeGrip = ({ id }) => <div onMouseDown={(e) => startResize(e, id)} onClick={(e) => e.stopPropagation()} title="drag to resize column"
    onMouseEnter={(e) => { e.currentTarget.style.background = PLAT.selBorder; }} onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
    style={{ position: "absolute", top: 0, bottom: 0, right: 0, width: 6, cursor: "col-resize", zIndex: 4, background: "transparent" }} />;

  // Precompute, for each compact, the band cells per depth.
  const bandCells = (cp, depth) => {
    // returns array of {label, span} for this factor depth across the compact's leaves
    if (depth >= cp.factors.length) return [{ label: "", span: cp.leaves.length, blank: true }];
    const sizes = cp.factors.map((f) => f.levels.length);
    const inner = sizes.slice(depth + 1).reduce((p, s) => p * s, 1);
    const repeats = sizes.slice(0, depth).reduce((p, s) => p * s, 1);
    const f = cp.factors[depth];
    const cells = [];
    for (let r = 0; r < repeats; r++) for (const lv of f.levels) cells.push({ label: lv, span: inner });
    return cells;
  };

  return (
    <div style={{ flex: 1, minHeight: 0, overflow: "auto", background: "#fff" }}>
      <table style={{ borderCollapse: "collapse", fontSize: 11, tableLayout: "fixed", width: COLW_GUTTER + orderedColIds.reduce((s, id) => s + widthOf(id), 0) }}>
        <colgroup>
          <col style={{ width: COLW_GUTTER }} />
          {orderedColIds.map((id) => <col key={id} style={{ width: widthOf(id) }} />)}
        </colgroup>
        <thead>
          {/* compact-name row */}
          {hasCompacts && (
            <tr>
              <th style={{ ...thStyle, width: 26, background: PLAT.face }} rowSpan={maxF + 2}></th>
              {units.map((u, ui) => u.kind === "col" ? (
                <ColHeadCell key={"c" + ui} c={u.col} rowSpan={maxF + 2} sel={isColSel(u.col.id)} onSelect={(e) => clickCol(u.col.id, e)} setColName={setColName} setColType={setColType} setColDecimals={setColDecimals} setColFormula={setColFormula} recomputeStatic={recomputeStatic} openFormulaBuilder={openFormulaBuilder} grip={<ColResizeGrip id={u.col.id} />} reorder={reorderFor("col:" + u.col.id)} />
              ) : (
                <th key={"k" + ui} colSpan={u.compact.leaves.length} {...reorderFor("cmp:" + u.compact.id).cellProps} style={{ ...thStyle, background: "#f3e7dc", color: COMPACT_CLR, fontWeight: "bold", cursor: "pointer", boxShadow: dropKey === ("cmp:" + u.compact.id) ? `inset 3px 0 0 ${PLAT.selBorder}` : undefined }}
                  onClick={() => clickCompact(u.compact)}>
                  <span {...reorderFor("cmp:" + u.compact.id).handleProps} onClick={(e) => e.stopPropagation()} style={{ cursor: "grab", color: "#a8866a", fontSize: 10, marginRight: 3 }}>⠿</span>
                  <span style={{ background: u.compact.leaves.length && u.compact.leaves.every((l) => colSel.has(l)) ? PLAT.sel : "transparent", padding: "0 4px" }}>
                    ▣ {u.compact.name}
                  </span>
                  <span title="expand" onClick={(e) => { e.stopPropagation(); expand(u.compact.id); }} style={{ cursor: "pointer", color: "#900", marginLeft: 6 }}>✕</span>
                </th>
              ))}
            </tr>
          )}
          {/* factor band rows */}
          {hasCompacts && Array.from({ length: maxF }, (_, depth) => (
            <tr key={"band" + depth}>
              {units.map((u, ui) => {
                if (u.kind === "col") return null; // rowSpanned by name row
                const cp = u.compact;
                const fname = depth < cp.factors.length ? cp.factors[depth].name : null;
                const cells = bandCells(cp, depth);
                let off = 0;
                return cells.map((cell, ci) => {
                  const start = off; off += cell.span;
                  const leafIds = cp.leaves.slice(start, start + cell.span);
                  const bandSel = leafIds.length > 0 && leafIds.every((l) => colSel.has(l));
                  return (
                  <th key={"b" + ui + "_" + ci} colSpan={cell.span} onClick={(e) => { e.stopPropagation(); clickBand(cp, leafIds); }}
                    style={{ ...thStyle, background: bandSel ? PLAT.sel : (cell.blank ? "#faf5f0" : "#f7ede3"), fontSize: 10, color: "#634", fontWeight: cell.blank ? "normal" : "bold", whiteSpace: "nowrap", cursor: "pointer" }}>
                    {ci === 0 && fname && (
                      <span onClick={(e) => { e.stopPropagation(); expandFactor(cp.id, fname); }} title={"expand " + fname}
                        style={{ cursor: "pointer", color: "#933", fontSize: 8, marginRight: 4, fontWeight: "bold", border: "1px solid #cba", borderRadius: 2, padding: "0 2px", background: "#fff" }}>
                        ✕{fname}
                      </span>
                    )}
                    {cell.blank ? "" : cell.label}
                  </th>
                  );
                });
              })}
            </tr>
          ))}
          {/* base leaf row */}
          <tr>
            {!hasCompacts && <th style={{ ...thStyle, width: 26, background: PLAT.face }}></th>}
            {units.map((u, ui) => u.kind === "col"
              ? (hasCompacts ? null : <ColHeadCell key={"bc" + ui} c={u.col} rowSpan={1} sel={isColSel(u.col.id)} onSelect={(e) => clickCol(u.col.id, e)} setColName={setColName} setColType={setColType} setColDecimals={setColDecimals} setColFormula={setColFormula} recomputeStatic={recomputeStatic} openFormulaBuilder={openFormulaBuilder} grip={<ColResizeGrip id={u.col.id} />} reorder={reorderFor("col:" + u.col.id)} />)
              : u.compact.leaves.map((id, li) => (
                <th key={"lf" + ui + "_" + li} onClick={(e) => clickCol(id, e)} style={{ ...thStyle, background: isColSel(id) ? PLAT.sel : "#fbf6f1", fontSize: 9, color: "#856", whiteSpace: "nowrap", overflow: "hidden", cursor: "pointer" }}>
                  {id}
                  <ColResizeGrip id={id} />
                </th>
              ))
            )}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => {
            const ex = excluded && excluded.has(ri);
            return (
            <tr key={ri}>
              <td onClick={() => { setSelRow(ri); setCellSel(null); }} onDoubleClick={() => toggleExcluded(ri)} title="click to select this case; double-click to exclude / include it" style={{ ...tdStyle, textAlign: "center", background: ex ? "#d2d2d2" : (ri === selRow ? PLAT.sel : PLAT.face), color: ex ? "#999" : "#555", cursor: "pointer", textDecoration: ex ? "line-through" : "none", fontWeight: ri === selRow ? "bold" : "normal" }}>{ex ? "⊘" : ri + 1}</td>
              {orderedColIds.map((cid, ci) => {
                const c = colById[cid] || { id: cid, type: "real" };
                const fx = isFormula(c.type);
                const val = r[cid] ?? "";
                const inRange = cellSel && ri >= cellSel.r0 && ri <= cellSel.r1 && ci >= cellSel.c0 && ci <= cellSel.c1;
                return (
                  <td key={cid} data-cell={ri + ":" + ci} style={{ ...tdStyle, background: ex ? "#f0f0f0" : inRange ? "#dce6fb" : fx ? "#f3f6ee" : isColSel(cid) ? "#eef2fb" : "#fff" }}>
                    {fx ? (
                      <div title={c.formula ? (c.type === "formula_static" ? "static formula (frozen): " : "dynamic formula: ") + c.formula : "set a formula in the column header"} style={{ padding: "0 2px", textAlign: "right", fontFamily: FONT, fontSize: 11, color: val === "#ERR" ? "#a00" : ex ? "#aaa" : "#225", fontStyle: c.type === "formula_static" ? "normal" : "italic", overflow: "hidden", whiteSpace: "nowrap", textDecoration: ex ? "line-through" : "none" }}>{val}</div>
                    ) : (
                      <input value={focusCell === ri + ":" + cid ? val : fmtCell(val, c)} onMouseDown={(e) => { if (e.shiftKey && cellAnchorRef.current) { e.preventDefault(); const a = cellAnchorRef.current; setCellSel({ r0: Math.min(a.ri, ri), r1: Math.max(a.ri, ri), c0: Math.min(a.ci, ci), c1: Math.max(a.ci, ci) }); } else { cellAnchorRef.current = { ri, ci }; setCellSel({ r0: ri, r1: ri, c0: ci, c1: ci }); if (setColSel) setColSel(new Set()); startCellDrag(); } }} onFocus={() => { setFocusCell(ri + ":" + cid); onCellFocus && onCellFocus(ri, cid); }} onBlur={() => setFocusCell(null)} onChange={(e) => setCell(ri, cid, e.target.value)} onPaste={(e) => { const cd = e.clipboardData || window.clipboardData; const t = cd ? cd.getData("text") : ""; if (/[\t\n]/.test(t)) { e.preventDefault(); onPasteCell(ri, cid, t); } }} style={{ width: "100%", border: "none", background: "transparent", fontFamily: FONT, fontSize: 11, textAlign: isContinuous(c.type) ? "right" : "left", color: ex ? "#aaa" : "#000", textDecoration: ex ? "line-through" : "none" }} />
                    )}
                  </td>
                );
              })}
            </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
function ColHeadCell({ c, rowSpan, sel, onSelect, setColName, setColType, setColDecimals, setColFormula, recomputeStatic, openFormulaBuilder, grip, reorder = {} }) {
  const fx = isFormula(c.type);
  return (
    <th rowSpan={rowSpan} onClick={(e) => onSelect(e)} {...(reorder.cellProps || {})} style={{ ...thStyle, background: sel ? PLAT.sel : PLAT.face, cursor: "pointer", verticalAlign: "top", overflow: "hidden", boxShadow: reorder.dropping ? `inset 3px 0 0 ${PLAT.selBorder}` : undefined }}>
      <div style={{ display: "flex", alignItems: "center", gap: 3, justifyContent: "center" }}>
        {reorder.handleProps && <span {...reorder.handleProps} onClick={(e) => e.stopPropagation()} style={{ cursor: "grab", color: "#9a9a9a", fontSize: 10, lineHeight: 1 }}>⠿</span>}
        {typeBadge(c.type)}
        <input value={c.name} onClick={(e) => e.stopPropagation()} onChange={(e) => setColName(c.id, e.target.value)} style={{ width: 50, minWidth: 0, border: "none", background: "transparent", fontWeight: "bold", fontFamily: FONT, fontSize: 11, textAlign: "center" }} />
      </div>
      <select value={c.type} onChange={(e) => { const t = e.target.value; setColType(c.id, t); if (t === "formula_static" && c.formula) setTimeout(() => recomputeStatic(c.id), 0); }} onClick={(e) => e.stopPropagation()} style={{ fontSize: 9, fontFamily: FONT, width: "100%", marginTop: 1, border: "none", background: "transparent", color: "#444" }}>
        <option value="real">real</option><option value="integer">integer</option><option value="string">string</option><option value="category">category</option>
        <option value="formula">formula (dynamic)</option><option value="formula_static">formula (static)</option>
      </select>
      {fx && (
        <div onClick={(e) => e.stopPropagation()} style={{ display: "flex", alignItems: "center", gap: 2, marginTop: 1 }}>
          <input value={c.formula || ""} placeholder="=AVERAGE([RT (ms)])"
            onChange={(e) => setColFormula(c.id, e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && c.type === "formula_static") recomputeStatic(c.id); }}
            onBlur={() => { if (c.type === "formula_static") recomputeStatic(c.id); }}
            title="row formula — refs in [brackets]; a compact name expands to all its leaves"
            style={{ flex: 1, minWidth: 0, border: `1px solid ${PLAT.dark}`, background: "#fffef6", fontFamily: "monospace", fontSize: 9, padding: "0 2px" }} />
          <span onClick={() => openFormulaBuilder(c.id)} title="pick cells / build formula" style={{ cursor: "pointer", fontSize: 10, border: `1px solid ${PLAT.dark}`, borderRadius: 2, padding: "0 2px", background: "#fff" }}>⊞</span>
          {c.type === "formula_static" && <span onClick={() => recomputeStatic(c.id)} title="recompute frozen values" style={{ cursor: "pointer", fontSize: 10 }}>↻</span>}
        </div>
      )}
      {c.type === "real" && (
        <select value={c.decimals == null ? "" : String(c.decimals)} onClick={(e) => e.stopPropagation()} onChange={(e) => setColDecimals(c.id, e.target.value === "" ? null : parseInt(e.target.value, 10))} title="decimal places to display (display only — stored values are unchanged)" style={{ fontSize: 9, fontFamily: FONT, width: "100%", marginTop: 1, border: "none", background: "transparent", color: "#456" }}>
          <option value="">dec: auto</option>
          {[0, 1, 2, 3, 4, 5, 6].map((d) => <option key={d} value={d}>{d} dp</option>)}
        </select>
      )}
      {grip}
    </th>
  );
}
const thStyle = { border: `1px solid ${PLAT.dark}`, padding: "1px 2px", position: "sticky", top: 0, fontSize: 11 };
const tdStyle = { border: "1px solid #d8d8d8", padding: "0 2px", height: 17 };

/* =========================================================================
   COMPACT DIALOG
   ========================================================================= */
function CompactDialog({ dlg, setDlg, columns, onCreate }) {
  const set = (patch) => setDlg((d) => ({ ...d, ...patch }));
  const picks = dlg.picks;
  const factors = dlg.factors.map((f) => ({ name: f.name.trim(), levels: f.levels.split(",").map((s) => s.trim()).filter(Boolean) }));
  const product = factors.reduce((p, f) => p * Math.max(1, f.levels.length), 1);
  const ok = product === picks.length && factors.every((f) => f.name && f.levels.length >= 1) && factors.length >= 1;

  const build = () => {
    const compact = {
      id: "cp" + Math.random().toString(36).slice(2, 7),
      name: dlg.name.trim() || "Measure",
      factors, leaves: [...picks],
    };
    onCreate(compact);
  };
  const movePick = (i, dir) => set({ picks: (() => { const a = [...picks]; const j = i + dir; if (j < 0 || j >= a.length) return a; [a[i], a[j]] = [a[j], a[i]]; return a; })() });
  const setFactor = (i, patch) => set({ factors: dlg.factors.map((f, k) => (k === i ? { ...f, ...patch } : f)) });
  // Set the number of within factors: keep any rows already edited, fill the rest with the default name/levels format.
  const setFactorCount = (n) => set({ factors: Array.from({ length: n }, (_, k) => dlg.factors[k] || { name: "Factor" + (k + 1), levels: "level1, level2" }) });

  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 2000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ width: 470, background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11 }}>
        <div style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>Compact Columns → Repeated-Measures Variable</span>
        </div>
        <div style={{ padding: 12, background: "#f4f4f4" }}>
          <div style={{ marginBottom: 8 }}>
            <b>Measure name:</b>{" "}
            <input value={dlg.name} onChange={(e) => set({ name: e.target.value })} style={{ fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 4px", width: 180 }} />
            <span style={{ marginLeft: 10, color: "#456" }}>{picks.length} columns selected</span>
          </div>
          <div style={{ marginBottom: 8 }}>
            <b>Number of within factors:</b>{" "}
            <select value={dlg.factors.length} onChange={(e) => setFactorCount(parseInt(e.target.value, 10))} style={{ fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 4px" }}>
              {[1, 2, 3, 4].map((n) => <option key={n} value={n}>{n}</option>)}
            </select>
            <span style={{ marginLeft: 10, color: "#789", fontSize: 10 }}>rows are created as Factor1…FactorN with default levels; rename them and set their levels below</span>
          </div>

          <div style={{ display: "flex", gap: 12 }}>
            {/* factors */}
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: "bold", marginBottom: 3 }}>Within factors (outer → inner, up to 4)</div>
              {dlg.factors.map((f, i) => (
                <div key={i} style={{ display: "flex", gap: 4, marginBottom: 4, alignItems: "center" }}>
                  <span style={{ color: "#789", width: 14 }}>{i + 1}.</span>
                  <input value={f.name} onChange={(e) => setFactor(i, { name: e.target.value })} placeholder="factor" style={{ width: 70, fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 3px" }} />
                  <input value={f.levels} onChange={(e) => setFactor(i, { levels: e.target.value })} placeholder="lvlA, lvlB" style={{ flex: 1, fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 3px" }} />
                  {dlg.factors.length > 1 && <span onClick={() => set({ factors: dlg.factors.filter((_, k) => k !== i) })} style={{ cursor: "pointer", color: "#900" }}>✕</span>}
                </div>
              ))}
              {dlg.factors.length < 4 && <MacBtn onClick={() => set({ factors: [...dlg.factors, { name: "Factor" + (dlg.factors.length + 1), levels: "level1, level2" }] })}>+ Add factor</MacBtn>}
              <div style={{ marginTop: 8, color: ok ? "#161" : "#a00", fontSize: 10 }}>
                {factors.map((f) => `${f.name || "?"}(${f.levels.length})`).join(" × ")} = <b>{product}</b> cells {ok ? "✓ matches column count" : `≠ ${picks.length} columns`}
              </div>
            </div>

            {/* mapping preview */}
            <div style={{ width: 190 }}>
              <div style={{ fontWeight: "bold", marginBottom: 3 }}>Column → level map</div>
              <div style={{ ...bevelIn, background: "#fff", maxHeight: 150, overflowY: "auto", padding: 3 }}>
                {picks.map((id, i) => {
                  const col = columns.find((c) => c.id === id);
                  const lv = ok ? leafLevels({ factors, leaves: picks }, i) : null;
                  return (
                    <div key={id} style={{ display: "flex", alignItems: "center", gap: 3, padding: "1px 0", fontSize: 10 }}>
                      <span onClick={() => movePick(i, -1)} style={{ cursor: "pointer", color: "#678" }}>▲</span>
                      <span onClick={() => movePick(i, 1)} style={{ cursor: "pointer", color: "#678" }}>▼</span>
                      <span style={{ width: 54, fontFamily: "monospace" }}>{col?.name || id}</span>
                      <span style={{ color: "#a76" }}>→ {lv ? Object.entries(lv).map(([k, v]) => `${v}`).join(" · ") : "—"}</span>
                    </div>
                  );
                })}
              </div>
              <div style={{ fontSize: 9, color: "#789", marginTop: 3 }}>innermost factor varies fastest down the list. reorder with ▲▼.</div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 12 }}>
            <MacBtn onClick={() => setDlg(null)}>Cancel</MacBtn>
            <MacBtn onClick={build} disabled={!ok} primary={ok}>Compact</MacBtn>
          </div>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   ANALYSIS BROWSER
   ========================================================================= */
function AnalysisBrowser({ expanded, setExpanded, addAnalysis }) {
  const byCat = {};
  BROWSER_ITEMS.forEach((it) => (byCat[it.cat] = byCat[it.cat] || []).push(it));
  return (
    <div style={{ flex: 1, minHeight: 0, overflowY: "auto", background: "#fff", padding: "4px 2px" }}>
      {CATEGORY_ORDER.map((cat) => (
        <div key={cat}>
          <div onClick={() => setExpanded((e) => ({ ...e, [cat]: !e[cat] }))} style={{ display: "flex", alignItems: "center", gap: 4, cursor: "pointer", padding: "1px 4px", fontWeight: "bold" }}>
            <span style={{ fontSize: 9, transform: expanded[cat] ? "rotate(90deg)" : "none", display: "inline-block", transition: "transform .1s" }}>▶</span>
            {cat}
          </div>
          {expanded[cat] && (byCat[cat] || []).map((it, i) => (
            it.nested ? (
              <div key={i}>
                <div onClick={() => setExpanded((e) => ({ ...e, ["sub:" + it.label]: !e["sub:" + it.label] }))} style={{ padding: "1px 4px 1px 14px", cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                  onMouseEnter={(e) => { e.currentTarget.style.background = PLAT.sel; }} onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}>
                  <span style={{ fontSize: 9, transform: expanded["sub:" + it.label] ? "rotate(90deg)" : "none", display: "inline-block" }}>▶</span>{it.label}
                </div>
                {expanded["sub:" + it.label] && it.nested.map((sub, j) => (
                  <div key={j} onClick={() => addAnalysis(sub)} title="add post-hoc comparisons" style={{ padding: "1px 4px 1px 34px", cursor: "pointer" }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = PLAT.sel; }} onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}>
                    {sub.label}
                  </div>
                ))}
              </div>
            ) : (
              <div key={i} onClick={() => addAnalysis(it)} title="click to add to view" style={{ padding: "1px 4px 1px 22px", cursor: "pointer" }}
                onMouseEnter={(e) => { e.currentTarget.style.background = PLAT.sel; }} onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}>
                {it.label}
              </div>
            )
          ))}
        </div>
      ))}
    </div>
  );
}

/* =========================================================================
   VARIABLES BROWSER  (compact variables + loose columns)
   ========================================================================= */
function VariablesBrowser({ columns, compacts, leafToCompact, sel, setSel, def, accepts, assignRole, analysis, colById, compactById, expand, expandFactor, colSel, setColSel, openCompact, setCompactDecimals, onNeedHeight }) {
  const looseCols = columns.filter((c) => !leafToCompact[c.id]);
  const isSel = (kind, id) => sel && sel.kind === kind && sel.id === id;
  const looseOrder = looseCols.map((c) => c.id);
  const vAnchor = useRef(null);
  const colSelected = (id) => !!(colSel && colSel.has(id));
  const pickCol = (id, e) => {
    if (e && e.shiftKey && vAnchor.current && looseOrder.includes(vAnchor.current)) {
      const a = looseOrder.indexOf(vAnchor.current), b = looseOrder.indexOf(id);
      const [lo, hi] = a < b ? [a, b] : [b, a];
      setColSel(new Set(looseOrder.slice(lo, hi + 1))); setSel({ kind: "col", id });
    } else if (e && (e.metaKey || e.ctrlKey || e.altKey)) {
      const n = new Set(colSel || []); n.has(id) ? n.delete(id) : n.add(id); setColSel(n); vAnchor.current = id; setSel({ kind: "col", id });
    } else {
      setColSel(new Set([id])); vAnchor.current = id; setSel({ kind: "col", id });
    }
  };
  const compactEligible = looseCols.filter((c) => colSel && colSel.has(c.id) && (c.type === "real" || c.type === "integer"));
  useEffect(() => { if (sel && sel.kind === "compact" && onNeedHeight) onNeedHeight(); }, [sel, onNeedHeight]);
  const roleSummary = (r) => {
    if (!analysis) return "";
    const v = analysis.roles[r.key];
    if (!v) return "—";
    if (r.accept === "depOrCompact") return v.kind === "compact" ? (compactById[v.id]?.name || v.id) : (colById[v.id]?.name || v.id);
    if (r.accept === "compactOnly") return compactById[v.id]?.name || v.id;
    if (r.multiple) return (v || []).map((id) => colById[id]?.name || id).join(", ") || "—";
    return colById[v]?.name || v;
  };
  return (
    <div style={{ background: "#fff", flex: 1, minHeight: 0, display: "flex", flexDirection: "column" }}>
      <div style={{ flex: 1, minHeight: 110, overflowY: "auto", borderBottom: `1px solid ${PLAT.dark}`, padding: 2 }}>
        {compacts.map((c) => (
          <div key={c.id} onClick={() => { setSel({ kind: "compact", id: c.id }); setColSel(new Set()); }} style={{ padding: "1px 5px", cursor: "pointer", display: "flex", gap: 5, alignItems: "center", background: isSel("compact", c.id) ? PLAT.selBorder : "transparent", color: isSel("compact", c.id) ? "#fff" : COMPACT_CLR, fontWeight: "bold" }}>
            <span style={{ fontFamily: "monospace", fontSize: 9, color: "#fff", background: COMPACT_CLR, borderRadius: 2, padding: "0 2px" }}>R</span>
            {c.name} <span style={{ fontSize: 9, opacity: 0.8 }}>[{c.factors.map((f) => f.name).join("×")}]</span>
          </div>
        ))}
        {looseCols.map((c) => {
          const on = colSelected(c.id) || isSel("col", c.id);
          return (
          <div key={c.id} onClick={(e) => pickCol(c.id, e)} style={{ padding: "1px 5px", cursor: "pointer", display: "flex", gap: 5, alignItems: "center", background: on ? PLAT.selBorder : "transparent", color: on ? "#fff" : "#000" }}>
            {typeBadge(c.type)} {c.name}
          </div>
          );
        })}
      </div>
      {compactEligible.length >= 2 && (
        <div style={{ padding: "4px 6px", borderBottom: `1px solid ${PLAT.dark}`, background: "#eef4ff", display: "flex", alignItems: "center", gap: 6 }}>
          <button onClick={openCompact} title="combine the selected continuous columns into a repeated-measures variable" style={{ fontFamily: FONT, fontSize: 11, padding: "2px 10px", ...bevelOut, borderRadius: 6, background: PLAT.faceLite, cursor: "pointer", fontWeight: "bold", color: COMPACT_CLR }}>▣ Compact {compactEligible.length} selected…</button>
          <span style={{ fontSize: 9, color: "#667" }}>Shift- or ⌘/⌥-click to select multiple</span>
        </div>
      )}
      {sel && sel.kind === "compact" && compactById[sel.id] && (
        <div style={{ padding: "5px 6px", borderBottom: `1px solid ${PLAT.dark}`, background: "#fbf4ee" }}>
          <div style={{ fontWeight: "bold", color: COMPACT_CLR, marginBottom: 3 }}>Expand “{compactById[sel.id].name}”</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
            {compactById[sel.id].factors.map((f) => (
              <button key={f.name} onClick={() => expandFactor(sel.id, f.name)} title={`remove ${f.name} as a repeated factor`}
                style={{ fontFamily: FONT, fontSize: 10, padding: "1px 6px", ...bevelOut, borderRadius: 6, background: PLAT.faceLite, cursor: "pointer" }}>
                ⊟ {f.name} ({f.levels.length})
              </button>
            ))}
            <button onClick={() => expand(sel.id)} title="dissolve completely (all leaves become loose columns)"
              style={{ fontFamily: FONT, fontSize: 10, padding: "1px 6px", ...bevelOut, borderRadius: 6, background: PLAT.faceLite, cursor: "pointer", color: "#900" }}>⊟ all</button>
          </div>
          <div style={{ fontSize: 9, color: "#866", marginTop: 3 }}>removing a factor splits this into one variable per level of that factor</div>
          <div style={{ marginTop: 5, display: "flex", alignItems: "center", gap: 5, fontSize: 10, color: "#634" }}>
            Decimal places (all values):
            {(() => {
              const ds = compactById[sel.id].leaves.map((id) => colById[id] && colById[id].decimals);
              const common = ds.every((d) => d === ds[0]) ? ds[0] : undefined;
              return (
                <select value={common == null ? "" : String(common)} onChange={(e) => setCompactDecimals(sel.id, e.target.value === "" ? null : parseInt(e.target.value, 10))}
                  title="decimal places to display for every value in this compact variable (display only)"
                  style={{ fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" }}>
                  <option value="">auto{common === undefined ? " (mixed)" : ""}</option>
                  {[0, 1, 2, 3, 4, 5, 6].map((d) => <option key={d} value={d}>{d} dp</option>)}
                </select>
              );
            })()}
          </div>
        </div>
      )}
      <div style={{ padding: 6, display: "flex", flexDirection: "column", gap: 4, minHeight: 96 }}>
        {!analysis && <div style={{ color: "#888" }}>Select an analysis in the View, then assign variables to its roles.</div>}
        {analysis && def && def.roles.map((r) => {
          const ok = accepts(r.accept, sel);
          return (
            <button key={r.key} onClick={() => assignRole(r.key)} disabled={!ok} style={{ fontFamily: FONT, fontSize: 11, padding: "2px 8px", textAlign: "left", background: ok ? PLAT.faceLite : "#eee", ...bevelOut, borderRadius: 6, cursor: ok ? "pointer" : "default", color: ok ? "#000" : "#999" }}>
              → {r.label}: <b style={{ fontWeight: "normal", color: "#357" }}>{roleSummary(r)}</b>
            </button>
          );
        })}
        {analysis && def && <div style={{ color: "#888", fontSize: 10, marginTop: 2 }}>Pick a variable above, then click a role. A compact variable carries its own within factors.</div>}
      </div>
    </div>
  );
}

/* =========================================================================
   ANALYSIS PANE
   ========================================================================= */
// Group analyses so an ANOVA and its inherited follow-ups (means/plots/post-hoc) cluster
// together: same dependent + between-subjects roles => same group. Non-ANOVA = its own group.
function analysisSig(a) {
  if (!a || a.type !== "anova") return "id:" + (a ? a.id : "");
  const dep = a.roles && a.roles.dep;
  return "anova|" + (dep ? dep.kind + ":" + dep.id : "?") + "|" + (((a.roles && a.roles.between) || []).slice().sort().join(","));
}
function analysisNavLabel(a, colById, compactById) {
  if (a.type === "anova") {
    const dep = a.roles && a.roles.dep;
    const dn = dep ? (((dep.kind === "compact" ? compactById[dep.id] : colById[dep.id]) || {}).name || "?") : "(unset)";
    const btw = ((a.roles && a.roles.between) || []).map((id) => (colById[id] || {}).name || id);
    return "ANOVA \u00b7 " + dn + (btw.length ? " \u00d7 " + btw.join(" \u00d7 ") : "");
  }
  return (ANALYSES[a.type] || {}).name || a.type;
}
// Analysis View body: a docked navigator (resizable + collapsible) driving a filtered
// detail pane that shows just the selected analysis's cluster (or all, via "All results").
function AnalysisView({ analyses, selAnalysis, setSelAnalysis, colById, compactById, valuesOf, rows, removeAnalysis, onOutput, onConfig, onSwap, onMethod, onPhEffect, onPlot, onPlotResize, onPlotLegend, onOpt }) {
  const [navW, setNavW] = useState(196);
  const [navOpen, setNavOpen] = useState(true);
  const [viewAll, setViewAll] = useState(false);
  const [collapsed, setCollapsed] = useState(() => new Set());
  const dragRef = useRef(null);
  const detailRef = useRef(null);
  useEffect(() => {
    const mv = (e) => { if (dragRef.current) setNavW(Math.max(140, Math.min(380, dragRef.current.ow + e.clientX - dragRef.current.sx))); };
    const up = () => (dragRef.current = null);
    window.addEventListener("mousemove", mv); window.addEventListener("mouseup", up);
    return () => { window.removeEventListener("mousemove", mv); window.removeEventListener("mouseup", up); };
  }, []);
  useEffect(() => {
    if (viewAll) return;
    const c = detailRef.current; if (!c) return;
    const r = window.requestAnimationFrame(() => { const el = c.querySelector('[data-aid="' + selAnalysis + '"]'); if (el) c.scrollTop = Math.max(0, el.offsetTop - 8); });
    return () => window.cancelAnimationFrame(r);
  }, [selAnalysis, viewAll]);
  if (!analyses.length) return <div style={{ flex: 1, color: "#666", padding: 30, textAlign: "center", background: "#fff" }}>Empty view. Add an analysis from the Analysis Browser or the Analyze menu.</div>;
  const groups = []; const bySig = {};
  analyses.forEach((a) => { const sg = analysisSig(a); if (!bySig[sg]) { bySig[sg] = { sig: sg, members: [] }; groups.push(bySig[sg]); } bySig[sg].members.push(a); });
  groups.forEach((g) => { g.head = g.members.find((m) => m.type === "anova" && (m.output || "table") === "table") || g.members[0]; g.children = g.members.filter((m) => m !== g.head); });
  const selA = analyses.find((a) => a.id === selAnalysis);
  const selSig = selA ? analysisSig(selA) : groups[0].sig;
  const shown = viewAll ? analyses : analyses.filter((a) => analysisSig(a) === selSig);
  const toggleGroup = (sig) => setCollapsed((set) => { const n = new Set(set); n.has(sig) ? n.delete(sig) : n.add(sig); return n; });
  const navRow = (a, label, depth, hasKids, sig) => {
    const isSel = !viewAll && selAnalysis === a.id;
    const isCol = hasKids && collapsed.has(sig);
    return (
      <div key={a.id} onClick={() => { setViewAll(false); setSelAnalysis(a.id); }} title={label}
        style={{ display: "flex", alignItems: "center", padding: "2px 6px", paddingLeft: 4 + depth * 14, cursor: "pointer", background: isSel ? PLAT.selBorder : "transparent", color: isSel ? "#fff" : "#000" }}>
        <span onClick={hasKids ? (e) => { e.stopPropagation(); toggleGroup(sig); } : undefined} title={hasKids ? (isCol ? "expand" : "collapse") : undefined}
          style={{ width: 12, textAlign: "center", fontSize: 8, flexShrink: 0, cursor: hasKids ? "pointer" : "default", color: isSel ? "#fff" : "#555" }}>{hasKids ? (isCol ? "\u25b8" : "\u25be") : ""}</span>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</span>
      </div>
    );
  };
  return (
    <div style={{ flex: 1, minHeight: 0, display: "flex", overflow: "hidden", background: PLAT.face }}>
      {navOpen ? (
        <div style={{ width: navW, flexShrink: 0, display: "flex", flexDirection: "column", background: "#ededed", overflow: "hidden" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "2px 5px 2px 7px", borderBottom: `1px solid ${PLAT.dark}`, fontSize: 10, fontWeight: "bold", color: "#555", background: PLAT.faceLite, flexShrink: 0 }}>
            <span>Analyses</span>
            <span onClick={() => setNavOpen(false)} title="hide navigator" style={{ cursor: "pointer", padding: "0 3px" }}>{"\u2039"}</span>
          </div>
          <div style={{ overflow: "auto", flex: 1, fontSize: 11, padding: "2px 0" }}>
            <div onClick={() => setViewAll(true)} style={{ padding: "2px 6px 2px 21px", cursor: "pointer", fontStyle: "italic", background: viewAll ? PLAT.selBorder : "transparent", color: viewAll ? "#fff" : "#444" }}>All results ({analyses.length})</div>
            {groups.map((g) => (
              <div key={g.sig}>
                {navRow(g.head, analysisNavLabel(g.head, colById, compactById), 0, g.children.length > 0, g.sig)}
                {!collapsed.has(g.sig) && g.children.map((c) => navRow(c, ANOVA_OUTPUTS[c.output || "table"], 1, false, g.sig))}
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div onClick={() => setNavOpen(true)} title="show navigator" style={{ width: 15, flexShrink: 0, cursor: "pointer", background: "#ededed", borderRight: `1px solid ${PLAT.dark}`, display: "flex", justifyContent: "center", paddingTop: 4, fontSize: 11, color: "#555" }}>{"\u203a"}</div>
      )}
      {navOpen && <div onMouseDown={(e) => { dragRef.current = { sx: e.clientX, ow: navW }; e.preventDefault(); }} title="drag to resize" style={{ width: 4, flexShrink: 0, cursor: "col-resize", background: PLAT.face, borderLeft: `1px solid ${PLAT.faceLite}`, borderRight: `1px solid ${PLAT.dark}` }} />}
      <div ref={detailRef} style={{ flex: 1, minWidth: 0, overflow: "auto", background: "#fff", padding: 8, position: "relative" }}>
        {shown.map((a) => (
          <div key={a.id} data-aid={a.id}>
            <AnalysisPane a={a} colById={colById} compactById={compactById} valuesOf={valuesOf} rows={rows} selected={selAnalysis === a.id} onSelect={() => { setViewAll(false); setSelAnalysis(a.id); }} onRemove={() => removeAnalysis(a.id)} onOutput={onOutput} onConfig={onConfig} onSwap={onSwap} onMethod={onMethod} onPhEffect={onPhEffect} onPlot={onPlot} onPlotResize={onPlotResize} onPlotLegend={onPlotLegend} onOpt={onOpt} />
          </div>
        ))}
      </div>
    </div>
  );
}
function AnalysisPane({ a, colById, compactById, valuesOf, rows, selected, onSelect, onRemove, onOutput, onConfig, onSwap, onMethod, onPhEffect, onPlot, onPlotResize, onPlotLegend, onOpt }) {
  const def = ANALYSES[a.type];
  const body = useMemo(() => renderResult(a, def, colById, compactById, valuesOf, rows, onPlotResize, onPlotLegend, onOpt), [a, def, colById, compactById, valuesOf, rows, onPlotResize, onPlotLegend, onOpt]);
  const title = a.type === "anova" ? ANOVA_OUTPUTS[a.output || "table"] : def.name;
  const phEffects = (() => {
    if (a.type !== "anova" || !a.roles.dep) return [];
    const within = a.roles.dep.kind === "compact" ? (compactById[a.roles.dep.id]?.factors || []).map((f) => f.name) : [];
    const between = (a.roles.between || []).map((id) => colById[id]?.name || id);
    return effectList([...between, ...within], "all");
  })();
  const effKey = (e) => e.join("\u0001");
  const _singles = phEffects.filter((e) => e.length === 1);
  const curEff = a.phEffect || (_singles.length ? effKey(_singles[_singles.length - 1]) : (phEffects[0] ? effKey(phEffects[0]) : ""));
  return (
    <div onClick={onSelect} style={{ ...bevelIn, background: "#fff", marginBottom: 10, outline: selected ? `2px solid ${PLAT.selBorder}` : "none" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: selected ? STRIPES : PLAT.face, borderBottom: `1px solid ${PLAT.dark}`, padding: "2px 6px", fontWeight: "bold", fontSize: 11 }}>
        <span style={{ background: selected ? PLAT.face : "transparent", padding: "0 4px" }}>{title}</span>
        <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
          {a.type === "anova" && (
            <>
              {a.output === "posthoc" && (
                <>
                  <select value={a.method || "tukey"} onClick={(e) => e.stopPropagation()} onChange={(e) => onMethod(a.id, e.target.value)} title="correction method" style={{ fontFamily: FONT, fontSize: 10, border: "none", background: "transparent" }}>
                    {Object.entries(PH_METHODS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                  </select>
                  {phEffects.length > 0 && (
                    <select value={curEff} onClick={(e) => e.stopPropagation()} onChange={(e) => onPhEffect(a.id, e.target.value)} title="effect to compare (main effect or interaction)" style={{ fontFamily: FONT, fontSize: 10, border: "none", background: "transparent" }}>
                      {phEffects.map((e) => <option key={effKey(e)} value={effKey(e)}>{e.join(" × ")}</option>)}
                    </select>
                  )}
                </>
              )}
              <select value={a.output || "table"} onClick={(e) => e.stopPropagation()} onChange={(e) => onOutput(a.id, e.target.value)} style={{ fontFamily: FONT, fontSize: 10, border: "none", background: "transparent" }}>
                {Object.entries(ANOVA_OUTPUTS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
              {(a.output === "bar" || a.output === "line") && (
                <span onClick={(e) => { e.stopPropagation(); onSwap(a.id); }} style={{ cursor: "pointer" }} title="swap x-axis factor and side-by-side factor">⇄</span>
              )}
              {(a.output === "bar" || a.output === "line") && (
                <span onClick={(e) => { e.stopPropagation(); onPlot(a.id); }} style={{ cursor: "pointer" }} title="plot settings (axes, symbols, grid)">🎨</span>
              )}
              <span onClick={(e) => { e.stopPropagation(); onConfig(a.id); }} style={{ cursor: "pointer" }} title="analysis setup">⚙</span>
            </>
          )}
          <span onClick={(e) => { e.stopPropagation(); onRemove(); }} style={{ cursor: "pointer", color: "#900" }} title="remove">✕</span>
        </span>
      </div>
      <div style={{ padding: 8 }}>{body}</div>
    </div>
  );
}

const fmt = (v, d = 3) => (v === null || v === undefined || Number.isNaN(v) ? "—" : Math.abs(v) >= 1e5 || (Math.abs(v) < 1e-3 && v !== 0) ? v.toExponential(2) : v.toFixed(d));
const fmtP = (p) => (p === null || Number.isNaN(p) ? "—" : p < 0.0001 ? "<.0001" : p.toFixed(4));

function needRoles(a, def) {
  return def.roles.every((r) => { if (r.optional) return true; const v = a.roles[r.key]; return r.multiple ? v && v.length > 0 : !!v; });
}

const _POWER_FAM = [["t2", "Two independent means (t)"], ["t1", "Paired / one-sample mean (t)"], ["anova", "One-way ANOVA (F)"], ["corr", "Correlation (r)"], ["prop", "Two proportions (z)"], ["chi2", "Chi-square / goodness-of-fit (w)"]];
const _POWER_ESNAME = { t2: "Cohen's d", t1: "Cohen's d", anova: "Cohen's f", corr: "r", prop: "Cohen's h", chi2: "w" };
const _POWER_ESHINT = { t2: "d: 0.2 small · 0.5 medium · 0.8 large", t1: "d: 0.2 small · 0.5 medium · 0.8 large", anova: "f: 0.10 small · 0.25 medium · 0.40 large", corr: "r: 0.10 small · 0.30 medium · 0.50 large", prop: "h: 0.20 small · 0.50 medium · 0.80 large", chi2: "w: 0.10 small · 0.30 medium · 0.50 large" };
const _POWER_NOTE = {
  t2: "Independent-samples t (equal n per group); noncentrality δ = d·√(n/2), df = 2n−2.",
  t1: "Paired / one-sample t; noncentrality δ = d·√N, df = N−1.",
  anova: "One-way fixed-effects ANOVA; noncentral F with λ = f²·N, df = (k−1, N−k).",
  corr: "Pearson correlation; Fisher-z approximation with bias correction (pwr convention).",
  prop: "Two independent proportions; arcsine effect size h, normal (z) approximation.",
  chi2: "χ² goodness-of-fit / association; noncentral χ² with λ = w²·N at the given df.",
};
function PowerPanel({ a, onOpt }) {
  const D = { family: "t2", solveFor: "n", tail: "two", alpha: 0.05, es: 0.5, n: 30, k: 3, df: 1, power: 0.8 };
  const o = { ...D, ...(a.opts || {}) };
  const set = (patch) => onOpt(a.id, patch);
  const fam = o.family;
  const hasTail = fam === "t2" || fam === "t1" || fam === "corr" || fam === "prop";
  const totalFam = fam === "corr" || fam === "chi2";
  const nLabel = totalFam ? "Sample size N" : (fam === "t1" ? "N (pairs)" : "n per group");
  const r = solvePower(o);
  const SELS = { fontFamily: FONT, fontSize: 11, border: `1px solid ${PLAT.dark}`, background: "#fff", padding: "1px 2px" };
  const INP = { width: 66, fontFamily: FONT, fontSize: 11, border: `1px solid ${PLAT.dark}`, padding: "1px 3px" };
  const field = (key, label, step) => (
    <label style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>{label}:
      <input type="number" step={step || "any"} value={o[key]} onChange={(e) => set({ [key]: e.target.value })} style={INP} /></label>
  );
  const f4 = (x) => (isFinite(x) ? x.toFixed(4) : "\u2014");
  const tailCtl = hasTail ? (
    <label>Tails: <select value={o.tail} onChange={(e) => set({ tail: e.target.value })} style={SELS}>
      <option value="two">two</option>
      <option value="one">one</option>
    </select></label>
  ) : null;
  let headline;
  if (o.solveFor === "power") headline = `Power = ${f4(r.power)}`;
  else if (o.solveFor === "n") headline = !isFinite(r.n) ? "Required sample exceeds 1,000,000" :
    `Required ${nLabel} = ${r.n}` + (fam === "anova" ? `  (total N = ${r.n * r.k})` : (fam === "t2" || fam === "prop") ? `  (total N = ${r.n * 2})` : "") + `  ·  achieved power ${f4(r.power)}`;
  else headline = !isFinite(r.es) ? "Target power not attainable at this sample size" : `Minimum detectable ${_POWER_ESNAME[fam]} = ${r.es.toFixed(4)}`;
  const data = r.curve.points.map((p) => ({ n: p.n, power: p.power }));
  return (
    <div style={{ fontFamily: FONT, fontSize: 11, maxWidth: 640 }}>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center", marginBottom: 6 }}>
        <label>Test: <select value={fam} onChange={(e) => set({ family: e.target.value })} style={SELS}>{_POWER_FAM.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></label>
        <label>Solve for: <select value={o.solveFor} onChange={(e) => set({ solveFor: e.target.value })} style={SELS}><option value="power">Power (post hoc)</option><option value="n">Sample size (a priori)</option><option value="es">Min. detectable effect</option></select></label>
        {tailCtl}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", marginBottom: 6 }}>
        {field("alpha", "\u03b1", "0.01")}
        {o.solveFor !== "es" && field("es", _POWER_ESNAME[fam], "0.05")}
        {o.solveFor !== "n" && field("n", nLabel, "1")}
        {fam === "anova" && field("k", "groups (k)", "1")}
        {fam === "chi2" && field("df", "df", "1")}
        {o.solveFor !== "power" && field("power", "target power", "0.05")}
      </div>
      <div style={{ background: "#eef3fb", border: "1px solid #aac3e6", borderRadius: 4, padding: "8px 10px", marginBottom: 6 }}>
        <div style={{ fontSize: 14, fontWeight: "bold", color: "#1a3b6e" }}>{headline}</div>
        <div style={{ fontSize: 10, color: "#556", marginTop: 2 }}>{_POWER_ESHINT[fam]}</div>
      </div>
      <div style={{ height: 200 }}>
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={data} margin={{ top: 6, right: 16, bottom: 18, left: -4 }}>
            <CartesianGrid stroke="#eee" />
            <XAxis dataKey="n" type="number" domain={["dataMin", "dataMax"]} tick={{ fontSize: 9, fontFamily: FONT }} label={{ value: r.curve.xLabel, position: "insideBottom", offset: -8, fontSize: 10 }} />
            <YAxis domain={[0, 1]} tick={{ fontSize: 9, fontFamily: FONT }} label={{ value: "power", angle: -90, position: "insideLeft", fontSize: 10 }} />
            <Tooltip contentStyle={{ fontFamily: FONT, fontSize: 11 }} formatter={(v) => Number(v).toFixed(4)} />
            <Line type="monotone" dataKey="power" stroke="#1a3b6e" strokeWidth={2} dot={false} isAnimationActive={false} />
            {r.curve.opN != null && isFinite(r.curve.opPow) && <Scatter data={[{ n: r.curve.opN, power: r.curve.opPow }]} dataKey="power" fill="#c0653a" />}
          </ComposedChart>
        </ResponsiveContainer>
      </div>
      <div style={{ fontSize: 10, color: "#555", marginTop: 4 }}>
        {_POWER_NOTE[fam]} Curve: power vs {r.curve.xLabel}{r.curve.target != null ? ` (target ${(+o.power).toFixed(2)})` : ""}; the point marks the current scenario.
      </div>
    </div>
  );
}

function renderResult(a, def, colById, compactById, valuesOf, rows, onPlotResize, onPlotLegend, onOpt) {
  if (a.type === "anova") {
    const out = a.output || "table";
    if (out === "table") return renderAnovaTable(a, colById, compactById, rows, onOpt);
    if (out === "means") return renderAnovaMeans(a, colById, compactById, rows);
    if (out === "posthoc") return renderAnovaPosthoc(a, colById, compactById, rows);
    return renderAnovaGraph(a, colById, compactById, rows, out, onPlotResize, onPlotLegend); // "bar" | "line"
  }
  if (a.type === "power") return <PowerPanel a={a} onOpt={onOpt} />;
  if (a.type === "importnotes") {
    const info = a.info || {};
    return (
      <div style={{ fontFamily: FONT, fontSize: 11, maxWidth: 660, lineHeight: 1.45 }}>
        <div style={{ fontWeight: "bold", fontSize: 13, marginBottom: 2 }}>Imported from {info.source || "GraphPad Prism"}</div>
        <div style={{ color: "#556", marginBottom: 6 }}>{[info.createdBy, info.prismVer ? "PrismXMLVersion " + info.prismVer : ""].filter(Boolean).join("  \u00b7  ")}</div>
        {info.activeTable && (
          <div style={{ marginBottom: 6 }}>Loaded table: <b>{info.activeTable.title}</b> — {info.activeTable.typeDesc} <span style={{ color: "#667" }}>({info.activeTable.shape})</span></div>
        )}
        <div style={{ background: "#fff6e0", border: "1px solid #e3b341", borderRadius: 4, padding: "7px 9px", margin: "6px 0" }}>
          <b>Between- vs within-subjects:</b> Prism files do not record which factors are repeated measures, so values are imported with their original group/category labels. To treat a factor as <i>within-subjects</i> (repeated), tick its columns and choose <b>Compact</b>; treat independent factors as <i>between-subjects</i> grouping variables.
        </div>
        {(info.infoSheets || []).map((s, i) => (
          <div key={i} style={{ margin: "8px 0" }}>
            <div style={{ fontWeight: "bold" }}>{s.title}</div>
            {s.notes ? <div style={{ whiteSpace: "pre-wrap", color: "#334", margin: "2px 0 4px" }}>{s.notes}</div> : null}
            {s.constants && s.constants.length ? <StatTable head={["Constant", "Value"]} rows={s.constants.map((c) => [c.name, c.value || "\u2014"])} /> : null}
          </div>
        ))}
        {info.otherTables && info.otherTables.length ? (
          <div style={{ margin: "8px 0" }}>
            <div style={{ fontWeight: "bold" }}>Other tables in this file (not loaded)</div>
            <StatTable head={["Table", "Type"]} rows={info.otherTables.map((t) => [t.title, t.typeDesc])} />
            <div style={{ color: "#667", fontSize: 10, marginTop: 3 }}>VibeStat loaded the selected/first data table. To bring in another, make it the active table in Prism and re-save, then import again.</div>
          </div>
        ) : null}
      </div>
    );
  }
  if (!CMP_TYPES.has(a.type) && !DIST_TYPES.has(a.type) && !needRoles(a, def)) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign {def.roles.map((r) => r.label).join(", ")} in the Variables window.</div>;

  if (a.type === "descriptive") {
    const stats = ["n", "mean", "sd", "sem", "min", "max", "range", "median", "sum"];
    const labels = { n: "Count", mean: "Mean", sd: "Std. Dev.", sem: "Std. Error", min: "Minimum", max: "Maximum", range: "Range", median: "Median", sum: "Sum" };
    const cols = a.roles.y.map((id) => ({ id, name: colById[id]?.name || id, d: describe(valuesOf(id)) }));
    return <StatTable head={["", ...cols.map((c) => c.name)]} rows={stats.map((s) => [labels[s], ...cols.map((c) => s === "n" ? c.d.n : fmt(c.d[s]))])} />;
  }
  if (a.type === "histogram") {
    const h = histogram(valuesOf(a.roles.x));
    return (
      <div>
        <div style={{ fontSize: 11, marginBottom: 4 }}>{colById[a.roles.x]?.name} — n = {h.n}</div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={h.bins} margin={{ top: 4, right: 10, bottom: 4, left: -18 }}>
            <CartesianGrid stroke="#eee" /><XAxis dataKey="label" tick={{ fontSize: 9, fontFamily: FONT }} /><YAxis tick={{ fontSize: 9, fontFamily: FONT }} allowDecimals={false} />
            <Tooltip contentStyle={{ fontFamily: FONT, fontSize: 11 }} /><Bar dataKey="count" fill="#5577aa" stroke="#223355" />
          </BarChart>
        </ResponsiveContainer>
      </div>
    );
  }
  if (a.type === "scattergram") {
    const reg = regression(valuesOf(a.roles.x), valuesOf(a.roles.y));
    if (!reg.pairs) return <Few n={reg.n} />;
    const pts = reg.pairs.map(([x, y]) => ({ x, y }));
    const xs = pts.map((p) => p.x); const lo = Math.min(...xs), hi = Math.max(...xs);
    const lineData = [{ x: lo, y: reg.intercept + reg.slope * lo }, { x: hi, y: reg.intercept + reg.slope * hi }];
    return (
      <div>
        <div style={{ fontSize: 11, marginBottom: 4 }}>{colById[a.roles.y]?.name} vs {colById[a.roles.x]?.name} &nbsp; (R² = {fmt(reg.r2)})</div>
        <ResponsiveContainer width="100%" height={210}>
          <ComposedChart margin={{ top: 6, right: 12, bottom: 4, left: -18 }}>
            <CartesianGrid stroke="#eee" /><XAxis type="number" dataKey="x" domain={["auto", "auto"]} tick={{ fontSize: 9, fontFamily: FONT }} /><YAxis type="number" dataKey="y" domain={["auto", "auto"]} tick={{ fontSize: 9, fontFamily: FONT }} />
            <Tooltip contentStyle={{ fontFamily: FONT, fontSize: 11 }} /><Scatter data={pts} fill="#aa3344" />
            <Line data={lineData} dataKey="y" stroke="#223355" dot={false} strokeWidth={2} isAnimationActive={false} legendType="none" />
          </ComposedChart>
        </ResponsiveContainer>
      </div>
    );
  }
  if (a.type === "regression") {
    const r = regression(valuesOf(a.roles.x), valuesOf(a.roles.y));
    if (!r.pairs) return <Few n={r.n} />;
    const X = colById[a.roles.x]?.name, Y = colById[a.roles.y]?.name;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <StatTable head={["", "Value"]} rows={[["Count", r.n], ["R", fmt(r.r)], ["R Squared", fmt(r.r2)], ["Adjusted R²", fmt(r.adjR2)], ["RMS Residual", fmt(r.rmsResidual)]]} />
        <div style={{ fontWeight: "bold" }}>ANOVA</div>
        <StatTable head={["Source", "DF", "Sum Sq", "Mean Sq", "F", "P"]} rows={[["Regression", r.dfReg, fmt(r.ssReg, 2), fmt(r.msReg, 2), fmt(r.F, 3), fmtP(r.pF)], ["Residual", r.dfRes, fmt(r.ssRes, 2), fmt(r.msRes, 2), "", ""], ["Total", r.dfReg + r.dfRes, fmt(r.ssTotal, 2), "", "", ""]]} />
        <div style={{ fontWeight: "bold" }}>Coefficients</div>
        <StatTable head={["Term", "Coeff.", "Std. Err", "t", "P"]} rows={[["Intercept", fmt(r.intercept), "", "", ""], [X + " (slope)", fmt(r.slope), fmt(r.seSlope), fmt(r.tSlope), fmtP(r.pSlope)]]} />
        <div style={{ fontStyle: "italic" }}>{Y} = {fmt(r.intercept, 2)} {r.slope >= 0 ? "+" : "−"} {fmt(Math.abs(r.slope), 3)} · {X}</div>
      </div>
    );
  }
  if (a.type === "mreg") {
    const o = { ci: 0.95, showStd: true, ...(a.opts || {}) };
    const Y = colById[a.roles.y]?.name || a.roles.y;
    const xids = a.roles.x || [];
    const xnames = xids.map((id) => colById[id]?.name || id);
    const optsStrip = (
      <div style={{ display: "flex", gap: 14, fontSize: 10, color: "#444", alignItems: "center" }}>
        <label onClick={(e) => e.stopPropagation()}>CI level:
          <select value={String(o.ci)} onChange={(e) => onOpt(a.id, { ci: parseFloat(e.target.value) })} style={{ fontFamily: FONT, fontSize: 10, marginLeft: 4, border: `1px solid ${PLAT.dark}`, background: "#fff" }}>
            <option value="0.9">90%</option><option value="0.95">95%</option><option value="0.99">99%</option>
          </select>
        </label>
        <label onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={!!o.showStd} onChange={(e) => onOpt(a.id, { showStd: e.target.checked })} /> standardized β</label>
        <label onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={!!o.diag} onChange={(e) => onOpt(a.id, { diag: e.target.checked })} /> residual diagnostics</label>
      </div>
    );
    if (!a.roles.y || xids.length === 0) return (<div>{optsStrip}<div style={{ color: "#a00", marginTop: 8 }}>Assign a dependent variable (Y) and one or more predictors (X).</div></div>);
    const r = multipleRegression(valuesOf(a.roles.y), xids.map((id) => valuesOf(id)), xnames, o);
    if (r.error) return (<div>{optsStrip}<div style={{ color: "#a00", marginTop: 8 }}>{r.error}</div></div>);
    const pct = Math.round(o.ci * 100);
    const coefHead = o.showStd ? ["Term", "Coeff.", "Std. Err", `${pct}% CI`, "Std. β", "t", "P"] : ["Term", "Coeff.", "Std. Err", `${pct}% CI`, "t", "P"];
    const coefRows = r.coefs.map((c) => { const base = [c.name, fmt(c.b, 4), fmt(c.se, 4), `[${fmt(c.lo, 3)}, ${fmt(c.hi, 3)}]`], tail = [fmt(c.t, 3), fmtP(c.p)]; return o.showStd ? [...base, c.std == null ? "—" : fmt(c.std, 4), ...tail] : [...base, ...tail]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {optsStrip}
        <div style={{ fontSize: 11 }}>{Y} ~ {xnames.join(" + ")} &nbsp; (n = {r.n})</div>
        <StatTable head={["", "Value"]} rows={[["R", fmt(Math.sqrt(r.r2), 4)], ["R Squared", fmt(r.r2, 4)], ["Adjusted R²", fmt(r.adjR2, 4)], ["RMS Residual", fmt(r.rmse, 4)]]} />
        <div style={{ fontWeight: "bold" }}>ANOVA</div>
        <StatTable head={["Source", "DF", "Sum Sq", "Mean Sq", "F", "P"]} rows={[["Regression", r.dfM, fmt(r.ssr, 2), fmt(r.msr, 2), fmt(r.F, 3), fmtP(r.pF)], ["Residual", r.dfR, fmt(r.sse, 2), fmt(r.mse, 2), "", ""], ["Total", r.dfT, fmt(r.sst, 2), "", "", ""]]} />
        <div style={{ fontWeight: "bold" }}>Coefficients</div>
        <StatTable head={coefHead} rows={coefRows} />
        <div style={{ fontSize: 10, color: "#555" }}>OLS with listwise deletion across Y and all predictors. Std. β are fully standardized (z-scored X and Y); the intercept has none. CIs and P use the t-distribution on {r.dfR} residual df.</div>
        {o.diag && <ResidualDiagnostics resid={r.resid} fitted={r.fitted} />}
      </div>
    );
  }
  if (a.type === "wcontrast") {
    const m = anovaModel(a, colById, compactById, rows);
    if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
    if (!a.roles.dep || a.roles.dep.kind !== "compact" || m.withinNames.length === 0) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign a repeated-measures (R / compact) variable with at least one within-subjects factor.</div>;
    const o = a.opts || {};
    const target = (o.factor && m.withinNames.includes(o.factor)) ? o.factor : m.withinNames[0];
    const tLevels = m.levelOrder[target], k = tLevels.length;
    const wc = withinContrastScores(m.long, target, tLevels), n = wc.subjects.length;
    if (k < 2) return <div style={{ color: "#a00" }}>The factor needs at least two levels.</div>;
    if (n < 2) return <div style={{ color: "#a00" }}>Not enough subjects with complete data.</div>;
    const mode = o.mode === "custom" ? "custom" : "trend";
    const strip = (
      <div style={{ display: "flex", gap: 12, fontSize: 11, flexWrap: "wrap", alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
        <span style={{ fontWeight: "bold" }}>Mode:</span>
        <label><input type="radio" checked={mode === "trend"} onChange={() => onOpt(a.id, { mode: "trend" })} /> polynomial trend</label>
        <label><input type="radio" checked={mode === "custom"} onChange={() => onOpt(a.id, { mode: "custom" })} /> custom contrasts</label>
        {m.withinNames.length > 1 && <span>· factor: <select value={target} onChange={(e) => onOpt(a.id, { factor: e.target.value })} style={{ fontFamily: FONT, fontSize: 11 }}>{m.withinNames.map((f) => <option key={f} value={f}>{f}</option>)}</select></span>}
      </div>
    );
    const ctx = <div style={{ fontSize: 10, color: "#555" }}>n subjects = {n} · factor “{target}” — levels: {tLevels.join(", ")}{m.withinNames.length > 1 ? " (marginalized over other within factors)" : ""}. Each contrast uses its own error term (df = {n - 1}); no sphericity assumption.</div>;
    if (mode === "trend") {
      let scores = tLevels.map((l, i) => isFinite(Number(l)) ? Number(l) : i + 1);
      if (Array.isArray(o.scores) && o.scores.length === k && o.scores.every((s) => isFinite(Number(s)))) scores = o.scores.map(Number);
      const polys = orthoPoly(scores), names = ["Linear", "Quadratic", "Cubic", "Quartic", "Quintic", "Sextic", "Septic"];
      let ssSum = 0;
      const rrows = polys.map((p, i) => { const r = withinContrastTest(p, wc.cells); ssSum += r.ss; return [names[i] || ("Order " + (i + 1)), fmt(r.ss, 4), 1, r.df, fmt(r.F, 4), fmtP(r.p)]; });
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {strip}{ctx}
          <div style={{ fontWeight: "bold" }}>Within-subjects polynomial trend <span style={{ fontWeight: "normal", fontSize: 10, color: "#555" }}>(scores: {scores.join(", ")})</span></div>
          <StatTable head={["Component", "SS", "df", "df error", "F", "P"]} rows={rrows} />
          <div style={{ fontSize: 10, color: "#555" }}>Orthonormal polynomial contrasts on per-subject scores; each F has (1, {n - 1}) df. Components partition the {target} main-effect SS exactly (Σ = {fmt(ssSum, 3)}). Scores assume equal spacing by default.</div>
        </div>
      );
    }
    let C = (Array.isArray(o.contrasts) ? o.contrasts : []).map((row) => Array.from({ length: k }, (_, j) => (row && row[j] != null) ? String(row[j]) : "0"));
    if (C.length === 0) C = [Array.from({ length: k }, (_, j) => j === 0 ? "1" : (j === 1 ? "-1" : "0"))];
    const setC = (next) => onOpt(a.id, { contrasts: next });
    const cell = { width: 46, fontSize: 11, textAlign: "center" };
    const resultRows = C.map((row, ri) => { const cv = row.map((v) => Number(v) || 0), sum = cv.reduce((s, v) => s + v, 0); if (Math.abs(sum) < 1e-9) { const r = withinContrastTest(cv, wc.cells); return ["C" + (ri + 1), fmt(r.L, 4), fmt(r.se, 4), fmt(r.t, 4), r.df, fmtP(r.p), fmt(r.t / Math.sqrt(r.t * r.t + r.df), 4)]; } return ["C" + (ri + 1), "Σc = " + fmt(sum, 3) + " ≠ 0 (not a valid contrast)", "", "", "", "", ""]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {strip}{ctx}
        <div style={{ fontWeight: "bold" }}>Contrast coefficients <span style={{ fontWeight: "normal", fontSize: 10, color: "#555" }}>(columns = levels of {target})</span></div>
        <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
          <thead><tr><th style={{ padding: "2px 6px" }}></th>{tLevels.map((l) => <th key={l} style={{ padding: "2px 6px", borderBottom: "1px solid #999" }}>{l}</th>)}<th style={{ padding: "2px 6px" }}>Σ</th><th /></tr></thead>
          <tbody>{C.map((row, ri) => { const sum = row.reduce((s, v) => s + (Number(v) || 0), 0); return (
            <tr key={ri}><td style={{ padding: "2px 6px", fontWeight: "bold" }}>C{ri + 1}</td>
              {row.map((v, ci) => <td key={ci} style={{ padding: "1px 2px" }}><input type="text" value={v} style={cell} onClick={(e) => e.stopPropagation()} onChange={(e) => { const next = C.map((r) => r.slice()); next[ri][ci] = e.target.value; setC(next); }} /></td>)}
              <td style={{ padding: "2px 6px", color: Math.abs(sum) < 1e-9 ? "#070" : "#c00", fontWeight: "bold" }}>{fmt(sum, 2)}</td>
              <td style={{ padding: "2px 6px" }}>{C.length > 1 && <span style={{ cursor: "pointer", color: "#a00" }} onClick={(e) => { e.stopPropagation(); setC(C.filter((_, i) => i !== ri)); }}>✕</span>}</td>
            </tr>); })}</tbody>
        </table>
        <div><span style={{ cursor: "pointer", fontSize: 11, border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", borderRadius: 2, ...bevelOut }} onClick={(e) => { e.stopPropagation(); setC([...C, new Array(k).fill("0")]); }}>+ add contrast</span></div>
        <div style={{ fontWeight: "bold" }}>Contrast tests</div>
        <StatTable head={["Contrast", "Estimate", "SE", "t", "df", "P", "r"]} rows={resultRows} />
        <div style={{ fontSize: 10, color: "#555" }}>L = Σ cᵢ·(subject mean at level i), tested against zero with a one-sample t (df = {n - 1}). Coefficients should sum to zero.</div>
      </div>
    );
  }
  if (a.type === "contrast") {
    const yRaw = valuesOf(a.roles.y), gRaw = valuesOf(a.roles.factor), gmap = new Map();
    for (let i = 0; i < yRaw.length; i++) { const yy = num(yRaw[i]), g = gRaw[i]; if (yy != null && g !== "" && g != null) { const key = String(g); if (!gmap.has(key)) gmap.set(key, []); gmap.get(key).push(yy); } }
    let levels = [...gmap.keys()];
    const allNum = levels.length > 0 && levels.every((l) => isFinite(Number(l)));
    levels.sort(allNum ? (x, y) => Number(x) - Number(y) : (x, y) => x.localeCompare(y));
    const k = levels.length;
    if (k < 2) return <div style={{ color: "#a00" }}>Need at least two non-empty groups.</div>;
    const groups = levels.map((l) => gmap.get(l)), ow = oneWayGroups(groups), o = a.opts || {};
    const mode = o.mode === "custom" ? "custom" : "trend";
    const strip = (
      <div style={{ display: "flex", gap: 12, fontSize: 11, flexWrap: "wrap", alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
        <span style={{ fontWeight: "bold" }}>Mode:</span>
        <label><input type="radio" checked={mode === "trend"} onChange={() => onOpt(a.id, { mode: "trend" })} /> polynomial trend</label>
        <label><input type="radio" checked={mode === "custom"} onChange={() => onOpt(a.id, { mode: "custom" })} /> custom contrasts</label>
      </div>
    );
    const omnibus = <StatTable head={["Source", "df", "SS", "MS", "F", "P"]} rows={[["Between (factor)", ow.dfb, fmt(ow.ssb, 4), fmt(ow.ssb / ow.dfb, 4), fmt(ow.F, 4), fmtP(ow.p)], ["Within (error)", ow.dfw, fmt(ow.ssw, 4), fmt(ow.mse, 4), "", ""]]} />;
    const levelHdr = <div style={{ fontSize: 10, color: "#555" }}>Levels (in order): {levels.join(", ")} &nbsp;·&nbsp; n = {ow.ni.join(", ")}</div>;
    if (mode === "trend") {
      let scores = levels.map((l, i) => allNum ? Number(l) : i + 1);
      if (Array.isArray(o.scores) && o.scores.length === k && o.scores.every((s) => isFinite(Number(s)))) scores = o.scores.map(Number);
      const polys = orthoPoly(scores), names = ["Linear", "Quadratic", "Cubic", "Quartic", "Quintic", "Sextic", "Septic"];
      let ssSum = 0;
      const rows = polys.map((p, i) => { const r = contrastStat(p, ow); ssSum += r.ss; return [names[i] || ("Order " + (i + 1)), fmt(r.ss, 4), 1, fmt(r.F, 4), fmtP(r.p)]; });
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {strip}{levelHdr}
          <div style={{ fontWeight: "bold" }}>One-way ANOVA</div>{omnibus}
          <div style={{ fontWeight: "bold" }}>Orthogonal polynomial trend <span style={{ fontWeight: "normal", fontSize: 10, color: "#555" }}>(scores: {scores.join(", ")})</span></div>
          <StatTable head={["Component", "SS", "df", "F", "P"]} rows={rows} />
          <div style={{ fontSize: 10, color: "#555" }}>Single-df orthogonal-polynomial contrasts vs the pooled within error (df = {ow.dfw}). {Math.abs(ssSum - ow.ssb) < 1e-6 ? `Components partition the between SS exactly (Σ = ${fmt(ssSum, 3)} = ${fmt(ow.ssb, 3)}).` : `Σ components = ${fmt(ssSum, 3)} vs between SS ${fmt(ow.ssb, 3)}; an exact partition requires equal n.`} Scores assume equal spacing by default.</div>
        </div>
      );
    }
    let C = (Array.isArray(o.contrasts) ? o.contrasts : []).map((row) => Array.from({ length: k }, (_, j) => (row && row[j] != null) ? String(row[j]) : "0"));
    if (C.length === 0) C = [Array.from({ length: k }, (_, j) => j === 0 ? "1" : (j === 1 ? "-1" : "0"))];
    const setC = (next) => onOpt(a.id, { contrasts: next });
    const cell = { width: 46, fontSize: 11, textAlign: "center" };
    const resultRows = C.map((row, ri) => { const cv = row.map((v) => Number(v) || 0), sum = cv.reduce((s, v) => s + v, 0); if (Math.abs(sum) < 1e-9) { const r = contrastStat(cv, ow); return ["C" + (ri + 1), fmt(r.L, 4), fmt(r.se, 4), fmt(r.t, 4), r.df2, fmtP(r.p), fmt(r.ss, 4), fmt(r.F, 4), fmt(r.t / Math.sqrt(r.t * r.t + r.df2), 4)]; } return ["C" + (ri + 1), "Σc = " + fmt(sum, 3) + " ≠ 0 (not a valid contrast)", "", "", "", "", "", "", ""]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {strip}{levelHdr}
        <div style={{ fontWeight: "bold" }}>One-way ANOVA</div>{omnibus}
        <div style={{ fontWeight: "bold" }}>Contrast coefficients</div>
        <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
          <thead><tr><th style={{ padding: "2px 6px" }}></th>{levels.map((l) => <th key={l} style={{ padding: "2px 6px", borderBottom: "1px solid #999" }}>{l}</th>)}<th style={{ padding: "2px 6px" }}>Σ</th><th /></tr></thead>
          <tbody>{C.map((row, ri) => { const sum = row.reduce((s, v) => s + (Number(v) || 0), 0); return (
            <tr key={ri}><td style={{ padding: "2px 6px", fontWeight: "bold" }}>C{ri + 1}</td>
              {row.map((v, ci) => <td key={ci} style={{ padding: "1px 2px" }}><input type="text" value={v} style={cell} onClick={(e) => e.stopPropagation()} onChange={(e) => { const next = C.map((r) => r.slice()); next[ri][ci] = e.target.value; setC(next); }} /></td>)}
              <td style={{ padding: "2px 6px", color: Math.abs(sum) < 1e-9 ? "#070" : "#c00", fontWeight: "bold" }}>{fmt(sum, 2)}</td>
              <td style={{ padding: "2px 6px" }}>{C.length > 1 && <span style={{ cursor: "pointer", color: "#a00" }} onClick={(e) => { e.stopPropagation(); setC(C.filter((_, i) => i !== ri)); }}>✕</span>}</td>
            </tr>); })}</tbody>
        </table>
        <div><span style={{ cursor: "pointer", fontSize: 11, border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", borderRadius: 2, ...bevelOut }} onClick={(e) => { e.stopPropagation(); setC([...C, new Array(k).fill("0")]); }}>+ add contrast</span></div>
        <div style={{ fontWeight: "bold" }}>Contrast tests</div>
        <StatTable head={["Contrast", "Estimate", "SE", "t", "df", "P", "SS", "F", "r"]} rows={resultRows} />
        <div style={{ fontSize: 10, color: "#555" }}>L = Σ cᵢ·meanᵢ, tested against the pooled within error (df = {ow.dfw}); F = t². Coefficients should sum to zero. For several non-orthogonal or post-hoc comparisons, consider a multiplicity correction.</div>
      </div>
    );
  }
  if (a.type === "piechart" || a.type === "barchart") {
    const id = a.roles.var;
    if (!id) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign a categorical Variable.</div>;
    const name = colById[id]?.name || id, o = a.opts || {};
    const uniq = (arr) => { const s = [...new Set(arr)], allNum = s.length > 0 && s.every((v) => isFinite(Number(v))); return s.sort(allNum ? (x, y) => Number(x) - Number(y) : (x, y) => String(x).localeCompare(String(y))); };
    const gid = a.roles.group;
    if (a.type === "piechart" || !gid) {
      const vals = valuesOf(id).map((v) => (v == null ? "" : String(v))).filter((v) => v !== "");
      if (!vals.length) return <div style={{ color: "#a00" }}>No data to plot.</div>;
      const levels = uniq(vals), counts = Object.fromEntries(levels.map((l) => [l, 0])); vals.forEach((v) => counts[v]++);
      const total = vals.length;
      const stop = (e) => e.stopPropagation(), sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
      if (a.type === "piechart") {
        return (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <div style={{ fontSize: 10 }} onClick={stop}><label><input type="checkbox" checked={!!o.showPct} onChange={(e) => onOpt(a.id, { showPct: e.target.checked })} /> legend as %</label></div>
            <ExportFrame name={"pie_" + name}><PieChart slices={levels.map((l) => ({ label: l, value: counts[l] }))} title={name} showPct={!!o.showPct} /></ExportFrame>
            <StatTable head={[name, "Count", "%"]} rows={levels.map((l) => [l, counts[l], fmt(counts[l] / total * 100, 1) + "%"])} />
          </div>
        );
      }
      const measure = o.measure || "count";
      const data = levels.map((l) => [measure === "prop" ? counts[l] / total * 100 : counts[l]]);
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ fontSize: 10 }} onClick={stop}><label>Show: <select value={measure} onChange={(e) => onOpt(a.id, { measure: e.target.value })} style={sel}><option value="count">Counts</option><option value="prop">Percent</option></select></label></div>
          <ExportFrame name={"bar_" + name}><CatBarChart categories={levels} seriesNames={[""]} data={data} mode="grouped" title={name} xLabel={name} yLabel={measure === "prop" ? "Percent" : "Count"} /></ExportFrame>
          <StatTable head={[name, "Count", "%"]} rows={levels.map((l) => [l, counts[l], fmt(counts[l] / total * 100, 1) + "%"])} />
        </div>
      );
    }
    // grouped/stacked bar from two categorical variables
    const rRaw = valuesOf(id), cRaw = valuesOf(gid), rV = [], cV = [];
    for (let i = 0; i < rRaw.length; i++) { const rr = rRaw[i], cc = cRaw[i]; if (rr !== "" && rr != null && cc !== "" && cc != null) { rV.push(String(rr)); cV.push(String(cc)); } }
    if (!rV.length) return <div style={{ color: "#a00" }}>No complete cases.</div>;
    const ct = contingency(rV, cV), gName = colById[gid]?.name || gid;
    const measure = o.measure || "count", mode = o.mode || "grouped";
    const data = ct.O.map((row, ri) => row.map((v) => measure === "prop" ? (ct.rt[ri] ? v / ct.rt[ri] * 100 : 0) : v));
    const stop = (e) => e.stopPropagation(), sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10 }} onClick={stop}>
          <label>Bars: <select value={mode} onChange={(e) => onOpt(a.id, { mode: e.target.value })} style={sel}><option value="grouped">Grouped</option><option value="stacked">Stacked</option></select></label>
          <label>Show: <select value={measure} onChange={(e) => onOpt(a.id, { measure: e.target.value })} style={sel}><option value="count">Counts</option><option value="prop">% within {name}</option></select></label>
        </div>
        <ExportFrame name={"bar_" + name + "_" + gName}><CatBarChart categories={ct.rl} seriesNames={ct.cl} data={data} mode={mode} title={name + " by " + gName} xLabel={name} yLabel={measure === "prop" ? "Percent within " + name : "Count"} /></ExportFrame>
        <div style={{ fontSize: 10, color: "#555" }}>Series = {gName}. {measure === "prop" ? "Each " + name + " category's bars sum to 100%." : "Cell counts."}</div>
      </div>
    );
  }
  if (a.type === "forest") {
    const did = a.roles.dep, fid = a.roles.factor;
    if (!did || !fid) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign a Dependent variable and a Factor.</div>;
    const gr = groupArrays(valuesOf(did), valuesOf(fid));
    const stats = gr.keys.map((k) => { const d = describe(gr.map[k]); return { label: String(k), n: d.n, mean: d.mean, sd: d.sd }; }).filter((s) => s.n >= 2);
    if (stats.length < 2) return <div style={{ color: "#a00" }}>Need at least two groups with n ≥ 2.</div>;
    const o = a.opts || {}, conf = o.conf || 0.95, esKind = o.es === "g" ? "g" : "d";
    const refI = (() => { const i = stats.findIndex((s) => s.label === o.ref); return i >= 0 ? i : 0; })();
    const ref = stats[refI], others = stats.filter((_, i) => i !== refI);
    const fRows = others.map((s) => { const es = cohenTwoGroup(s.mean, s.sd, s.n, ref.mean, ref.sd, ref.n, conf); const est = esKind === "g" ? es.g : es.d, ci = esKind === "g" ? es.gCI : es.dCI; return { label: s.label, est, lo: ci[0], hi: ci[1] }; });
    const depName = colById[did]?.name || did, esLbl = esKind === "g" ? "Hedges g" : "Cohen's d", pct = Math.round(conf * 100);
    const stop = (e) => e.stopPropagation(), sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10, flexWrap: "wrap" }} onClick={stop}>
          <label>Reference: <select value={ref.label} onChange={(e) => onOpt(a.id, { ref: e.target.value })} style={sel}>{stats.map((s) => <option key={s.label} value={s.label}>{s.label}</option>)}</select></label>
          <label>Effect: <select value={esKind} onChange={(e) => onOpt(a.id, { es: e.target.value })} style={sel}><option value="d">Cohen's d</option><option value="g">Hedges g</option></select></label>
          <label>CI: <select value={String(conf)} onChange={(e) => onOpt(a.id, { conf: parseFloat(e.target.value) })} style={sel}><option value="0.9">90%</option><option value="0.95">95%</option><option value="0.99">99%</option></select></label>
        </div>
        <ExportFrame name={"forest_" + depName}><ForestPlot rows={fRows} title={`${depName}: ${esLbl} vs ${ref.label}`} xLabel={`${esLbl} (vs ${ref.label}), ${pct}% CI`} nullValue={0} /></ExportFrame>
        <StatTable head={["Group", "n", "Mean", "SD", esLbl, `${pct}% CI`]} rows={[[ref.label + " (reference)", ref.n, fmt(ref.mean, 2), fmt(ref.sd, 2), "—", "—"], ...others.map((s, i) => [s.label, s.n, fmt(s.mean, 2), fmt(s.sd, 2), fmt(fRows[i].est, 3), `[${fmt(fRows[i].lo, 2)}, ${fmt(fRows[i].hi, 2)}]`])]} />
        <div style={{ fontSize: 10, color: "#555" }}>Standardized mean difference of each group vs the reference, with noncentral-t {pct}% confidence intervals. Intervals crossing 0 (dashed line) are consistent with no difference.</div>
      </div>
    );
  }
  if (a.type === "crosstab") {
    const rid = a.roles.row, cid = a.roles.col;
    if (!rid || !cid) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign a Row and a Column variable.</div>;
    const rRaw = valuesOf(rid), cRaw = valuesOf(cid), rV = [], cV = [];
    for (let i = 0; i < rRaw.length; i++) { const rr = rRaw[i], cc = cRaw[i]; if (rr !== "" && rr != null && cc !== "" && cc != null) { rV.push(String(rr)); cV.push(String(cc)); } }
    if (rV.length < 1) return <div style={{ color: "#a00" }}>No complete cases.</div>;
    const ct = contingency(rV, cV), rowName = colById[rid]?.name || rid, colName = colById[cid]?.name || cid;
    const o = a.opts || {}, show = o.show || "count";
    const cellVal = (i, j) => { const O = ct.O[i][j]; if (show === "expected") return fmt(ct.E[i][j], 2); if (show === "rowpct") return ct.rt[i] ? fmt(O / ct.rt[i] * 100, 1) + "%" : "—"; if (show === "colpct") return ct.ct[j] ? fmt(O / ct.ct[j] * 100, 1) + "%" : "—"; return String(O); };
    const head = [rowName + " \\ " + colName, ...ct.cl, "Total"];
    const tblRows = ct.rl.map((rl, i) => [rl, ...ct.cl.map((_, j) => cellVal(i, j)), String(ct.rt[i])]);
    tblRows.push(["Total", ...ct.cl.map((_, j) => String(ct.ct[j])), String(ct.N)]);
    const statRows = [["Pearson χ²", fmt(ct.chi2, 4), ct.df, fmtP(ct.p)], ["Likelihood-ratio G²", fmt(ct.g2, 4), ct.df, fmtP(ct.pG)]];
    if (ct.R === 2 && ct.C === 2) statRows.push(["Yates' continuity χ²", fmt(ct.chi2y, 4), ct.df, fmtP(ct.pY)]);
    const effRows = [["Cramér's V", fmt(ct.V, 4)]];
    if (ct.R === 2 && ct.C === 2) { effRows.push(["φ (phi)", fmt(ct.phi, 4)], ["Fisher's exact (2-tailed) P", fmtP(ct.fisher)]); }
    const sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10, alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
          <label>Show: <select value={show} onChange={(e) => onOpt(a.id, { show: e.target.value })} style={sel}><option value="count">Counts</option><option value="expected">Expected</option><option value="rowpct">Row %</option><option value="colpct">Column %</option></select></label>
        </div>
        <StatTable head={head} rows={tblRows} />
        <div style={{ fontWeight: "bold" }}>Tests of independence</div>
        <StatTable head={["Test", "Value", "df", "P"]} rows={statRows} />
        <div style={{ fontWeight: "bold" }}>Effect size</div>
        <StatTable head={["Measure", "Value"]} rows={effRows} />
        {ct.lowE > 0 && <div style={{ fontSize: 10, color: "#a40" }}>⚠ {ct.lowE} of {ct.cells} cells ({fmt(ct.lowE / ct.cells * 100, 0)}%) have expected count &lt; 5 (min {fmt(ct.minE, 2)}); the χ² approximation may be unreliable.{ct.R === 2 && ct.C === 2 ? " Prefer Fisher's exact test." : ""}</div>}
        <div style={{ fontSize: 10, color: "#555" }}>χ² test of independence (H₀: the two variables are independent). Cramér's V ∈ [0, 1] gauges association strength.{ct.R === 2 && ct.C === 2 ? " Yates' correction and Fisher's exact apply to 2×2 tables." : ""}</div>
      </div>
    );
  }
  if (a.type === "corrviz") {
    const ids = a.roles.vars || [];
    if (ids.length < 2) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign at least two continuous variables.</div>;
    const names = ids.map((id) => colById[id]?.name || id), cols = ids.map((id) => valuesOf(id).map(num));
    const o = a.opts || {}, mode = ["heat", "scatter", "both"].includes(o.mode) ? o.mode : "both", method = o.method === "spearman" ? "spearman" : "pearson";
    const sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10, alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
          <label>Display: <select value={mode} onChange={(e) => onOpt(a.id, { mode: e.target.value })} style={sel}><option value="both">Combined</option><option value="heat">Heatmap</option><option value="scatter">Scatter matrix</option></select></label>
          <label>Method: <select value={method} onChange={(e) => onOpt(a.id, { method: e.target.value })} style={sel}><option value="pearson">Pearson</option><option value="spearman">Spearman</option></select></label>
        </div>
        <ExportFrame name="corrviz"><PairsPlot names={names} cols={cols} mode={mode} method={method} /></ExportFrame>
        <div style={{ fontSize: 10, color: "#555" }}>{mode === "scatter" ? "Pairwise scatterplots" : mode === "heat" ? (method === "spearman" ? "Spearman" : "Pearson") + " correlation heatmap" : "Lower triangle: pairwise scatterplots · upper triangle: " + (method === "spearman" ? "Spearman" : "Pearson") + " r"}. Color: red = positive, blue = negative, white ≈ 0. Pairwise deletion.</div>
      </div>
    );
  }
  if (a.type === "profile") {
    const m = anovaModel(a, colById, compactById, rows);
    if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
    if (m.withinNames.length === 0) return <div style={{ color: "#a00" }}>Assign a repeated-measures (compact “R”) variable with at least one within factor.</div>;
    const o = a.opts || {};
    const xFac = m.withinNames.includes(o.xFac) ? o.xFac : m.withinNames[0], levels = m.levelOrder[xFac];
    const subjects = buildProfile(m.long, xFac, levels, m.betweenNames);
    const groupOn = m.betweenNames.length > 0 && o.groupBy !== false, groupsOf = groupOn ? [...new Set(subjects.map((s) => s.group))] : [null];
    const errType = ["none", "se", "ci"].includes(o.err) ? o.err : "se";
    const means = groupsOf.map((g) => ({ group: g, ys: levels.map((_, li) => { const vals = subjects.filter((s) => g == null || s.group === g).map((s) => s.ys[li]).filter((v) => v != null), nn = vals.length; if (!nn) return { mean: null, err: 0, n: 0 }; const mean = vals.reduce((x, y) => x + y, 0) / nn, sd = nn > 1 ? Math.sqrt(vals.reduce((x, y) => x + (y - mean) ** 2, 0) / (nn - 1)) : 0, se = sd / Math.sqrt(nn); return { mean, err: errType === "none" ? 0 : errType === "ci" ? tCrit(0.05, Math.max(1, nn - 1)) * se : se, n: nn }; }) }));
    const sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" }, stop = (e) => e.stopPropagation();
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10, alignItems: "center", flexWrap: "wrap" }} onClick={stop}>
          {m.withinNames.length > 1 && <label>X-axis: <select value={xFac} onChange={(e) => onOpt(a.id, { xFac: e.target.value })} style={sel}>{m.withinNames.map((f) => <option key={f} value={f}>{f}</option>)}</select></label>}
          {m.betweenNames.length > 0 && <label><input type="checkbox" checked={o.groupBy !== false} onChange={(e) => onOpt(a.id, { groupBy: e.target.checked })} /> color by {m.betweenNames.join(" · ")}</label>}
          <label><input type="checkbox" checked={o.showSubjects !== false} onChange={(e) => onOpt(a.id, { showSubjects: e.target.checked })} /> show subjects</label>
          <label>Error bars: <select value={errType} onChange={(e) => onOpt(a.id, { err: e.target.value })} style={sel}><option value="none">none</option><option value="se">± SE</option><option value="ci">± 95% CI</option></select></label>
        </div>
        <ExportFrame name={"profile_" + m.depName}><ProfilePlot levels={levels} subjects={subjects} means={means} xLabel={xFac} yLabel={m.depName} title={m.depName + " across " + xFac} showSubjects={o.showSubjects !== false} /></ExportFrame>
        <div style={{ fontSize: 10, color: "#555" }}>Thin lines: individual subjects{groupOn ? " (colored by group)" : ""}. Thick line{groupsOf.length > 1 ? "s" : ""}: {groupOn ? "group means" : "grand mean"}{errType !== "none" ? ` (± ${errType === "ci" ? "95% CI" : "SE"})` : ""}. Any other within factors are averaged per subject.</div>
      </div>
    );
  }
  if (a.type === "corrmatrix") {
    const ids = a.roles.vars || [];
    if (ids.length < 2) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign at least two continuous variables.</div>;
    const names = ids.map((id) => colById[id]?.name || id), raw = ids.map((id) => valuesOf(id).map(num)), N = raw[0].length;
    const o = a.opts || {}, method = o.method === "spearman" ? "spearman" : "pearson", deletion = o.deletion === "listwise" ? "listwise" : "pairwise";
    let mask = null, listN = 0;
    if (deletion === "listwise") { mask = []; for (let i = 0; i < N; i++) { const ok = raw.every((c) => c[i] != null); mask.push(ok); if (ok) listN++; } }
    const corr = (ai, bi) => { const xs = [], ys = []; for (let i = 0; i < N; i++) { if (deletion === "listwise" && !mask[i]) continue; const a1 = raw[ai][i], b1 = raw[bi][i]; if (a1 != null && b1 != null) { xs.push(a1); ys.push(b1); } } if (xs.length < 3) return { r: NaN, p: NaN, n: xs.length }; if (method === "spearman") { const s = spearman(xs, ys, "two"); return { r: s.rho, p: s.p, n: s.n }; } return pearsonPair(xs, ys); };
    const stars = (p) => p < 0.001 ? "***" : p < 0.01 ? "**" : p < 0.05 ? "*" : "";
    let nMin = Infinity, nMax = 0;
    const rows = names.map((nm, i) => [nm, ...names.map((_, j) => { if (i === j) return "1"; const c = corr(i, j); if (!isFinite(c.r)) return "—"; nMin = Math.min(nMin, c.n); nMax = Math.max(nMax, c.n); return fmt(c.r, 3) + stars(c.p); })]);
    const sel = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", gap: 12, fontSize: 10, alignItems: "center", flexWrap: "wrap" }} onClick={(e) => e.stopPropagation()}>
          <label>Method: <select value={method} onChange={(e) => onOpt(a.id, { method: e.target.value })} style={sel}><option value="pearson">Pearson</option><option value="spearman">Spearman</option></select></label>
          <label>Missing: <select value={deletion} onChange={(e) => onOpt(a.id, { deletion: e.target.value })} style={sel}><option value="pairwise">Pairwise</option><option value="listwise">Listwise</option></select></label>
        </div>
        <StatTable head={["", ...names]} rows={rows} />
        <div style={{ fontSize: 10, color: "#555" }}>{method === "spearman" ? "Spearman rank" : "Pearson"} correlations. * P &lt; .05, ** P &lt; .01, *** P &lt; .001 (two-tailed). {deletion === "listwise" ? `Listwise deletion, N = ${listN}.` : (nMin === nMax ? `Pairwise deletion, N = ${isFinite(nMin) ? nMin : 0}.` : `Pairwise deletion, N = ${nMin}–${nMax}.`)}</div>
      </div>
    );
  }
  if (a.type === "reliability") {
    const ids = a.roles.items || [];
    if (ids.length < 2) return <div style={{ color: "#999", fontStyle: "italic" }}>Assign at least two scale items.</div>;
    const names = ids.map((id) => colById[id]?.name || id), raw = ids.map((id) => valuesOf(id).map(num)), N = raw[0].length, k = ids.length;
    const items = ids.map(() => []); let n = 0;
    for (let i = 0; i < N; i++) { if (raw.every((c) => c[i] != null)) { ids.forEach((_, j) => items[j].push(raw[j][i])); n++; } }
    if (n < 3) return <div style={{ color: "#a00" }}>Need at least 3 complete cases (have {n}).</div>;
    const alpha = cronbachAlpha(items);
    const inter = []; for (let i = 0; i < k; i++) for (let j = i + 1; j < k; j++) { const r = pearsonPair(items[i], items[j]).r; if (isFinite(r)) inter.push(r); }
    const rbar = inter.reduce((s, v) => s + v, 0) / inter.length, stdAlpha = k * rbar / (1 + (k - 1) * rbar);
    const itemRows = items.map((col, j) => { const rest = []; for (let i = 0; i < n; i++) { let s = 0; for (let m = 0; m < k; m++) if (m !== j) s += items[m][i]; rest.push(s); } const r = pearsonPair(col, rest).r, ad = cronbachAlpha(items.filter((_, m) => m !== j)), mean = col.reduce((s, v) => s + v, 0) / n; return [names[j], fmt(mean, 3), fmt(Math.sqrt(_variance(col)), 4), fmt(r, 4), fmt(ad, 4)]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <StatTable head={["Reliability", "Value"]} rows={[["Cronbach's α", fmt(alpha, 4)], ["Standardized α", fmt(stdAlpha, 4)], ["Items (k)", k], ["Cases (n)", n], ["Mean inter-item r", fmt(rbar, 4)]]} />
        <div style={{ fontWeight: "bold" }}>Item-total statistics</div>
        <StatTable head={["Item", "Mean", "SD", "Corrected item–total r", "α if item deleted"]} rows={itemRows} />
        <div style={{ fontSize: 10, color: "#555" }}>α uses listwise-complete cases. Corrected item–total r correlates each item with the sum of the others. “α if item deleted” shows the scale's α with that item removed — values above the overall α flag items that may be weakening the scale. Standardized α is based on the mean inter-item correlation.</div>
      </div>
    );
  }
  if (a.type === "normtest") {
    const yRaw = valuesOf(a.roles.y), gid = a.roles.group;
    const fmtW = (r) => r.error ? r.error : `${fmt(r.W, 4)}`;
    if (!gid) {
      const yv = numArr(yRaw), r = shapiroWilk(yv);
      if (r.error) return <div style={{ color: "#a00" }}>{r.error}</div>;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <StatTable head={["Shapiro–Wilk", ""]} rows={[["n", r.n], ["W", fmt(r.W, 4)], ["P", fmtP(r.p)]]} />
          <ExportFrame name="qq-plot"><QQPlot data={yv} title="Normal Q–Q Plot" /></ExportFrame>
          <div style={{ fontSize: 10, color: "#555" }}>H₀: the data are normally distributed. A small P (e.g. &lt; 0.05) is evidence against normality. Points near the line ⇒ approximately normal.</div>
        </div>
      );
    }
    const gRaw = valuesOf(gid), groups = new Map();
    for (let i = 0; i < yRaw.length; i++) { const yy = num(yRaw[i]), g = gRaw[i]; if (yy != null && g !== "" && g != null) { if (!groups.has(g)) groups.set(g, []); groups.get(g).push(yy); } }
    const entries = [...groups.entries()];
    const rows = entries.map(([g, v]) => { const r = shapiroWilk(v); return [g, v.length, r.error ? "—" : fmt(r.W, 4), r.error ? r.error : fmtP(r.p)]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <StatTable head={["Group", "n", "W", "P"]} rows={rows} />
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
          {entries.filter(([, v]) => v.length >= 3).map(([g, v]) => <ExportFrame key={g} name={"qq-" + g}><QQPlot data={v} title={"Q–Q: " + g} /></ExportFrame>)}
        </div>
        <div style={{ fontSize: 10, color: "#555" }}>Shapiro–Wilk per group. Small P ⇒ evidence against normality within that group.</div>
      </div>
    );
  }
  if (a.type === "homovar") {
    const yRaw = valuesOf(a.roles.y), gRaw = valuesOf(a.roles.group), groups = new Map();
    for (let i = 0; i < yRaw.length; i++) { const yy = num(yRaw[i]), g = gRaw[i]; if (yy != null && g !== "" && g != null) { if (!groups.has(g)) groups.set(g, []); groups.get(g).push(yy); } }
    const entries = [...groups.entries()].filter(([, v]) => v.length >= 2);
    if (entries.length < 2) return <div style={{ color: "#a00" }}>Need at least two groups with ≥ 2 observations each.</div>;
    const arrs = entries.map(([, v]) => v);
    const lev = leveneTest(arrs, "median"), levMean = leveneTest(arrs, "mean"), bart = bartlettTest(arrs);
    const sdRows = entries.map(([g, v]) => { const m = v.reduce((s, x) => s + x, 0) / v.length, sd = Math.sqrt(v.reduce((s, x) => s + (x - m) ** 2, 0) / (v.length - 1)); return [g, v.length, fmt(m, 3), fmt(sd, 4), fmt(sd * sd, 4)]; });
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <StatTable head={["Group", "n", "Mean", "SD", "Variance"]} rows={sdRows} />
        <div style={{ fontWeight: "bold" }}>Tests of equal variance</div>
        <StatTable head={["Test", "Statistic", "df", "P"]} rows={[
          ["Levene (median)", fmt(lev.W, 4), `${lev.df1}, ${lev.df2}`, fmtP(lev.p)],
          ["Levene (mean)", fmt(levMean.W, 4), `${levMean.df1}, ${levMean.df2}`, fmtP(levMean.p)],
          ["Bartlett", fmt(bart.chi2, 4), `${bart.df}`, fmtP(bart.p)],
        ]} />
        <div style={{ fontSize: 10, color: "#555" }}>H₀: all groups share a common variance. Small P ⇒ heterogeneity. Levene (median) [Brown–Forsythe] is robust to non-normality and is the usual default; Bartlett is more powerful but assumes normality.</div>
      </div>
    );
  }
  if (a.type === "glm") {
    const o = { factorial: true, ci: 0.95, ...(a.opts || {}) };
    const Y = colById[a.roles.y]?.name || a.roles.y;
    const fids = a.roles.factors || [], cids = a.roles.covs || [];
    const fnames = fids.map((id) => colById[id]?.name || id), cnames = cids.map((id) => colById[id]?.name || id);
    const pct = Math.round(o.ci * 100);
    const strip = (
      <div style={{ display: "flex", gap: 14, fontSize: 10, color: "#444", alignItems: "center" }}>
        {fids.length >= 2 && <label onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={o.factorial !== false} onChange={(e) => onOpt(a.id, { factorial: e.target.checked })} /> factor interactions</label>}
        {fids.length >= 1 && cids.length >= 1 && <label onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={!!o.slopes} onChange={(e) => onOpt(a.id, { slopes: e.target.checked })} /> factor × covariate slopes</label>}
        <label onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={!!o.diag} onChange={(e) => onOpt(a.id, { diag: e.target.checked })} /> residual diagnostics</label>
        <label onClick={(e) => e.stopPropagation()}>CI level:
          <select value={String(o.ci)} onChange={(e) => onOpt(a.id, { ci: parseFloat(e.target.value) })} style={{ fontFamily: FONT, fontSize: 10, marginLeft: 4, border: `1px solid ${PLAT.dark}`, background: "#fff" }}>
            <option value="0.9">90%</option><option value="0.95">95%</option><option value="0.99">99%</option>
          </select>
        </label>
      </div>
    );
    if (fids.length === 0 && cids.length === 0) return (<div>{strip}<div style={{ color: "#a00", marginTop: 8 }}>Assign at least one factor and/or covariate.</div></div>);
    const r = glmAnalyze(valuesOf(a.roles.y), fids.map((id) => valuesOf(id)), cids.map((id) => valuesOf(id)), fnames, cnames, o);
    if (r.error) return (<div>{strip}<div style={{ color: "#a00", marginTop: 8 }}>{r.error}</div></div>);
    const termLabel = fnames.length ? (o.factorial !== false && fnames.length >= 2 ? fnames.join(" × ") : fnames.join(" + ")) : "";
    const modelDesc = `${Y} ~ ${[termLabel, ...cnames].filter(Boolean).join(" + ")}`;
    const effRows = [
      ["Corrected Model", r.dfModel, fmt(r.ssModel, 2), fmt(r.ssModel / r.dfModel, 2), fmt(r.Fmodel, 3), fmtP(r.pModel)],
      ...r.effects.map((e) => [e.name, e.df, fmt(e.ss, 2), fmt(e.ms, 2), fmt(e.F, 3), fmtP(e.p)]),
      ["Residual", r.residual.df, fmt(r.residual.ss, 2), fmt(r.residual.ms, 2), "", ""],
      ["Corrected Total", r.total.df, fmt(r.total.ss, 2), "", "", ""],
    ];
    const parRows = r.params.map((c) => [c.name, fmt(c.b, 4), fmt(c.se, 4), `[${fmt(c.lo, 3)}, ${fmt(c.hi, 3)}]`, fmt(c.t, 3), fmtP(c.p)]);
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {strip}
        <div style={{ fontSize: 11 }}>{modelDesc} &nbsp; (n = {r.n})</div>
        <StatTable head={["", "Value"]} rows={[["R Squared", fmt(r.r2, 4)], ["Adjusted R²", fmt(r.adjR2, 4)], ["RMS Residual", fmt(Math.sqrt(r.residual.ms), 4)]]} />
        <div style={{ fontWeight: "bold" }}>Type III Tests of Effects</div>
        <StatTable head={["Source", "DF", "Sum Sq", "Mean Sq", "F", "P"]} rows={effRows} />
        <div style={{ fontWeight: "bold" }}>Parameter Estimates (treatment coding)</div>
        <StatTable head={["Term", "Coeff.", "Std. Err", `${pct}% CI`, "t", "P"]} rows={parRows} />
        <div style={{ fontSize: 10, color: "#555" }}>Type III SS via sum-to-zero coding (marginal hypotheses); parameter estimates use treatment coding with each factor's first level as reference. Covariates enter as main effects (homogeneous-slopes ANCOVA).</div>
        {o.diag && <ResidualDiagnostics resid={r.resid} fitted={r.fitted} />}
      </div>
    );
  }
  if (a.type === "ttest") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    const twoCol = o.srcMode === "twoCol";
    const opts = <CmpOptions a={a} o={o} onOpt={onOpt} kind="unpaired" />;
    let res;
    if (twoCol) {
      if (!a.roles.x1 || !a.roles.x2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign two continuous variables (treated as independent samples).</div></div>);
      const l1 = colById[a.roles.x1]?.name || a.roles.x1, l2 = colById[a.roles.x2]?.name || a.roles.x2;
      res = unpairedCompTwo(valuesOf(a.roles.x1), valuesOf(a.roles.x2), l1, l2, o.mu0, o.tail, o.varAssume === "welch");
    } else {
      if (!a.roles.y || !a.roles.group) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign a Test Variable and a 2-level Grouping variable.</div></div>);
      res = unpairedComp(valuesOf(a.roles.y), valuesOf(a.roles.group), o.mu0, o.tail, o.varAssume === "welch");
    }
    if (res.error !== undefined) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>{res.error === "size" ? "Each group needs at least 2 values." : `Grouping variable must have exactly 2 levels (found ${Array.isArray(res.keys) ? res.keys.length : "?"}).`}</div></div>);
    const pL = tailLabel(o.tail);
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <StatTable head={[twoCol ? "Variable" : "Group", "Count", "Mean", "Std. Dev.", "Std. Error"]} rows={res.groups.map((g) => [g.key, g.n, fmt(g.mean), fmt(g.sd), fmt(g.sem)])} />
        <div style={{ fontWeight: "bold" }}>{res.welch ? "Unpaired t-test — Welch (unequal variances)" : "Unpaired t-test — pooled variance"}</div>
        <StatTable head={["Hyp. Diff.", "Mean Diff.", "DF", "t Value", pL]} rows={[[fmt(o.mu0), fmt(res.meanDiff), fmt(res.df, 2), fmt(res.t), fmtP(res.p)]]} />
        {o.showF && (<>
          <div style={{ fontWeight: "bold" }}>F-test for equality of variances</div>
          <StatTable head={["F (s₁²/s₂²)", "DF", tailLabel(o.tail)]} rows={[[fmt(res.F.F), `${res.F.df1}, ${res.F.df2}`, fmtP(res.F.p)]]} />
        </>)}
        {o.effSize && (() => {
          const g0 = res.groups[0], g1 = res.groups[1], conf = o.effConf || 0.95, pc = Math.round(conf * 100);
          const es = cohenTwoGroup(g0.mean, g0.sd, g0.n, g1.mean, g1.sd, g1.n, conf);
          return (<div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <div style={{ fontWeight: "bold" }}>Effect size ({g0.key} − {g1.key})</div>
            <StatTable head={["Measure", "Value", `${pc}% CI`]} rows={[
              ["Cohen's d", fmt(es.d, 4), `[${fmt(es.dCI[0], 4)}, ${fmt(es.dCI[1], 4)}]`],
              ["Hedges' g", fmt(es.g, 4), `[${fmt(es.gCI[0], 4)}, ${fmt(es.gCI[1], 4)}]`],
              [`Glass's Δ (${g0.key} SD)`, fmt(es.glass1, 4), ""],
              [`Glass's Δ (${g1.key} SD)`, fmt(es.glass2, 4), ""],
            ]} />
            <div style={{ fontSize: 10, color: "#555" }}>d uses the pooled SD; Hedges' g applies the small-sample bias correction. Glass's Δ standardizes by one group's SD (useful when variances differ). CI for d via the noncentral t.</div>
          </div>);
        })()}
        <div style={{ fontSize: 10, color: "#555" }}>{res.groups[0].key} − {res.groups[1].key}. {o.tail === "two" ? "Two-tailed." : o.tail === "upper" ? "One-tailed (group 1 > group 2 + hyp. diff)." : "One-tailed (group 1 < group 2 + hyp. diff)."}</div>
      </div>
    );
  }
  if (a.type === "paired") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    if (!a.roles.x1 || !a.roles.x2) return (<div><CmpOptions a={a} o={o} onOpt={onOpt} kind="paired" /><div style={{ color: "#a00", marginTop: 8 }}>Assign two paired variables.</div></div>);
    const res = pairedT(valuesOf(a.roles.x1), valuesOf(a.roles.x2), o.mu0, o.tail);
    if (res.n < 2) return (<div><CmpOptions a={a} o={o} onOpt={onOpt} kind="paired" /><div style={{ color: "#a00", marginTop: 8 }}>Need at least 2 complete pairs (have {res.n}).</div></div>);
    const n1 = colById[a.roles.x1]?.name || a.roles.x1, n2 = colById[a.roles.x2]?.name || a.roles.x2;
    const pL = tailLabel(o.tail);
    const corr = o.showCorr ? corrZTest(res.r, res.n, o.rho0, o.tail) : null;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <CmpOptions a={a} o={o} onOpt={onOpt} kind="paired" />
        <StatTable head={["Variable", "Count", "Mean", "Std. Dev.", "Std. Error"]} rows={[[n1, res.d1.n, fmt(res.d1.mean), fmt(res.d1.sd), fmt(res.d1.sem)], [n2, res.d2.n, fmt(res.d2.mean), fmt(res.d2.sd), fmt(res.d2.sem)]]} />
        <div style={{ fontWeight: "bold" }}>Paired t-test ({n1} − {n2})</div>
        <StatTable head={["Hyp. Diff.", "Mean Diff.", "DF", "t Value", pL]} rows={[[fmt(o.mu0), fmt(res.meanDiff), res.df, fmt(res.t), fmtP(res.p)]]} />
        {o.showCorr && (<>
          <div style={{ fontWeight: "bold" }}>Correlation — Fisher z-test</div>
          {corr.invalid
            ? <div style={{ fontSize: 10, color: "#a00" }}>Correlation test needs ≥ 4 pairs and |r| &lt; 1 (r = {fmt(res.r)}, n = {res.n}).</div>
            : <StatTable head={["Count", "Corr. (r)", "Hyp. ρ", "z Value", tailLabel(o.tail)]} rows={[[corr.n, fmt(corr.r), fmt(corr.rho0), fmt(corr.Z), fmtP(corr.p)]]} />}
        </>)}
        {o.effSize && res.se > 0 && (() => {
          const conf = o.effConf || 0.95, pc = Math.round(conf * 100), t0 = res.meanDiff / res.se;
          const es = cohenPaired(t0, res.n, conf);
          return (<div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <div style={{ fontWeight: "bold" }}>Effect size ({n1} − {n2})</div>
            <StatTable head={["Measure", "Value", `${pc}% CI`]} rows={[
              ["Cohen's dz", fmt(es.dz, 4), `[${fmt(es.dCI[0], 4)}, ${fmt(es.dCI[1], 4)}]`],
              ["Hedges' gz", fmt(es.gz, 4), `[${fmt(es.gCI[0], 4)}, ${fmt(es.gCI[1], 4)}]`],
            ]} />
            <div style={{ fontSize: 10, color: "#555" }}>dz standardizes the mean difference by the SD of the differences (the within-subject effect size); CI via the noncentral t.</div>
          </div>);
        })()}
        <div style={{ fontSize: 10, color: "#555" }}>{o.tail === "two" ? "Two-tailed." : o.tail === "upper" ? "One-tailed (upper)." : "One-tailed (lower)."} Differences computed pairwise over complete cases.</div>
      </div>
    );
  }
  if (a.type === "mannwhitney") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    const twoCol = o.srcMode === "twoCol";
    const opts = <CmpOptions a={a} o={o} onOpt={onOpt} kind="mannwhitney" />;
    let a1, a2, l1, l2;
    if (twoCol) {
      if (!a.roles.x1 || !a.roles.x2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign two continuous variables.</div></div>);
      l1 = colById[a.roles.x1]?.name || a.roles.x1; l2 = colById[a.roles.x2]?.name || a.roles.x2;
      a1 = numArr(valuesOf(a.roles.x1)); a2 = numArr(valuesOf(a.roles.x2));
    } else {
      if (!a.roles.y || !a.roles.group) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign a Test Variable and a 2-level Grouping variable.</div></div>);
      const gr = groupArrays(valuesOf(a.roles.y), valuesOf(a.roles.group));
      if (gr.keys.length !== 2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Grouping variable must have exactly 2 levels (found {gr.keys.length}).</div></div>);
      l1 = gr.keys[0]; l2 = gr.keys[1]; a1 = gr.map[l1]; a2 = gr.map[l2];
    }
    if (a1.length < 1 || a2.length < 1) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Each sample needs at least one value.</div></div>);
    const res = mannWhitney(a1, a2, o.tail, o.cc);
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <StatTable head={[twoCol ? "Variable" : "Group", "Count", "Rank Sum"]} rows={[[l1, res.n1, fmt(res.R1, 1)], [l2, res.n2, fmt(res.R2, 1)]]} />
        <div style={{ fontWeight: "bold" }}>Mann–Whitney U</div>
        <StatTable head={["U (" + l1 + ")", "U (" + l2 + ")", "z", tailLabel(o.tail)]} rows={[[fmt(res.U1, 1), fmt(res.U2, 1), fmt(res.z), fmtP(res.p)]]} />
        <div style={{ fontSize: 10, color: "#555" }}>Normal approximation, tie-corrected{o.cc ? ", continuity-corrected" : ""}. {l1} vs {l2}.</div>
      </div>
    );
  }
  if (a.type === "wilcoxon") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    const opts = <CmpOptions a={a} o={o} onOpt={onOpt} kind="wilcoxon" />;
    if (!a.roles.x1 || !a.roles.x2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign two paired variables.</div></div>);
    const res = wilcoxonSR(pairedDiffs(valuesOf(a.roles.x1), valuesOf(a.roles.x2), o.mu0), o.tail, o.cc);
    if (res.n < 1) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>No non-zero differences to rank.</div></div>);
    const n1 = colById[a.roles.x1]?.name || a.roles.x1, n2 = colById[a.roles.x2]?.name || a.roles.x2;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <div style={{ fontWeight: "bold" }}>Wilcoxon Signed-Rank ({n1} − {n2})</div>
        <StatTable head={["n (non-zero)", "W+", "W−", "z", tailLabel(o.tail)]} rows={[[res.n, fmt(res.rPlus, 1), fmt(res.rMinus, 1), fmt(res.z), fmtP(res.p)]]} />
        <div style={{ fontSize: 10, color: "#555" }}>Zero differences dropped; tie-corrected normal approximation{o.cc ? " with continuity correction" : ""}.</div>
      </div>
    );
  }
  if (a.type === "signtest") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    const opts = <CmpOptions a={a} o={o} onOpt={onOpt} kind="signtest" />;
    if (!a.roles.x1 || !a.roles.x2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign two paired variables.</div></div>);
    const res = signTest(pairedDiffs(valuesOf(a.roles.x1), valuesOf(a.roles.x2), o.mu0), o.tail);
    if (res.n < 1) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>No non-tied pairs.</div></div>);
    const n1 = colById[a.roles.x1]?.name || a.roles.x1, n2 = colById[a.roles.x2]?.name || a.roles.x2;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <div style={{ fontWeight: "bold" }}>Sign Test ({n1} − {n2})</div>
        <StatTable head={["n (±)", "n (+)", "n (−)", tailLabel(o.tail)]} rows={[[res.n, res.nplus, res.nminus, fmtP(res.p)]]} />
        <div style={{ fontSize: 10, color: "#555" }}>Exact binomial; tied pairs dropped.</div>
      </div>
    );
  }
  if (a.type === "spearman") {
    const o = { ...DEFAULT_CMP_OPTS, ...(a.opts || {}) };
    const opts = <CmpOptions a={a} o={o} onOpt={onOpt} kind="spearman" />;
    if (!a.roles.x1 || !a.roles.x2) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign two variables.</div></div>);
    const p1 = [], p2 = []; const v1 = valuesOf(a.roles.x1), v2 = valuesOf(a.roles.x2);
    for (let i = 0; i < Math.min(v1.length, v2.length); i++) { const x = num(v1[i]), y = num(v2[i]); if (x !== null && y !== null) { p1.push(x); p2.push(y); } }
    if (p1.length < 3) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Need at least 3 complete pairs (have {p1.length}).</div></div>);
    const res = spearman(p1, p2, o.tail);
    const n1 = colById[a.roles.x1]?.name || a.roles.x1, n2 = colById[a.roles.x2]?.name || a.roles.x2;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <div style={{ fontWeight: "bold" }}>Spearman Rank Correlation ({n1}, {n2})</div>
        <StatTable head={["Count", "Spearman ρ", "t", "DF", tailLabel(o.tail)]} rows={[[res.n, fmt(res.rho), fmt(res.t), res.df, fmtP(res.p)]]} />
        <div style={{ fontSize: 10, color: "#555" }}>Tie-corrected ρ; t-distribution approximation for p.</div>
      </div>
    );
  }
  if (a.type === "kruskal") {
    const gr = groupArrays(valuesOf(a.roles.y), valuesOf(a.roles.group));
    if (gr.keys.length < 2) return <div style={{ color: "#a00" }}>Grouping variable must have at least 2 levels (found {gr.keys.length}).</div>;
    if (gr.keys.some((k) => gr.map[k].length < 1)) return <div style={{ color: "#a00" }}>Each group needs at least one value.</div>;
    const res = kruskal(gr.keys, gr.map);
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <StatTable head={["Group", "Count", "Mean Rank"]} rows={res.groups.map((g) => [g.key, g.n, fmt(g.meanRank, 2)])} />
        <div style={{ fontWeight: "bold" }}>Kruskal–Wallis</div>
        <StatTable head={["H", "DF", "P (χ²)"]} rows={[[fmt(res.H), res.df, fmtP(res.p)]]} />
        <div style={{ fontSize: 10, color: "#555" }}>Tie-corrected; chi-square approximation, {res.k} groups, N = {res.N}.</div>
      </div>
    );
  }
  if (a.type === "friedman") {
    const dep = a.roles.dep, cp = dep && dep.kind === "compact" ? compactById[dep.id] : null;
    if (!cp) return <div style={{ color: "#a00" }}>Assign a repeated-measures (compact) variable.</div>;
    if (cp.leaves.length < 2) return <div style={{ color: "#a00" }}>Friedman needs at least 2 conditions (the compact has {cp.leaves.length}).</div>;
    const blocks = [];
    rows.forEach((r) => { const vals = cp.leaves.map((id) => num(r[id])); if (vals.every((v) => v !== null)) blocks.push(vals); });
    if (blocks.length < 2) return <div style={{ color: "#a00" }}>Need at least 2 complete cases (have {blocks.length}).</div>;
    const res = friedman(blocks);
    const condLabel = (j) => { const sizes = cp.factors.map((f) => f.levels.length); return cp.factors.map((f, fi) => { const inner = sizes.slice(fi + 1).reduce((p, s) => p * s, 1); return f.levels[Math.floor(j / inner) % f.levels.length]; }).join(" · ") || cp.leaves[j]; };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ fontSize: 11 }}>{cp.name}: {cp.leaves.length} conditions, n = {res.n} complete cases.</div>
        <StatTable head={["Q (χ²)", "DF", "P"]} rows={[[fmt(res.Q), res.df, fmtP(res.p)]]} />
        <StatTable head={["Condition", "Rank Sum", "Mean Rank"]} rows={cp.leaves.map((id, j) => [condLabel(j), fmt(res.Rj[j], 1), fmt(res.Rj[j] / res.n, 2)])} />
        <div style={{ fontSize: 10, color: "#555" }}>Tie-corrected Friedman across all leaf conditions; blocks = cases.</div>
      </div>
    );
  }
  if (a.type === "boxplot" || a.type === "violin") {
    const mode = a.type === "boxplot" ? "box" : "violin";
    const o = { ...DEFAULT_DIST_OPTS, ...(a.opts || {}) };
    const opts = <DistOptions a={a} o={o} onOpt={onOpt} mode={mode} />;
    if (!a.roles.y) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>Assign a Variable (and optionally a Grouping variable) in the Variables window.</div></div>);
    const yName = colById[a.roles.y]?.name || a.roles.y;
    let groups;
    if (a.roles.group) { const gr = groupArrays(valuesOf(a.roles.y), valuesOf(a.roles.group)); groups = gr.keys.map((k) => ({ label: String(k), values: gr.map[k] })); }
    else groups = [{ label: yName, values: numArr(valuesOf(a.roles.y)) }];
    groups = groups.filter((g) => g.values.length >= 1);
    if (!groups.length) return (<div>{opts}<div style={{ color: "#a00", marginTop: 8 }}>No numeric data to plot.</div></div>);
    const gName = a.roles.group ? (colById[a.roles.group]?.name || a.roles.group) : null;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {opts}
        <ExportFrame name={mode + "_" + yName}>
          <BoxViolinChart groups={groups} mode={mode} opts={o} title={yName + (gName ? " by " + gName : "")} yLabel={yName} />
        </ExportFrame>
        <StatTable head={["Group", "n", "Median", "Q1", "Q3", "Mean"]} rows={groups.map((g) => { const b = boxStats(g.values); return [g.label, b.n, fmt(b.med), fmt(b.q1), fmt(b.q3), fmt(b.mean)]; })} />
      </div>
    );
  }
  return null;
}

function DistOptions({ a, o, onOpt, mode }) {
  const set = (patch) => onOpt(a.id, patch);
  const stop = (e) => e.stopPropagation();
  const lbl = { fontSize: 10, color: "#444", display: "inline-flex", alignItems: "center", gap: 3 };
  const inp = { fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff", padding: "0 2px" };
  return (
    <div onClick={stop} style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", padding: "4px 6px", background: "#f1ece4", border: `1px solid ${PLAT.dark}`, borderRadius: 3 }}>
      <label style={lbl}><input type="checkbox" checked={!!o.showPoints} onChange={(e) => set({ showPoints: e.target.checked })} /> show points</label>
      {mode === "box" && <label style={lbl}><input type="checkbox" checked={o.showMean !== false} onChange={(e) => set({ showMean: e.target.checked })} /> show mean</label>}
      {mode === "violin" && <label style={lbl}><input type="checkbox" checked={o.showBox !== false} onChange={(e) => set({ showBox: e.target.checked })} /> inner box</label>}
      {mode === "violin" && <label style={lbl}>Bandwidth:
        <select value={o.bw || 1} onChange={(e) => set({ bw: parseFloat(e.target.value) })} style={{ ...inp, width: "auto" }}>
          <option value="0.5">0.5× (sharper)</option>
          <option value="1">1× (Silverman)</option>
          <option value="1.5">1.5× (smoother)</option>
        </select>
      </label>}
    </div>
  );
}

function CmpOptions({ a, o, onOpt, kind }) {
  const cap = CMP_CAP[kind] || {};
  const set = (patch) => onOpt(a.id, patch);
  const stop = (e) => e.stopPropagation();
  const lbl = { fontSize: 10, color: "#444", display: "inline-flex", alignItems: "center", gap: 3 };
  const inp = { fontFamily: FONT, fontSize: 10, width: 58, border: `1px solid ${PLAT.dark}`, background: "#fff", padding: "0 2px" };
  return (
    <div onClick={stop} style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", padding: "4px 6px", background: "#f1ece4", border: `1px solid ${PLAT.dark}`, borderRadius: 3 }}>
      {cap.srcMode && (<label style={lbl}>Groups from:
        <select value={o.srcMode} onChange={(e) => set({ srcMode: e.target.value })} style={{ ...inp, width: "auto" }}>
          <option value="grouped">Grouping variable</option>
          <option value="twoCol">Two variables</option>
        </select>
      </label>)}
      {cap.tail && (<label style={lbl}>Tail:
        <select value={o.tail} onChange={(e) => set({ tail: e.target.value })} style={{ ...inp, width: "auto" }}>
          <option value="two">Two-tailed</option>
          <option value="upper">One-tailed (upper, &gt;)</option>
          <option value="lower">One-tailed (lower, &lt;)</option>
        </select>
      </label>)}
      {cap.mu0 && (<label style={lbl}>Hyp. difference:
        <input type="number" step="any" value={o.mu0} onChange={(e) => set({ mu0: e.target.value === "" ? 0 : parseFloat(e.target.value) })} style={inp} />
      </label>)}
      {cap.variance && (<><label style={lbl}>Variance:
        <select value={o.varAssume} onChange={(e) => set({ varAssume: e.target.value })} style={{ ...inp, width: "auto" }}>
          <option value="pooled">Pooled (equal)</option>
          <option value="welch">Unequal (Welch)</option>
        </select>
      </label>
      <label style={lbl}><input type="checkbox" checked={o.showF} onChange={(e) => set({ showF: e.target.checked })} /> F-test for variances</label></>)}
      {cap.corr && (<><label style={lbl}><input type="checkbox" checked={o.showCorr} onChange={(e) => set({ showCorr: e.target.checked })} /> correlation z-test</label>
        {o.showCorr && (<label style={lbl}>Hyp. ρ:
          <input type="number" step="any" value={o.rho0} onChange={(e) => set({ rho0: e.target.value === "" ? 0 : parseFloat(e.target.value) })} style={inp} />
        </label>)}</>)}
      {cap.cc && (<label style={lbl}><input type="checkbox" checked={o.cc} onChange={(e) => set({ cc: e.target.checked })} /> continuity correction</label>)}
      {(kind === "unpaired" || kind === "paired") && (<>
        <label style={lbl}><input type="checkbox" checked={!!o.effSize} onChange={(e) => set({ effSize: e.target.checked })} /> effect size</label>
        {o.effSize && (<label style={lbl}>CI:
          <select value={o.effConf || 0.95} onChange={(e) => set({ effConf: Number(e.target.value) })} style={{ ...inp, width: "auto" }}>
            <option value={0.95}>95%</option><option value={0.9}>90%</option>
          </select>
        </label>)}
      </>)}
    </div>
  );
}

/* ---- ANOVA renderer ---- */
/* Build the shared long-format model + factor metadata for any ANOVA output. */
function anovaModel(a, colById, compactById, rows) {
  const dep = a.roles.dep;
  const betweenIds = a.roles.between || [];
  if (!dep) return { error: "Assign a Dependent/Repeated variable. Use a compact (R) variable for repeated measures, or a plain continuous column for a between-only factorial." };
  const betweenSpecs = betweenIds.map((id) => ({ id, name: colById[id]?.name || id }));
  let long, withinNames = [], depName, levelOrder = {};
  if (dep.kind === "compact") {
    const cp = compactById[dep.id];
    if (!cp) return { error: "Compact variable no longer exists." };
    long = compactToLong(cp, betweenSpecs, rows);
    withinNames = cp.factors.map((f) => f.name);
    cp.factors.forEach((f) => (levelOrder[f.name] = [...f.levels]));
    depName = cp.name;
  } else {
    const subject = [], y = []; const bvals = {}; betweenSpecs.forEach((b) => (bvals[b.name] = []));
    rows.forEach((r, ri) => {
      const v = num(r[dep.id]); if (v === null) return;
      if (betweenSpecs.some((b) => r[b.id] === "" || r[b.id] == null)) return;
      subject.push("S" + ri); y.push(v); betweenSpecs.forEach((b) => bvals[b.name].push(r[b.id]));
    });
    long = { subject, between: bvals, within: {}, y };
    depName = colById[dep.id]?.name;
  }
  const betweenNames = betweenSpecs.map((b) => b.name);
  // level order for between factors = order of first appearance
  betweenNames.forEach((nm) => {
    const seen = []; long.between[nm].forEach((v) => { if (!seen.includes(v)) seen.push(v); }); levelOrder[nm] = seen;
  });
  if (long.y.length < 2) return { error: "Not enough complete observations." };
  if (betweenNames.length === 0 && withinNames.length === 0) return { error: "Add at least one factor (a between factor and/or a compact variable with within factors)." };
  // factor order for effects: between first, then within (so within tends to land on the x-axis)
  const factors = [...betweenNames, ...withinNames];
  return { long, withinNames, betweenNames, factors, depName, levelOrder, dep };
}

/* Per-subject profiles across one within factor: averages over any other within factors / replicates. */
function buildProfile(long, xFac, levels, betweenNames) {
  const subj = long.subject, xv = long.within[xFac], y = long.y;
  const groupKey = (i) => betweenNames.length ? betweenNames.map((b) => String(long.between[b][i])).join(" · ") : null;
  const bySubj = new Map();
  for (let i = 0; i < subj.length; i++) { const s = subj[i]; if (!bySubj.has(s)) bySubj.set(s, { group: groupKey(i), cells: new Map() }); const c = bySubj.get(s).cells, L = String(xv[i]), e = c.get(L) || [0, 0]; e[0] += y[i]; e[1]++; c.set(L, e); }
  const subjects = [];
  for (const [s, o] of bySubj) subjects.push({ id: s, group: o.group, ys: levels.map((L) => { const e = o.cells.get(String(L)); return e ? e[0] / e[1] : null; }) });
  return subjects;
}

function ConfigLine({ a, depName }) {
  const c = a.cfg || DEFAULT_ANOVA_CFG;
  const errLabel = { none: "no error bars", se: "± std error", sd: "± std dev", ci: `± ${Math.round((1 - c.alpha) * 100)}% CI` }[c.errorBars];
  return <div style={{ fontSize: 10, color: "#456" }}>{depName} · α = {c.alpha} · {c.effects === "all" ? "all effects" : "highest-order effect"} · {errLabel}</div>;
}

function renderAnovaTable(a, colById, compactById, rows, onOpt) {
  const m = anovaModel(a, colById, compactById, rows);
  if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
  const ssType = a.opts && a.opts.ssType === "II" ? "II" : "III";
  let res; try { res = anova({ ...m.long, ssType }); } catch (e) { return <div style={{ color: "#a00" }}>Could not compute: {String(e.message || e)}</div>; }
  const alpha = (a.cfg || DEFAULT_ANOVA_CFG).alpha;
  const designLabel = `${m.betweenNames.length ? m.betweenNames.join(" × ") + " (between)" : ""}${m.betweenNames.length && m.withinNames.length ? " × " : ""}${m.withinNames.length ? m.withinNames.join(" × ") + " (within)" : ""} on ${m.depName}`;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <div style={{ fontSize: 10, color: "#456" }}>{designLabel}</div>
      <div style={{ fontSize: 10, color: "#456" }}>N subjects = {res.Nsubj} · within cells/subject = {res.nWithinCells} · between cells = {res.nBetweenCells}</div>
      {res.betweenUnbalanced && m.withinNames.length === 0 && onOpt && (
        <label onClick={(e) => e.stopPropagation()} style={{ fontSize: 10, color: "#444", display: "inline-flex", alignItems: "center", gap: 4 }}>Unbalanced SS:
          <select value={ssType} onChange={(e) => onOpt(a.id, { ssType: e.target.value })} style={{ fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" }}>
            <option value="III">Type III (default)</option>
            <option value="II">Type II</option>
          </select>
        </label>
      )}
      {res.messages.map((mm, i) => <div key={i} style={{ color: "#a40", fontSize: 10 }}>⚠ {mm}</div>)}
      {onOpt && (
        <div style={{ display: "flex", gap: 10, fontSize: 10, color: "#444", alignItems: "center", flexWrap: "wrap" }} onClick={(e) => e.stopPropagation()}>
          <label style={{ display: "inline-flex", alignItems: "center", gap: 4 }}><input type="checkbox" checked={!!(a.opts && a.opts.effSize)} onChange={(e) => onOpt(a.id, { effSize: e.target.checked })} /> effect sizes</label>
          {a.opts && a.opts.effSize && <label style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>CI:
            <select value={(a.opts && a.opts.effConf) || 0.9} onChange={(e) => onOpt(a.id, { effConf: Number(e.target.value) })} style={{ fontFamily: FONT, fontSize: 10, border: `1px solid ${PLAT.dark}`, background: "#fff" }}>
              <option value={0.9}>90%</option><option value={0.95}>95%</option>
            </select>
          </label>}
        </div>
      )}
      {(() => {
        const hasSph = res.sources.some((s) => s.pGG != null);
        const cols = hasSph ? ["Source", "df", "Sum of Sq.", "Mean Sq.", "F", "P", "P (G–G)", "P (H–F)"] : ["Source", "df", "Sum of Sq.", "Mean Sq.", "F", "P"];
        const expRows = res.sources.map((s) => { const isErr = s.type === "error"; const base = [s.name, s.df, fmt(s.ss, 4), fmt(s.ms, 4), isErr ? "" : fmt(s.F, 4), isErr ? "" : fmtP(s.p)]; return hasSph ? [...base, s.pGG != null ? fmtP(s.pGG) : "", s.pHF != null ? fmtP(s.pHF) : ""] : base; });
        expRows.push(hasSph ? ["Total", res.total.df, fmt(res.total.ss, 4), "", "", "", "", ""] : ["Total", res.total.df, fmt(res.total.ss, 4), "", "", ""]);
        return (
          <div style={{ position: "relative", display: "inline-block" }} className="exp-anova">
            <div style={{ position: "absolute", top: -3, right: 0, transform: "translateY(-100%)", display: "flex", gap: 3, zIndex: 6 }} onClick={(e) => e.stopPropagation()}>
              <span style={tableExpBtn} onClick={() => copyTable(cols, expRows)} title="Copy as tab-separated (paste into Excel)">copy</span>
              <span style={tableExpBtn} onClick={() => downloadTableCSV(cols, expRows, "anova")} title="Download CSV">CSV</span>
            </div>
      <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
        <thead>
          <tr>{cols.map((h, i) => (
            <th key={i} style={{ border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", textAlign: i === 0 ? "left" : "right", fontWeight: "bold" }}>{h}</th>
          ))}</tr>
        </thead>
        <tbody>
          {res.sources.map((s, i) => {
            const isErr = s.type === "error"; const sig = s.p != null && s.p < alpha;
            return (
              <tr key={i}>
                <td style={{ ...anvCell, textAlign: "left", paddingLeft: isErr ? 18 : 8, fontStyle: isErr ? "italic" : "normal", color: isErr ? "#667" : "#000" }}>{s.name}</td>
                <td style={anvCell}>{s.df}</td><td style={anvCell}>{fmt(s.ss, 2)}</td><td style={anvCell}>{fmt(s.ms, 2)}</td>
                <td style={{ ...anvCell, fontWeight: sig ? "bold" : "normal" }}>{isErr ? "" : fmt(s.F, 3)}</td>
                <td style={{ ...anvCell, fontWeight: sig ? "bold" : "normal", color: sig ? "#070" : "#000" }}>{isErr ? "" : fmtP(s.p)}</td>
                {hasSph && <td style={{ ...anvCell, color: s.pGG != null && s.pGG < alpha ? "#070" : "#445" }}>{s.pGG != null ? fmtP(s.pGG) : ""}</td>}
                {hasSph && <td style={{ ...anvCell, color: s.pHF != null && s.pHF < alpha ? "#070" : "#445" }}>{s.pHF != null ? fmtP(s.pHF) : ""}</td>}
              </tr>
            );
          })}
          <tr>
            <td style={{ ...anvCell, textAlign: "left", fontWeight: "bold" }}>Total</td>
            <td style={{ ...anvCell, fontWeight: "bold" }}>{res.total.df}</td>
            <td style={{ ...anvCell, fontWeight: "bold" }}>{fmt(res.total.ss, 2)}</td>
            <td style={anvCell}></td><td style={anvCell}></td><td style={anvCell}></td>
            {hasSph && <td style={anvCell}></td>}{hasSph && <td style={anvCell}></td>}
          </tr>
        </tbody>
      </table>
          </div>
        );
      })()}
      <div style={{ fontSize: 10, color: "#555" }}>Within-subjects effects use (effect × Subject) error terms; between effects use Subject(Groups). Bold P &lt; {alpha}.</div>
      {(res.sphericity || []).some((b) => !b.trivial) && (
        <div style={{ fontSize: 10, color: "#445", border: `1px solid ${PLAT.dark}`, padding: "4px 8px", background: "#fbfbf7" }}>
          <b>Sphericity (Mauchly's test) &amp; epsilon corrections</b>
          {res.sphericity.filter((b) => !b.trivial).map((b, i) => (
            <div key={i} style={{ marginTop: 2 }}>
              {b.unavailable ? `${b.block}: too few subjects to estimate (need error df ≥ ${b.df}).`
                : `${b.block}: Mauchly W = ${fmt(b.W, 4)}, χ²(${b.dfChi}) = ${fmt(b.chi2, 3)}, p = ${fmtP(b.pMauchly)}  ·  ε(G–G) = ${fmt(b.epsGG, 4)}, ε(H–F) = ${fmt(b.epsHF, 4)}`}
            </div>
          ))}
          <div style={{ marginTop: 3, color: "#778" }}>Per-effect orthonormal contrasts (SPSS/afex convention); mixed designs use the pooled within-cell covariance. When Mauchly's p &lt; {alpha}, prefer the G–G (conservative) or H–F corrected P. Effects with 1 df are unaffected by sphericity.</div>
        </div>
      )}
      {a.opts && a.opts.effSize && (() => {
        const conf = (a.opts && a.opts.effConf) || 0.9, ssT = res.total.ss, Ntot = res.total.df + 1;
        const effects = res.sources.filter((s) => s.type !== "error");
        const pureBetween = !res.sources.some((s) => s.type === "within");
        const fullyWithin = res.sources.some((s) => s.type === "within") && res.nBetweenCells === 1;
        const subjSrc = res.sources.find((s) => s.type === "error" && (s.name.startsWith("Subject") || s.name === "Residual"));
        const withinErrSum = res.sources.filter((s) => s.type === "error" && s.name.indexOf("\u00d7 Subj") >= 0).reduce((a2, s) => a2 + s.ss, 0);
        const Kgen = (subjSrc ? subjSrc.ss : 0) + withinErrSum;
        const head = ["Effect", "η²", "partial η²", `partial η² ${Math.round(conf * 100)}% CI`];
        if (pureBetween) head.push("ω²", "partial ω²");
        if (fullyWithin) head.push("generalized η²");
        const rows = effects.map((s) => {
          const eta2 = s.ss / ssT, peta = s.ss / (s.ss + s.errSS);
          const ci = etaSqPartialCI(s.F, s.df, s.errDF, conf);
          const row = [s.name, fmt(eta2, 4), fmt(peta, 4), `[${fmt(ci[0], 4)}, ${fmt(ci[1], 4)}]`];
          if (pureBetween) { const om = (s.ss - s.df * s.errMS) / (ssT + s.errMS), pom = (s.ss - s.df * s.errMS) / (s.ss + (Ntot - s.df) * s.errMS); row.push(fmt(om, 4), fmt(pom, 4)); }
          if (fullyWithin) row.push(fmt(s.ss / (s.ss + Kgen), 4));
          return row;
        });
        return (
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <div style={{ fontWeight: "bold", fontSize: 11 }}>Effect sizes</div>
            <StatTable head={head} rows={rows} />
            <div style={{ fontSize: 10, color: "#555" }}>
              partial η² = SS/(SS + SS<sub>error</sub>) (SPSS convention; matches the standard np2). η² = SS/SS<sub>total</sub>. CI for partial η² uses the noncentral F (Smithson). {pureBetween ? "ω²/partial ω² are the less-biased estimators for between-subjects fixed effects. " : ""}{fullyWithin ? "Generalized η² (Olejnik–Algina) puts all measured error in the denominator and is the recommended comparable measure for repeated-measures designs. " : ""}{!pureBetween && !fullyWithin ? "For this mixed design, partial η² and its CI are reported; ω² and generalized η² depend on additional design assumptions. " : ""}A 90% CI pairs with the conventional one-sided F-test at α = .05.
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function renderAnovaMeans(a, colById, compactById, rows) {
  const m = anovaModel(a, colById, compactById, rows);
  if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
  const cfg = { ...DEFAULT_ANOVA_CFG, ...a.cfg };
  const effects = effectList(m.factors, cfg.effects);
  const errCol = { none: null, se: "Std. Error", sd: "Std. Dev.", ci: `${Math.round((1 - cfg.alpha) * 100)}% CI ±` }[cfg.errorBars];
  const sortRows = (rws, fac) => rws.slice().sort((x, y) => {
    for (let i = 0; i < fac.length; i++) { const oi = m.levelOrder[fac[i]].indexOf(x.levels[i]) - m.levelOrder[fac[i]].indexOf(y.levels[i]); if (oi) return oi; }
    return 0;
  });
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      <ConfigLine a={a} depName={m.depName} />
      {effects.map((E, ei) => {
        const rws = sortRows(cellMeans(m.long, E, cfg.alpha, cfg.errorBars), E);
        const head = [...E, "Count", "Mean"].concat(errCol ? [errCol] : []);
        const expRows = rws.map((r) => [...r.levels, r.n, fmt(r.mean, 2)].concat(errCol ? [fmt(r.err, 2)] : []));
        return (
          <div key={ei}>
            <div style={{ fontWeight: "bold", marginBottom: 2 }}>{E.join(" × ")}</div>
            <ExportWrap head={head} rows={expRows} name="means">
            <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
              <thead><tr>{head.map((h, i) => <th key={i} style={{ border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", textAlign: i < E.length ? "left" : "right", fontWeight: "bold" }}>{h}</th>)}</tr></thead>
              <tbody>
                {rws.map((r, ri) => (
                  <tr key={ri}>
                    {r.levels.map((lv, i) => <td key={i} style={{ border: "1px solid #ddd", padding: "1px 8px" }}>{lv}</td>)}
                    <td style={anvCell}>{r.n}</td>
                    <td style={anvCell}>{fmt(r.mean, 2)}</td>
                    {errCol && <td style={anvCell}>{fmt(r.err, 2)}</td>}
                  </tr>
                ))}
              </tbody>
            </table>
            </ExportWrap>
          </div>
        );
      })}
      {effects.length === 0 && <div style={{ color: "#999" }}>No factors to summarize.</div>}
    </div>
  );
}

/* Chart facets for one effect. Default: x = outer factor (2nd-last), series = inner
   factor (last, drawn side-by-side); remaining factors become facet panels. `swap`
   exchanges the x-axis factor and the side-by-side (series) factor. */
function graphModel(E, rws, levelOrder, swap) {
  let xFac, seriesFac;
  const facetFacs = E.slice(0, Math.max(0, E.length - 2));
  if (E.length === 1) { xFac = E[0]; seriesFac = null; }
  else { xFac = E[E.length - 2]; seriesFac = E[E.length - 1]; if (swap) { const t = xFac; xFac = seriesFac; seriesFac = t; } }
  const xLevels = levelOrder[xFac];
  const seriesLevels = seriesFac ? levelOrder[seriesFac] : [null];
  const facetCombos = cartesian(facetFacs.map((f) => levelOrder[f]));
  const rowMap = new Map(rws.map((r) => [r.levels.join("\u0001"), r]));
  const facets = (facetCombos.length ? facetCombos : [[]]).map((combo) => {
    const label = facetFacs.map((f, i) => `${f} = ${combo[i]}`).join(", ");
    const data = xLevels.map((xl) => {
      const point = { x: xl };
      seriesLevels.forEach((sl) => {
        const lv = E.map((f) => (f === xFac ? xl : f === seriesFac ? sl : combo[facetFacs.indexOf(f)]));
        const r = rowMap.get(lv.join("\u0001"));
        const key = sl == null ? "mean" : sl;
        point[key] = r ? r.mean : null;
        point[key + "__e"] = r ? r.err : 0;
      });
      return point;
    });
    return { label, data, seriesKeys: seriesLevels.map((sl) => (sl == null ? "mean" : sl)) };
  });
  return { facets, xFac, seriesFac };
}

/* =========================================================================
   CLEAN SVG PLOT RENDERER — our own publication-grade vector output
   (every axis, tick, series path, symbol, error bar and label is a plain
   SVG element, so the on-screen graph IS the export-ready artwork)
   ========================================================================= */
function niceNum(range, round) {
  const r = range || 1; const exp = Math.floor(Math.log10(r)); const f = r / Math.pow(10, exp);
  const nf = round ? (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) : (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10);
  return nf * Math.pow(10, exp);
}
function rangeTicks(lo, hi, step) { const t = []; for (let v = lo; v <= hi + step * 0.5; v += step) t.push(+v.toFixed(10)); return t; }
function niceScale(min, max, maxTicks) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) { min = 0; max = 1; }
  if (min === max) { min -= 0.5; max += 0.5; }
  const range = niceNum(max - min, false); const step = niceNum(range / Math.max(1, (maxTicks || 5) - 1), true);
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  return { lo, hi, step, ticks: rangeTicks(lo, hi, step) };
}
function yScaleFor(data, keys, showErr, plot) {
  let mn = Infinity, mx = -Infinity;
  data.forEach((p) => keys.forEach((k) => { const v = p[k]; if (v == null || Number.isNaN(v)) return; const e = showErr ? (p[k + "__e"] || 0) : 0; mn = Math.min(mn, v - e); mx = Math.max(mx, v + e); }));
  if (!Number.isFinite(mn)) { mn = 0; mx = 1; }
  const num = (v) => (v !== "" && v != null && Number.isFinite(+v) ? +v : null);
  const pMin = num(plot.yMin), pMax = num(plot.yMax), pStep = num(plot.yTickInterval) > 0 ? +plot.yTickInterval : null;
  let lo, hi, ticks;
  if (pMin != null || pMax != null || pStep) {
    const ns = niceScale(pMin != null ? pMin : mn, pMax != null ? pMax : mx, 5);
    lo = pMin != null ? pMin : ns.lo; hi = pMax != null ? pMax : ns.hi;
    ticks = pStep ? rangeTicks(lo, hi, pStep) : ns.ticks.filter((t) => t >= lo - 1e-9 && t <= hi + 1e-9);
  } else { const ns = niceScale(mn, mx, 5); lo = ns.lo; hi = ns.hi; ticks = ns.ticks; }
  if (lo === hi) hi = lo + 1;
  return { lo, hi, ticks };
}
function symbolEl(shape, cx, cy, r, fill, key) {
  if (shape === "none") return null;
  if (shape === "square") return <rect key={key} x={cx - r} y={cy - r} width={2 * r} height={2 * r} fill={fill} />;
  if (shape === "diamond") return <path key={key} d={`M${cx} ${cy - r}L${cx + r} ${cy}L${cx} ${cy + r}L${cx - r} ${cy}Z`} fill={fill} />;
  if (shape === "triangle") return <path key={key} d={`M${cx} ${cy - r}L${cx + r} ${cy + r}L${cx - r} ${cy + r}Z`} fill={fill} />;
  return <circle key={key} cx={cx} cy={cy} r={r} fill={fill} />;
}
function errBarEls(key, cx, yA, yB, color) {
  const c = color || "#333", cap = 3;
  return [
    <line key={key + "v"} x1={cx} y1={yA} x2={cx} y2={yB} stroke={c} strokeWidth={1} />,
    <line key={key + "a"} x1={cx - cap} y1={yA} x2={cx + cap} y2={yA} stroke={c} strokeWidth={1} />,
    <line key={key + "b"} x1={cx - cap} y1={yB} x2={cx + cap} y2={yB} stroke={c} strokeWidth={1} />,
  ];
}
const fmtTickVal = (v) => (Math.abs(v) >= 1e4 || (Math.abs(v) < 1e-2 && v !== 0) ? v.toExponential(1) : (+v.toFixed(4)).toString());
// One clean chart (one effect, one facet). Returns a self-contained <svg>.
function CleanChart({ data, seriesKeys, kind, showErr, plot, depName, hasSeries, onCommitSize, onCommitLegend }) {
  const nL = data.length || 1, nS = seriesKeys.length || 1;
  const ys = yScaleFor(data, seriesKeys, showErr, plot);
  const xLabels = data.map((p) => String(p.x));
  const maxLen = xLabels.reduce((mx, s) => Math.max(mx, s.length), 0);
  const skip = Math.max(0, Math.round(+plot.xInterval || 0));
  const showLab = (i) => skip === 0 || i % (skip + 1) === 0;
  const xAngle = Math.max(0, Math.min(90, Math.round(+plot.xAngle || 0))); // 0 = horizontal (default)
  const FS = +plot.tickSize > 0 ? +plot.tickSize : 9, axW = +plot.axisWidth || 1, tkW = +plot.tickWidth || 1;
  const LBL = +plot.labelSize > 0 ? +plot.labelSize : 10, LBLW = plot.labelBold ? "bold" : "normal", LBLI = plot.labelItalic ? "italic" : "normal";
  const LEG = +plot.legendSize > 0 ? +plot.legendSize : 9, LEGW = plot.legendBold ? "bold" : "normal", LEGI = plot.legendItalic ? "italic" : "normal";
  const legCw = LEG * 0.62, legRow = LEG + 5;
  const legendPos = hasSeries ? (plot.legend || "top") : "none";
  const showLeg = legendPos !== "none";
  const svgRef = useRef(null);
  const [legDrag, setLegDrag] = useState(null);
  const legXY = legDrag || (plot.legendXY && typeof plot.legendXY === "object" ? plot.legendXY : null);
  const legFree = showLeg && !!legXY;
  const legTop = showLeg && legendPos === "top" && !legFree;
  const legRight = showLeg && legendPos === "right" && !legFree;
  const legLabels = seriesKeys.map((k) => (k === "mean" ? depName : k));
  const legMaxLen = legLabels.reduce((m, s) => Math.max(m, String(s).length), 0);
  const swW = kind === "bar" ? 13 : 18;
  const legColW = legRight ? swW + Math.ceil(legMaxLen * legCw) + 12 : 0;
  const padTop = 8 + (plot.title ? 14 : 0) + (legTop ? LEG + 7 : 0);
  const padBottom = (xAngle === 0 ? FS + 13 : xAngle >= 80 ? FS + 5 + Math.ceil(maxLen * FS * 0.62) : FS + 5 + Math.ceil(maxLen * FS * 0.47)) + (plot.xLabel ? LBL + 4 : 0);
  const padLeft = 44 + (plot.yLabel ? LBL + 4 : 0), padRight = 16 + legColW;
  const bandW = kind === "bar" ? Math.max(30, nS * 16 + 16) : 46;
  const autoPlotW = Math.max(180, nL * bandW), autoPlotH = 168;
  const autoW = padLeft + autoPlotW + padRight, autoH = padTop + autoPlotH + padBottom;
  const [size, setSize] = useState({ w: +plot.width > 0 ? +plot.width : autoW, h: +plot.height > 0 ? +plot.height : autoH });
  useEffect(() => { setSize({ w: +plot.width > 0 ? +plot.width : autoW, h: +plot.height > 0 ? +plot.height : autoH }); }, [plot.width, plot.height, autoW, autoH]);
  useEffect(() => { if (!plot.legendXY) setLegDrag(null); }, [plot.legendXY]);
  const minW = padLeft + 140 + padRight, minH = padTop + 90 + padBottom;
  const plotW = Math.max(140, size.w - padLeft - padRight), plotH = Math.max(90, size.h - padTop - padBottom);
  const W = padLeft + plotW + padRight, H = padTop + plotH + padBottom;
  const x0 = padLeft, x1 = padLeft + plotW, yTop = padTop, yBot = padTop + plotH;
  const onGripDown = (e) => {
    e.preventDefault(); e.stopPropagation();
    const sx = e.clientX, sy = e.clientY, sw = W, sh = H;
    const move = (ev) => setSize({ w: Math.max(minW, sw + (ev.clientX - sx)), h: Math.max(minH, sh + (ev.clientY - sy)) });
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); setSize((s) => { onCommitSize && onCommitSize(Math.round(s.w), Math.round(s.h)); return s; }); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  };
  const startLegDrag = (e, ox, oy) => {
    e.preventDefault(); e.stopPropagation();
    const svg = svgRef.current; const rect = svg ? svg.getBoundingClientRect() : null;
    const kx = rect && rect.width ? W / rect.width : 1, ky = rect && rect.height ? H / rect.height : 1;
    const px0 = e.clientX, py0 = e.clientY;
    const move = (ev) => setLegDrag({ x: Math.round(ox + (ev.clientX - px0) * kx), y: Math.round(oy + (ev.clientY - py0) * ky) });
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); setLegDrag((d) => { if (d && onCommitLegend) onCommitLegend(d); return d; }); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  };
  const yOf = (v) => yBot - ((v - ys.lo) / (ys.hi - ys.lo)) * plotH;
  const bw = plotW / nL, center = (i) => x0 + (i + 0.5) * bw;
  const els = [];
  els.push(<rect key="bg" x={0} y={0} width={W} height={H} fill="#ffffff" />);
  if (plot.grid) {
    if (plot.gridH) ys.ticks.forEach((t, i) => { const y = yOf(t); if (y >= yTop - 0.5 && y <= yBot + 0.5) els.push(<line key={"gh" + i} x1={x0} y1={y} x2={x1} y2={y} stroke="#e0e0e0" strokeDasharray={plot.gridDashed ? "3 3" : undefined} />); });
    if (plot.gridV) data.forEach((p, i) => els.push(<line key={"gv" + i} x1={center(i)} y1={yTop} x2={center(i)} y2={yBot} stroke="#e0e0e0" strokeDasharray={plot.gridDashed ? "3 3" : undefined} />));
  }
  if (kind === "bar") {
    const groupW = bw * 0.8, barW = groupW / nS;
    seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si); data.forEach((p, i) => {
      const v = p[k]; if (v == null || Number.isNaN(v)) return;
      const bx = center(i) - groupW / 2 + si * barW, top = yOf(v);
      els.push(<rect key={`b${si}-${i}`} x={bx} y={Math.min(top, yBot)} width={barW * 0.9} height={Math.abs(yBot - top)} fill={st.color} stroke="#00000033" />);
      if (showErr) { const e = p[k + "__e"] || 0; els.push(...errBarEls(`be${si}-${i}`, bx + barW * 0.45, yOf(v + e), yOf(v - e))); }
    }); });
  } else {
    seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si);
      const lw = (st.lineWidth === "" || st.lineWidth == null) ? 2 : +st.lineWidth;
      const ss = (st.symbolSize === "" || st.symbolSize == null) ? 3.5 : +st.symbolSize;
      let d = "", pen = false;
      data.forEach((p, i) => { const v = p[k]; if (v == null || Number.isNaN(v)) { pen = false; return; } d += (pen ? "L" : "M") + center(i) + " " + yOf(v) + " "; pen = true; });
      if (d && lw > 0) els.push(<path key={"ln" + si} d={d.trim()} fill="none" stroke={st.color} strokeWidth={lw} />);
      if (showErr) data.forEach((p, i) => { const v = p[k]; if (v == null || Number.isNaN(v)) return; const e = p[k + "__e"] || 0; els.push(...errBarEls(`le${si}-${i}`, center(i), yOf(v + e), yOf(v - e), st.color)); });
      data.forEach((p, i) => { const v = p[k]; if (v == null || Number.isNaN(v)) return; const s = symbolEl(st.symbol, center(i), yOf(v), ss, st.symbolColor, `sy${si}-${i}`); if (s) els.push(s); });
    });
  }
  // axes
  els.push(<line key="axb" x1={x0} y1={yBot} x2={x1} y2={yBot} stroke="#333" strokeWidth={axW} />);
  els.push(<line key="axl" x1={x0} y1={yTop} x2={x0} y2={yBot} stroke="#333" strokeWidth={axW} />);
  if (plot.frame) { els.push(<line key="axt" x1={x0} y1={yTop} x2={x1} y2={yTop} stroke="#333" strokeWidth={axW} />); els.push(<line key="axr" x1={x1} y1={yTop} x2={x1} y2={yBot} stroke="#333" strokeWidth={axW} />); }
  // y ticks + labels
  ys.ticks.forEach((t, i) => { const y = yOf(t); if (y < yTop - 0.5 || y > yBot + 0.5) return; els.push(<line key={"yt" + i} x1={x0 - 4} y1={y} x2={x0} y2={y} stroke="#333" strokeWidth={tkW} />); els.push(<text key={"ytl" + i} x={x0 - 6} y={y + 3} textAnchor="end" fontSize={FS} fill="#222">{fmtTickVal(t)}</text>); });
  // x ticks + labels
  data.forEach((p, i) => { const cx = center(i); els.push(<line key={"xt" + i} x1={cx} y1={yBot} x2={cx} y2={yBot + 4} stroke="#333" strokeWidth={tkW} />); if (showLab(i)) els.push(xAngle === 0
    ? <text key={"xl" + i} x={cx} y={yBot + 13} fontSize={FS} fill="#222" textAnchor="middle">{xLabels[i]}</text>
    : <text key={"xl" + i} x={cx} y={yBot + (xAngle >= 80 ? 4 : 7)} fontSize={FS} fill="#222" textAnchor="end" transform={`rotate(-${xAngle} ${cx} ${yBot + (xAngle >= 80 ? 4 : 7)})`}>{xLabels[i]}</text>); });
  // axis titles + chart title
  if (plot.xLabel) els.push(<text key="xlab" x={(x0 + x1) / 2} y={H - 3} fontSize={LBL} fontWeight={LBLW} fontStyle={LBLI} fill="#111" textAnchor="middle">{plot.xLabel}</text>);
  if (plot.yLabel) els.push(<text key="ylab" x={11} y={(yTop + yBot) / 2} fontSize={LBL} fontWeight={LBLW} fontStyle={LBLI} fill="#111" textAnchor="middle" transform={`rotate(-90 11 ${(yTop + yBot) / 2})`}>{plot.yLabel}</text>);
  if (plot.title) els.push(<text key="title" x={(x0 + x1) / 2} y={12} fontSize={11} fontWeight="bold" fill="#000" textAnchor="middle">{plot.title}</text>);
  // legend (positionable + draggable)
  if (showLeg) {
    const entry = (st, sx, sy, kp) => {
      if (kind === "bar") els.push(<rect key={kp + "r"} x={sx} y={sy - 5} width={10} height={10} fill={st.color} />);
      else { const llw = (st.lineWidth === "" || st.lineWidth == null) ? 2 : +st.lineWidth; if (llw > 0) els.push(<line key={kp + "l"} x1={sx} y1={sy} x2={sx + 14} y2={sy} stroke={st.color} strokeWidth={llw} />); const sm = symbolEl(st.symbol, sx + 7, sy, Math.min((st.symbolSize === "" || st.symbolSize == null) ? 3.5 : +st.symbolSize, 4.5), st.symbolColor, kp + "s"); if (sm) els.push(sm); }
    };
    const tp = { fontSize: LEG, fontWeight: LEGW, fontStyle: LEGI, fill: "#222" };
    let hit = null;
    if (legFree) {
      const boxW = swW + Math.ceil(legMaxLen * legCw) + 14, boxH = nS * legRow + 6;
      const bx = legXY.x, by = legXY.y;
      els.push(<rect key="lgbg" x={bx} y={by} width={boxW} height={boxH} fill="#ffffff" fillOpacity={0.82} stroke="#bbb" />);
      seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si); const yy = by + 10 + si * legRow; entry(st, bx + 6, yy, "lg" + si); els.push(<text key={"lgt" + si} x={bx + 6 + swW} y={yy + 3} {...tp}>{legLabels[si]}</text>); });
      hit = { x: bx, y: by, w: boxW, h: boxH, ox: bx, oy: by };
    } else if (legendPos === "top") {
      let lx = x0; const ly = plot.title ? 24 : 11;
      seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si); entry(st, lx, ly, "lg" + si); const tx = lx + swW; els.push(<text key={"lgt" + si} x={tx} y={ly + 2} {...tp}>{legLabels[si]}</text>); lx = tx + String(legLabels[si]).length * legCw + 12; });
      hit = { x: x0 - 2, y: ly - 9, w: Math.max(24, lx - x0), h: legRow + 5, ox: x0 - 2, oy: ly - 9 };
    } else if (legendPos === "right") {
      const lx = x1 + 12, ly0 = yTop + 8;
      seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si); const yy = ly0 + si * legRow; entry(st, lx, yy, "lg" + si); els.push(<text key={"lgt" + si} x={lx + swW} y={yy + 3} {...tp}>{legLabels[si]}</text>); });
      hit = { x: lx - 2, y: ly0 - 9, w: swW + Math.ceil(legMaxLen * legCw) + 10, h: nS * legRow + 4, ox: lx - 2, oy: ly0 - 9 };
    } else {
      const boxW = swW + Math.ceil(legMaxLen * legCw) + 14, boxH = nS * legRow + 6;
      const bx = legendPos === "tr" || legendPos === "br" ? x1 - boxW - 6 : x0 + 6;
      const by = legendPos === "bl" || legendPos === "br" ? yBot - boxH - 6 : yTop + 6;
      els.push(<rect key="lgbg" x={bx} y={by} width={boxW} height={boxH} fill="#ffffff" fillOpacity={0.82} stroke="#bbb" />);
      seriesKeys.forEach((k, si) => { const st = seriesStyleOf(plot, k, si); const yy = by + 10 + si * legRow; entry(st, bx + 6, yy, "lg" + si); els.push(<text key={"lgt" + si} x={bx + 6 + swW} y={yy + 3} {...tp}>{legLabels[si]}</text>); });
      hit = { x: bx, y: by, w: boxW, h: boxH, ox: bx, oy: by };
    }
    if (onCommitLegend && hit) els.push(<rect key="lghit" x={hit.x} y={hit.y} width={hit.w} height={hit.h} fill="transparent" style={{ cursor: "move" }} onMouseDown={(e) => startLegDrag(e, hit.ox, hit.oy)}><title>drag to reposition legend</title></rect>);
  }
  return (
    <div style={{ position: "relative", width: W, height: H }}>
      <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} width="100%" height="100%" preserveAspectRatio="xMinYMin meet" fontFamily={FONT} style={{ display: "block", background: "#fff" }}>{els}</svg>
      {onCommitSize && <div onMouseDown={onGripDown} title="drag to resize graph" style={{ position: "absolute", right: 0, bottom: 0, width: 16, height: 16, zIndex: 6, cursor: "nwse-resize", background: `linear-gradient(135deg, transparent 40%, ${PLAT.dark} 40%, ${PLAT.dark} 55%, transparent 55%, transparent 68%, ${PLAT.dark} 68%, ${PLAT.dark} 83%, transparent 83%)` }} />}
    </div>
  );
}

/* ---- graph export / copy (the clean SVG is the master; everything derives from it) ---- */
function svgToString(svgEl) {
  const clone = svgEl.cloneNode(true);
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  clone.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
  const vb = (clone.getAttribute("viewBox") || "").split(/\s+/).map(Number);
  if (vb.length === 4) { clone.setAttribute("width", vb[2]); clone.setAttribute("height", vb[3]); }
  return new XMLSerializer().serializeToString(clone); // no XML prolog (more reliable as an <img> source)
}
function downloadBlob(blob, name) { const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
// Native-save / native-copy seams. In the browser build these are no-ops (return false), so the
// callers fall back to the standard <a download> and Clipboard API. The desktop (Tauri) kit
// replaces these with versions that use the OS save dialog + filesystem and the system clipboard.
async function nativeSaveBytes(blob, name) {
  try {
    const ext = (String(name).split(".").pop() || "").toLowerCase();
    const p = await tauriSaveDialog({ defaultPath: name, filters: ext ? [{ name: ext.toUpperCase() + " file", extensions: [ext] }] : [] });
    if (!p) return true; // user cancelled — handled (don't also fire a phantom web download)
    await writeFile(p, new Uint8Array(await blob.arrayBuffer()));
    return true;
  } catch (e) { try { window.alert("Save failed: " + (e && e.message ? e.message : e)); } catch (_) {} return true; }
}
async function nativeCopyImage(blob) {
  try {
    const img = await TauriImage.fromBytes(new Uint8Array(await blob.arrayBuffer()));
    await tauriWriteImage(img);
    return true;
  } catch (e) { return false; }
}
// rasterize via a data: URL (blob: URLs frequently fail to load as <img> inside sandboxed iframes)
function rasterizeSvg(svgStr, w, h, scale, type, quality) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      try {
        const c = document.createElement("canvas"); c.width = Math.max(1, Math.round(w * scale)); c.height = Math.max(1, Math.round(h * scale));
        const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0, c.width, c.height);
        c.toBlob((b) => (b ? resolve(b) : reject(new Error("canvas toBlob returned null"))), type, quality);
      } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error("the browser could not render the SVG to an image"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);
  });
}
// minimal single-image PDF (embeds a JPEG via DCTDecode) — works fully offline, no pop-up needed
function buildImagePdf(jpegBytes, ptW, ptH, pxW, pxH) {
  const enc = (s) => { const a = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i) & 0xff; return a; };
  const parts = []; let len = 0; const off = [];
  const push = (u8) => { parts.push(u8); len += u8.length; };
  const pushStr = (s) => push(enc(s));
  const obj = (n) => { off[n] = len; };
  pushStr("%PDF-1.4\n%\xFF\xFF\xFF\xFF\n");
  obj(1); pushStr("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  obj(2); pushStr("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");
  obj(3); pushStr(`3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${ptW} ${ptH}] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>\nendobj\n`);
  const content = `q\n${ptW} 0 0 ${ptH} 0 0 cm\n/Im0 Do\nQ\n`;
  obj(4); pushStr(`4 0 obj\n<< /Length ${content.length} >>\nstream\n${content}endstream\nendobj\n`);
  obj(5); pushStr(`5 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pxW} /Height ${pxH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpegBytes.length} >>\nstream\n`);
  push(jpegBytes); pushStr("\nendstream\nendobj\n");
  const xrefAt = len;
  let xref = "xref\n0 6\n0000000000 65535 f \n";
  for (let n = 1; n <= 5; n++) xref += String(off[n]).padStart(10, "0") + " 00000 n \n";
  pushStr(xref);
  pushStr(`trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`);
  const out = new Uint8Array(len); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
function ExportFrame({ name, children }) {
  const ref = useRef(null);
  const [copied, setCopied] = useState("");
  const base = String(name || "graph").replace(/[^\w.-]+/g, "_").replace(/^_+|_+$/g, "") || "graph";
  const getSvg = () => ref.current && ref.current.querySelector("svg");
  const dims = (svg) => { const v = (svg.getAttribute("viewBox") || "0 0 100 100").split(/\s+/).map(Number); return { w: v[2], h: v[3] }; };
  const doSVG = async () => { const s = getSvg(); if (!s) return; const blob = new Blob([svgToString(s)], { type: "image/svg+xml" }); if (!(await nativeSaveBytes(blob, base + ".svg"))) downloadBlob(blob, base + ".svg"); };
  const doPNG = async () => { const s = getSvg(); if (!s) return; const { w, h } = dims(s); try { const blob = await rasterizeSvg(svgToString(s), w, h, 3, "image/png"); if (!(await nativeSaveBytes(blob, base + ".png"))) downloadBlob(blob, base + ".png"); } catch (e) { window.alert("PNG export failed: " + (e.message || e)); } };
  const doPDF = async () => { const s = getSvg(); if (!s) return; const { w, h } = dims(s); try { const scale = 3; const jpg = await rasterizeSvg(svgToString(s), w, h, scale, "image/jpeg", 0.95); const bytes = new Uint8Array(await jpg.arrayBuffer()); const pdf = buildImagePdf(bytes, Math.round(w), Math.round(h), Math.round(w * scale), Math.round(h * scale)); const blob = new Blob([pdf], { type: "application/pdf" }); if (!(await nativeSaveBytes(blob, base + ".pdf"))) downloadBlob(blob, base + ".pdf"); } catch (e) { window.alert("PDF export failed: " + (e.message || e)); } };
  const doCopy = async () => {
    const s = getSvg(); if (!s) return; const { w, h } = dims(s); const svgStr = svgToString(s);
    try { const png = await rasterizeSvg(svgStr, w, h, 2, "image/png"); if (await nativeCopyImage(png)) { setCopied("img"); setTimeout(() => setCopied(""), 1600); return; } } catch (e) { /* fall through */ }
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        const blob = await rasterizeSvg(svgStr, w, h, 2, "image/png");
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        setCopied("img"); setTimeout(() => setCopied(""), 1600); return;
      }
    } catch (e) { /* fall through */ }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(svgStr);
        setCopied("svg"); setTimeout(() => setCopied(""), 2200); return;
      }
    } catch (e) { /* fall through */ }
    setCopied("no"); setTimeout(() => setCopied(""), 2600);
  };
  const btn = { cursor: "pointer", fontSize: 10, border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "0 6px", borderRadius: 2, ...bevelOut };
  return (
    <div>
      <div style={{ display: "flex", gap: 5, marginBottom: 3, alignItems: "center" }}>
        <span onClick={doCopy} style={btn} title="copy the graph image to the clipboard, then paste into PowerPoint, Illustrator, etc.">{copied === "img" ? "Copied ✓" : copied === "svg" ? "Copied SVG ✓" : copied === "no" ? "Clipboard N/A" : "⧉ Copy"}</span>
        <span onClick={doSVG} style={btn} title="download editable vector SVG (open/place in Illustrator; insert in PowerPoint → Convert to Shape)">SVG</span>
        <span onClick={doPNG} style={btn} title="download high-resolution PNG (3×)">PNG</span>
        <span onClick={doPDF} style={btn} title="download a PDF of the graph">PDF</span>
      </div>
      <div ref={ref}>{children}</div>
    </div>
  );
}
function ChartWithExport({ name, ...chartProps }) {
  return <ExportFrame name={name}><CleanChart {...chartProps} /></ExportFrame>;
}

/* Normal Q–Q plot (sample quantiles vs theoretical normal quantiles + quartile reference line). */
function QQPlot({ data, title }) {
  const W = 380, H = 300, ml = 54, mr = 14, mt = 24, mb = 42;
  const s = data.map(Number).sort((a, b) => a - b), n = s.length;
  const theo = s.map((_, i) => invNorm((i + 1 - 0.375) / (n + 0.25)));
  const qx1 = invNorm(0.25), qx2 = invNorm(0.75), qy1 = quantileSorted(s, 0.25), qy2 = quantileSorted(s, 0.75);
  const slope = (qy2 - qy1) / ((qx2 - qx1) || 1), inter = qy1 - slope * qx1;
  const xmin = theo[0], xmax = theo[n - 1], ymin = s[0], ymax = s[n - 1];
  const px = (xmax - xmin) * 0.06 || 1, py = (ymax - ymin) * 0.06 || 1;
  const X0 = xmin - px, X1 = xmax + px, Y0 = ymin - py, Y1 = ymax + py;
  const sx = (v) => ml + (v - X0) / (X1 - X0) * (W - mr - ml);
  const sy = (v) => (H - mb) - (v - Y0) / (Y1 - Y0) * (H - mb - mt);
  const xt = [], yt = []; for (let i = 0; i <= 4; i++) { xt.push(X0 + i / 4 * (X1 - X0)); yt.push(Y0 + i / 4 * (Y1 - Y0)); }
  const f3 = (v) => (Math.abs(v) >= 100 ? v.toFixed(0) : v.toFixed(2));
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: W, height: H, fontFamily: FONT, background: "#fff", border: `1px solid ${PLAT.dark}` }}>
      {title && <text x={W / 2} y={15} textAnchor="middle" fontSize="12" fontWeight="bold">{title}</text>}
      {xt.map((t, i) => <g key={"x" + i}><line x1={sx(t)} y1={H - mb} x2={sx(t)} y2={mt} stroke="#eee" /><line x1={sx(t)} y1={H - mb} x2={sx(t)} y2={H - mb + 4} stroke="#000" /><text x={sx(t)} y={H - mb + 15} textAnchor="middle" fontSize="9">{f3(t)}</text></g>)}
      {yt.map((t, i) => <g key={"y" + i}><line x1={ml} y1={sy(t)} x2={W - mr} y2={sy(t)} stroke="#eee" /><line x1={ml - 4} y1={sy(t)} x2={ml} y2={sy(t)} stroke="#000" /><text x={ml - 7} y={sy(t) + 3} textAnchor="end" fontSize="9">{f3(t)}</text></g>)}
      <line x1={ml} y1={mt} x2={ml} y2={H - mb} stroke="#000" /><line x1={ml} y1={H - mb} x2={W - mr} y2={H - mb} stroke="#000" />
      <line x1={sx(X0)} y1={sy(slope * X0 + inter)} x2={sx(X1)} y2={sy(slope * X1 + inter)} stroke="#c00" strokeWidth="1.3" />
      {s.map((v, i) => <circle key={i} cx={sx(theo[i])} cy={sy(v)} r="2.6" fill="none" stroke="#1a3b6e" strokeWidth="1.1" />)}
      <text x={(ml + W - mr) / 2} y={H - 6} textAnchor="middle" fontSize="10">Theoretical Quantiles</text>
      <text x={13} y={(mt + H - mb) / 2} textAnchor="middle" fontSize="10" transform={`rotate(-90 13 ${(mt + H - mb) / 2})`}>Sample Quantiles</text>
    </svg>
  );
}

/* Generic diagnostic scatter (e.g. residuals vs fitted) with an optional horizontal reference line. */
function ScatterDiag({ points, xLabel, yLabel, title, hLine }) {
  const W = 380, H = 290, ml = 54, mr = 14, mt = 24, mb = 42;
  const xs = points.map((p) => p.x), ys = points.map((p) => p.y);
  let xmin = Math.min(...xs), xmax = Math.max(...xs), ymin = Math.min(...ys), ymax = Math.max(...ys);
  if (hLine != null) { ymin = Math.min(ymin, hLine); ymax = Math.max(ymax, hLine); }
  const px = (xmax - xmin) * 0.06 || 1, py = (ymax - ymin) * 0.08 || 1;
  const X0 = xmin - px, X1 = xmax + px, Y0 = ymin - py, Y1 = ymax + py;
  const sx = (v) => ml + (v - X0) / (X1 - X0) * (W - mr - ml);
  const sy = (v) => (H - mb) - (v - Y0) / (Y1 - Y0) * (H - mb - mt);
  const xt = [], yt = []; for (let i = 0; i <= 4; i++) { xt.push(X0 + i / 4 * (X1 - X0)); yt.push(Y0 + i / 4 * (Y1 - Y0)); }
  const f3 = (v) => (Math.abs(v) >= 100 ? v.toFixed(0) : v.toFixed(2));
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: W, height: H, fontFamily: FONT, background: "#fff", border: `1px solid ${PLAT.dark}` }}>
      {title && <text x={W / 2} y={15} textAnchor="middle" fontSize="12" fontWeight="bold">{title}</text>}
      {xt.map((t, i) => <g key={"x" + i}><line x1={sx(t)} y1={H - mb} x2={sx(t)} y2={H - mb + 4} stroke="#000" /><text x={sx(t)} y={H - mb + 15} textAnchor="middle" fontSize="9">{f3(t)}</text></g>)}
      {yt.map((t, i) => <g key={"y" + i}><line x1={ml} y1={sy(t)} x2={W - mr} y2={sy(t)} stroke="#eee" /><line x1={ml - 4} y1={sy(t)} x2={ml} y2={sy(t)} stroke="#000" /><text x={ml - 7} y={sy(t) + 3} textAnchor="end" fontSize="9">{f3(t)}</text></g>)}
      <line x1={ml} y1={mt} x2={ml} y2={H - mb} stroke="#000" /><line x1={ml} y1={H - mb} x2={W - mr} y2={H - mb} stroke="#000" />
      {hLine != null && <line x1={ml} y1={sy(hLine)} x2={W - mr} y2={sy(hLine)} stroke="#c00" strokeWidth="1.1" strokeDasharray="4 3" />}
      {points.map((p, i) => <circle key={i} cx={sx(p.x)} cy={sy(p.y)} r="2.6" fill="none" stroke="#1a3b6e" strokeWidth="1.1" />)}
      <text x={(ml + W - mr) / 2} y={H - 6} textAnchor="middle" fontSize="10">{xLabel}</text>
      <text x={13} y={(mt + H - mb) / 2} textAnchor="middle" fontSize="10" transform={`rotate(-90 13 ${(mt + H - mb) / 2})`}>{yLabel}</text>
    </svg>
  );
}

/* Residual diagnostics shared by regression & GLM: Shapiro–Wilk on residuals + Q–Q + residuals-vs-fitted. */
function ResidualDiagnostics({ resid, fitted }) {
  if (!resid || resid.length < 3) return null;
  const sw = shapiroWilk(resid), pts = fitted.map((f, i) => ({ x: f, y: resid[i] }));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <div style={{ fontWeight: "bold" }}>Residual diagnostics</div>
      {!sw.error && <div style={{ fontSize: 11 }}>Shapiro–Wilk on residuals: W = {fmt(sw.W, 4)}, P = {fmtP(sw.p)} {sw.p < 0.05 ? "— residuals may be non-normal" : "— consistent with normality"}</div>}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        <ExportFrame name="resid-qq"><QQPlot data={resid} title="Residuals: Normal Q–Q" /></ExportFrame>
        <ExportFrame name="resid-vs-fitted"><ScatterDiag points={pts} xLabel="Fitted values" yLabel="Residuals" title="Residuals vs Fitted" hLine={0} /></ExportFrame>
      </div>
      <div style={{ fontSize: 10, color: "#555" }}>Q–Q points near the line ⇒ normal residuals; residuals-vs-fitted should show no pattern or funnel (a funnel suggests non-constant variance).</div>
    </div>
  );
}
// Publication-style box / violin plot. groups: [{label, values:[number]}]. mode: "box"|"violin".
function BoxViolinChart({ groups, mode, opts, title, yLabel }) {
  const o = opts || {};
  const all = [].concat(...groups.map((g) => g.values));
  if (!all.length) return <svg viewBox="0 0 100 40" xmlns="http://www.w3.org/2000/svg" />;
  let lo = Math.min(...all), hi = Math.max(...all); if (lo === hi) { lo -= 1; hi += 1; }
  const pad = (hi - lo) * 0.06; const ns = niceScale(lo - pad, hi + pad, 6);
  const dmin = ns.lo, dmax = ns.hi, ticks = ns.ticks;
  const nG = groups.length, colW = 92, padL = 60, padR = 16, padT = title ? 28 : 14, padB = 46;
  const W = padL + nG * colW + padR, H = 326, plotH = H - padT - padB;
  const yOf = (v) => padT + plotH * (1 - (v - dmin) / (dmax - dmin));
  const cxOf = (i) => padL + colW * (i + 0.5);
  const els = [];
  els.push(<rect key="bg" x="0" y="0" width={W} height={H} fill="#fff" />);
  ticks.forEach((t, i) => {
    const y = yOf(t);
    els.push(<line key={"g" + i} x1={padL} y1={y} x2={W - padR} y2={y} stroke="#eee" strokeWidth="1" />);
    els.push(<text key={"yt" + i} x={padL - 6} y={y + 3} textAnchor="end" fontFamily={FONT} fontSize="10" fill="#555">{fmtTickVal(t)}</text>);
  });
  els.push(<line key="ay" x1={padL} y1={padT} x2={padL} y2={H - padB} stroke="#333" strokeWidth="1" />);
  els.push(<line key="ax" x1={padL} y1={H - padB} x2={W - padR} y2={H - padB} stroke="#333" strokeWidth="1" />);
  let gmax = 0, grid = [];
  if (mode === "violin") {
    const G = 64; for (let i = 0; i < G; i++) grid.push(dmin + (dmax - dmin) * i / (G - 1));
    groups.forEach((g) => { if (g.values.length > 1) { g._d = kdeDensity(g.values, grid, o.bw || 1); gmax = Math.max(gmax, ...g._d); } });
  }
  const jitter = (i, k) => { const h = Math.sin(i * 131.7 + k * 977.3) * 43758.5453; return (h - Math.floor(h)) - 0.5; };
  groups.forEach((g, i) => {
    const cx = cxOf(i), color = PALETTE[i % PALETTE.length], vals = g.values;
    if (!vals.length) return;
    const bs = boxStats(vals), halfBox = 22;
    if (mode === "violin" && g._d && gmax > 0) {
      const halfW = colW * 0.4, sc = halfW / gmax;
      let path = "";
      grid.forEach((x, gi) => { path += (gi ? "L" : "M") + (cx - g._d[gi] * sc).toFixed(2) + "," + yOf(x).toFixed(2) + " "; });
      for (let gi = grid.length - 1; gi >= 0; gi--) path += "L" + (cx + g._d[gi] * sc).toFixed(2) + "," + yOf(grid[gi]).toFixed(2) + " ";
      path += "Z";
      els.push(<path key={"v" + i} d={path} fill={color} fillOpacity="0.25" stroke={color} strokeWidth="1.4" />);
      if (o.showBox !== false) {
        els.push(<line key={"vb" + i} x1={cx} y1={yOf(bs.q1)} x2={cx} y2={yOf(bs.q3)} stroke="#333" strokeWidth="6" />);
        els.push(<line key={"vw" + i} x1={cx} y1={yOf(bs.wlo)} x2={cx} y2={yOf(bs.whi)} stroke="#333" strokeWidth="1.4" />);
        els.push(<circle key={"vm" + i} cx={cx} cy={yOf(bs.med)} r="2.6" fill="#fff" stroke="#333" strokeWidth="1.2" />);
      }
    } else if (mode === "box") {
      els.push(<line key={"wu" + i} x1={cx} y1={yOf(bs.whi)} x2={cx} y2={yOf(bs.q3)} stroke="#333" strokeWidth="1.2" />);
      els.push(<line key={"wd" + i} x1={cx} y1={yOf(bs.q1)} x2={cx} y2={yOf(bs.wlo)} stroke="#333" strokeWidth="1.2" />);
      els.push(<line key={"c1" + i} x1={cx - 10} y1={yOf(bs.whi)} x2={cx + 10} y2={yOf(bs.whi)} stroke="#333" strokeWidth="1.2" />);
      els.push(<line key={"c2" + i} x1={cx - 10} y1={yOf(bs.wlo)} x2={cx + 10} y2={yOf(bs.wlo)} stroke="#333" strokeWidth="1.2" />);
      els.push(<rect key={"bx" + i} x={cx - halfBox} y={yOf(bs.q3)} width={halfBox * 2} height={yOf(bs.q1) - yOf(bs.q3)} fill={color} fillOpacity="0.28" stroke={color} strokeWidth="1.5" />);
      els.push(<line key={"md" + i} x1={cx - halfBox} y1={yOf(bs.med)} x2={cx + halfBox} y2={yOf(bs.med)} stroke="#222" strokeWidth="2" />);
      if (o.showMean) els.push(symbolEl("diamond", cx, yOf(bs.mean), 4, "#b00", "mn" + i));
      bs.outliers.forEach((ov, k) => els.push(<circle key={"o" + i + "_" + k} cx={cx} cy={yOf(ov)} r="2" fill="none" stroke="#a33" strokeWidth="1" />));
    }
    if (o.showPoints) {
      const pj = mode === "box" ? halfBox * 0.7 : colW * 0.16;
      vals.forEach((v, k) => els.push(<circle key={"p" + i + "_" + k} cx={cx + jitter(i, k) * 2 * pj} cy={yOf(v)} r="1.6" fill="#333" fillOpacity="0.5" />));
    }
    els.push(<text key={"gl" + i} x={cx} y={H - padB + 16} textAnchor="middle" fontFamily={FONT} fontSize="11" fill="#222">{g.label}</text>);
  });
  if (yLabel) els.push(<text key="yl" x={15} y={padT + plotH / 2} textAnchor="middle" fontFamily={FONT} fontSize="11" fill="#333" transform={`rotate(-90 15 ${padT + plotH / 2})`}>{yLabel}</text>);
  if (title) els.push(<text key="ti" x={W / 2} y={17} textAnchor="middle" fontFamily={FONT} fontSize="12" fontWeight="bold" fill="#222">{title}</text>);
  return <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
}

/* Pie chart of category frequencies. slices: [{label, value}]. */
function PieChart({ slices, title, showPct }) {
  const total = slices.reduce((s, x) => s + x.value, 0) || 1;
  const W = 420, H = title ? 312 : 296, cx = 148, cy = title ? 160 : 150, r = 116;
  let ang = -Math.PI / 2;
  const arcs = slices.map((s, i) => {
    const frac = s.value / total, a0 = ang, a1 = ang + frac * 2 * Math.PI; ang = a1;
    const large = (a1 - a0) > Math.PI ? 1 : 0, color = s.color || PALETTE[i % PALETTE.length];
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0), x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const mid = (a0 + a1) / 2, lr = r * 0.6, lx = cx + lr * Math.cos(mid), ly = cy + lr * Math.sin(mid);
    const full = frac >= 0.9999;
    const path = full ? null : `M${cx},${cy} L${x0.toFixed(2)},${y0.toFixed(2)} A${r},${r} 0 ${large} 1 ${x1.toFixed(2)},${y1.toFixed(2)} Z`;
    return { path, color, frac, label: s.label, lx, ly, value: s.value, full };
  });
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">
      <rect x="0" y="0" width={W} height={H} fill="#fff" />
      {title && <text x={W / 2} y={17} textAnchor="middle" fontSize="12" fontWeight="bold">{title}</text>}
      {arcs.map((s, i) => s.full
        ? <circle key={i} cx={cx} cy={cy} r={r} fill={s.color} stroke="#fff" strokeWidth="1.5" />
        : <path key={i} d={s.path} fill={s.color} stroke="#fff" strokeWidth="1.5" />)}
      {arcs.map((s, i) => s.frac >= 0.05 ? <text key={"l" + i} x={s.lx} y={s.ly + 3} textAnchor="middle" fontSize="10" fill="#fff" fontWeight="bold">{(s.frac * 100).toFixed(showPct ? 1 : 0) + "%"}</text> : null)}
      {arcs.map((s, i) => (
        <g key={"lg" + i} transform={`translate(${W - 138}, ${(title ? 40 : 28) + i * 19})`}>
          <rect x="0" y="-9" width="12" height="12" fill={s.color} />
          <text x="18" y="1" fontSize="11" fill="#222">{s.label}</text>
          <text x="130" y="1" fontSize="10" fill="#555" textAnchor="end">{showPct ? (s.frac * 100).toFixed(1) + "%" : s.value}</text>
        </g>
      ))}
    </svg>
  );
}

/* Bar chart of category frequencies; grouped or stacked when a second (series) factor is present.
   data[c][s] = value for category c, series s. */
function CatBarChart({ categories, seriesNames, data, mode, title, xLabel, yLabel, pct }) {
  const nC = categories.length, nS = seriesNames.length, grouped = mode !== "stacked" || nS === 1;
  let vmax = 0;
  if (grouped) data.forEach((row) => row.forEach((v) => { if (v > vmax) vmax = v; }));
  else data.forEach((row) => { const sum = row.reduce((a, b) => a + b, 0); if (sum > vmax) vmax = sum; });
  if (vmax <= 0) vmax = 1;
  const ns = niceScale(0, vmax, 6), dmax = ns.hi, ticks = ns.ticks;
  const colW = Math.max(46, Math.min(110, 520 / nC)), padL = 56, padR = nS > 1 ? 134 : 16, padT = title ? 28 : 14, padB = 52;
  const W = padL + nC * colW + padR, H = 320, plotH = H - padT - padB;
  const yOf = (v) => padT + plotH * (1 - v / dmax);
  const cxOf = (i) => padL + colW * i;
  const els = [<rect key="bg" x="0" y="0" width={W} height={H} fill="#fff" />];
  ticks.forEach((t, i) => { const y = yOf(t); els.push(<line key={"g" + i} x1={padL} y1={y} x2={W - padR} y2={y} stroke="#eee" />); els.push(<text key={"yt" + i} x={padL - 6} y={y + 3} textAnchor="end" fontSize="10" fill="#555">{fmtTickVal(t)}</text>); });
  els.push(<line key="ay" x1={padL} y1={padT} x2={padL} y2={H - padB} stroke="#333" />);
  els.push(<line key="ax" x1={padL} y1={H - padB} x2={W - padR} y2={H - padB} stroke="#333" />);
  categories.forEach((cat, c) => {
    const x0 = cxOf(c), inner = colW * 0.8, gap = colW * 0.1;
    if (grouped) {
      const bw = inner / nS;
      data[c].forEach((v, s) => { const bx = x0 + gap + s * bw, by = yOf(v); els.push(<rect key={`b${c}_${s}`} x={bx + 1} y={by} width={Math.max(1, bw - 2)} height={H - padB - by} fill={PALETTE[s % PALETTE.length]} stroke="#0003" />); });
    } else {
      let acc = 0;
      data[c].forEach((v, s) => { const yTop = yOf(acc + v), yBot = yOf(acc); acc += v; els.push(<rect key={`b${c}_${s}`} x={x0 + gap} y={yTop} width={inner} height={Math.max(0, yBot - yTop)} fill={PALETTE[s % PALETTE.length]} stroke="#0003" />); });
    }
    els.push(<text key={"ct" + c} x={x0 + colW / 2} y={H - padB + 16} textAnchor="middle" fontSize="10" fill="#222">{cat}</text>);
  });
  if (nS > 1) seriesNames.forEach((nm, s) => (els.push(
    <g key={"lg" + s} transform={`translate(${W - padR + 8}, ${padT + 4 + s * 18})`}>
      <rect x="0" y="-9" width="12" height="12" fill={PALETTE[s % PALETTE.length]} /><text x="17" y="1" fontSize="11" fill="#222">{nm}</text>
    </g>)));
  if (yLabel) els.push(<text key="yl" x={14} y={padT + plotH / 2} textAnchor="middle" fontSize="11" fill="#333" transform={`rotate(-90 14 ${padT + plotH / 2})`}>{yLabel}</text>);
  if (xLabel) els.push(<text key="xl" x={padL + (nC * colW) / 2} y={H - 8} textAnchor="middle" fontSize="11" fill="#333">{xLabel}</text>);
  if (title) els.push(<text key="ti" x={W / 2} y={17} textAnchor="middle" fontSize="12" fontWeight="bold" fill="#222">{title}</text>);
  return <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
}

/* Forest plot of point estimates with confidence intervals. rows: [{label, est, lo, hi}]. */
function ForestPlot({ rows, title, xLabel, nullValue }) {
  const nv = nullValue == null ? 0 : nullValue;
  let lo = Math.min(nv, ...rows.map((r) => r.lo)), hi = Math.max(nv, ...rows.map((r) => r.hi));
  if (lo === hi) { lo -= 1; hi += 1; }
  const pad = (hi - lo) * 0.08; const ns = niceScale(lo - pad, hi + pad, 6), X0 = ns.lo, X1 = ns.hi, ticks = ns.ticks;
  const padL = 150, padR = 96, padT = title ? 30 : 16, rowH = 30, padB = 40;
  const W = 560, H = padT + rows.length * rowH + padB;
  const sx = (v) => padL + (v - X0) / (X1 - X0) * (W - padL - padR);
  const yOf = (i) => padT + rowH * (i + 0.5);
  const els = [<rect key="bg" x="0" y="0" width={W} height={H} fill="#fff" />];
  ticks.forEach((t, i) => { els.push(<line key={"g" + i} x1={sx(t)} y1={padT} x2={sx(t)} y2={H - padB} stroke="#eee" />); els.push(<text key={"xt" + i} x={sx(t)} y={H - padB + 15} textAnchor="middle" fontSize="9" fill="#555">{fmtTickVal(t)}</text>); });
  els.push(<line key="nv" x1={sx(nv)} y1={padT} x2={sx(nv)} y2={H - padB} stroke="#c00" strokeWidth="1.1" strokeDasharray="4 3" />);
  rows.forEach((r, i) => {
    const y = yOf(i), col = PALETTE[i % PALETTE.length];
    els.push(<text key={"rl" + i} x={padL - 10} y={y + 3} textAnchor="end" fontSize="11" fill="#222">{r.label}</text>);
    els.push(<line key={"ci" + i} x1={sx(r.lo)} y1={y} x2={sx(r.hi)} y2={y} stroke={col} strokeWidth="1.6" />);
    els.push(<line key={"cl" + i} x1={sx(r.lo)} y1={y - 4} x2={sx(r.lo)} y2={y + 4} stroke={col} strokeWidth="1.4" />);
    els.push(<line key={"cr" + i} x1={sx(r.hi)} y1={y - 4} x2={sx(r.hi)} y2={y + 4} stroke={col} strokeWidth="1.4" />);
    els.push(<rect key={"pt" + i} x={sx(r.est) - 4} y={y - 4} width="8" height="8" fill={col} transform={`rotate(45 ${sx(r.est)} ${y})`} />);
    els.push(<text key={"vt" + i} x={W - padR + 8} y={y + 3} fontSize="10" fill="#333">{fmt(r.est, 2)} [{fmt(r.lo, 2)}, {fmt(r.hi, 2)}]</text>);
  });
  els.push(<line key="ax" x1={padL} y1={padT} x2={padL} y2={H - padB} stroke="#333" />);
  if (xLabel) els.push(<text key="xl" x={padL + (W - padL - padR) / 2} y={H - 6} textAnchor="middle" fontSize="11" fill="#333">{xLabel}</text>);
  if (title) els.push(<text key="ti" x={W / 2} y={18} textAnchor="middle" fontSize="12" fontWeight="bold" fill="#222">{title}</text>);
  return <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
}

// Diverging color for a correlation value: blue (−1) — white (0) — red (+1).
function rColor(r) {
  const t = Math.max(-1, Math.min(1, r || 0)), w = [247, 247, 247], pos = [178, 24, 43], neg = [33, 102, 172], e = t >= 0 ? pos : neg, b = Math.abs(t);
  const L = (i) => Math.round(w[i] + (e[i] - w[i]) * b);
  return `rgb(${L(0)},${L(1)},${L(2)})`;
}
/* Correlation matrix visual: heatmap, scatterplot matrix (SPLOM), or combined (scatters below diagonal, r-cells above). */
function PairsPlot({ names, cols, mode, method }) {
  const n = names.length, cell = n <= 4 ? 96 : n <= 6 ? 78 : 64, pad = 78, W = pad + n * cell + 8, H = pad + n * cell + 8;
  const complete = (i, j) => { const xs = [], ys = []; const N = cols[i].length; for (let k = 0; k < N; k++) { const a = cols[i][k], b = cols[j][k]; if (a != null && b != null) { xs.push(a); ys.push(b); } } return [xs, ys]; };
  const rOf = (i, j) => { const [xs, ys] = complete(i, j); if (xs.length < 3) return NaN; return method === "spearman" ? spearman(xs, ys, "two").rho : pearsonPair(xs, ys).r; };
  const rng = cols.map((c) => { const v = c.filter((x) => x != null); return v.length ? [Math.min(...v), Math.max(...v)] : [0, 1]; });
  const els = [<rect key="bg" x="0" y="0" width={W} height={H} fill="#fff" />];
  names.forEach((nm, i) => {
    els.push(<text key={"ct" + i} x={pad + cell * (i + 0.5)} y={pad - 8} textAnchor="middle" fontSize="10" fontWeight="bold" fill="#222">{nm.length > 10 ? nm.slice(0, 9) + "…" : nm}</text>);
    els.push(<text key={"rt" + i} x={pad - 8} y={pad + cell * (i + 0.5) + 3} textAnchor="end" fontSize="10" fontWeight="bold" fill="#222">{nm.length > 10 ? nm.slice(0, 9) + "…" : nm}</text>);
  });
  for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) {
    const x0 = pad + j * cell, y0 = pad + i * cell;
    const heatCell = mode === "heat" || (mode === "both" && j > i), scatterCell = mode === "scatter" || (mode === "both" && j < i), diag = i === j;
    els.push(<rect key={`f${i}_${j}`} x={x0} y={y0} width={cell} height={cell} fill="none" stroke="#ccc" />);
    if (diag && mode !== "heat") { els.push(<text key={`d${i}`} x={x0 + cell / 2} y={y0 + cell / 2 + 3} textAnchor="middle" fontSize="10" fontWeight="bold" fill="#444">{names[i].length > 9 ? names[i].slice(0, 8) + "…" : names[i]}</text>); continue; }
    if (heatCell || (diag && mode === "heat")) {
      const r = diag ? 1 : rOf(i, j);
      els.push(<rect key={`h${i}_${j}`} x={x0 + 1} y={y0 + 1} width={cell - 2} height={cell - 2} fill={rColor(r)} />);
      if (isFinite(r)) els.push(<text key={`ht${i}_${j}`} x={x0 + cell / 2} y={y0 + cell / 2 + 4} textAnchor="middle" fontSize={n <= 6 ? 12 : 10} fill={Math.abs(r) > 0.5 ? "#fff" : "#222"} fontWeight="bold">{diag ? "1" : r.toFixed(2)}</text>);
    } else if (scatterCell) {
      const [xs, ys] = complete(i, j); // y = var i, x = var j
      const [xlo, xhi] = rng[j], [ylo, yhi] = rng[i], sx = (v) => x0 + 5 + (xhi > xlo ? (v - xlo) / (xhi - xlo) : 0.5) * (cell - 10), sy = (v) => y0 + cell - 5 - (yhi > ylo ? (v - ylo) / (yhi - ylo) : 0.5) * (cell - 10);
      for (let k = 0; k < xs.length; k++) els.push(<circle key={`p${i}_${j}_${k}`} cx={sx(xs[k])} cy={sy(ys[k])} r="1.5" fill="#1a3b6e" fillOpacity="0.5" />);
    }
  }
  return <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
}

/* Repeated-measures profile ("spaghetti") plot: per-subject trajectories across a within factor, with group/grand mean overlays. */
function ProfilePlot({ levels, subjects, means, xLabel, yLabel, title, showSubjects }) {
  const nL = levels.length, padL = 58, padR = means.length > 1 ? 130 : 16, padT = title ? 28 : 14, padB = 46;
  const W = padL + Math.max(220, nL * 80) + padR, H = 330, plotW = W - padL - padR, plotH = H - padT - padB;
  const allY = [];
  subjects.forEach((s) => s.ys.forEach((v) => { if (v != null) allY.push(v); }));
  means.forEach((m) => m.ys.forEach((c) => { if (c.mean != null) { allY.push(c.mean + (c.err || 0)); allY.push(c.mean - (c.err || 0)); } }));
  let lo = Math.min(...allY), hi = Math.max(...allY); if (lo === hi) { lo -= 1; hi += 1; }
  const ns = niceScale(lo, hi, 6), dmin = ns.lo, dmax = ns.hi, ticks = ns.ticks;
  const xOf = (i) => padL + (nL === 1 ? plotW / 2 : plotW * i / (nL - 1));
  const yOf = (v) => padT + plotH * (1 - (v - dmin) / (dmax - dmin));
  const els = [<rect key="bg" x="0" y="0" width={W} height={H} fill="#fff" />];
  ticks.forEach((t, i) => { const y = yOf(t); els.push(<line key={"g" + i} x1={padL} y1={y} x2={W - padR} y2={y} stroke="#eee" />); els.push(<text key={"yt" + i} x={padL - 6} y={y + 3} textAnchor="end" fontSize="10" fill="#555">{fmtTickVal(t)}</text>); });
  els.push(<line key="ay" x1={padL} y1={padT} x2={padL} y2={H - padB} stroke="#333" />);
  els.push(<line key="ax" x1={padL} y1={H - padB} x2={W - padR} y2={H - padB} stroke="#333" />);
  levels.forEach((L, i) => els.push(<text key={"xl" + i} x={xOf(i)} y={H - padB + 16} textAnchor="middle" fontSize="10" fill="#222">{String(L)}</text>));
  const groupColor = (gi) => PALETTE[gi % PALETTE.length];
  if (showSubjects) subjects.forEach((s, si) => {
    const gi = Math.max(0, means.findIndex((m) => m.group === s.group)), col = means.length > 1 ? groupColor(gi) : "#888";
    let d = "", started = false;
    s.ys.forEach((v, i) => { if (v == null) { started = false; return; } d += (started ? "L" : "M") + xOf(i).toFixed(1) + "," + yOf(v).toFixed(1) + " "; started = true; });
    if (d) els.push(<path key={"s" + si} d={d} fill="none" stroke={col} strokeWidth="0.8" strokeOpacity="0.32" />);
  });
  means.forEach((m, gi) => {
    const col = means.length > 1 ? groupColor(gi) : "#b00";
    let d = "", started = false;
    m.ys.forEach((c, i) => { if (c.mean == null) { started = false; return; } const x = xOf(i), y = yOf(c.mean); d += (started ? "L" : "M") + x.toFixed(1) + "," + y.toFixed(1) + " "; started = true; if (c.err) { els.push(<line key={`e${gi}_${i}`} x1={x} y1={yOf(c.mean - c.err)} x2={x} y2={yOf(c.mean + c.err)} stroke={col} strokeWidth="1.3" />); els.push(<line key={`ec1${gi}_${i}`} x1={x - 4} y1={yOf(c.mean + c.err)} x2={x + 4} y2={yOf(c.mean + c.err)} stroke={col} strokeWidth="1.3" />); els.push(<line key={`ec2${gi}_${i}`} x1={x - 4} y1={yOf(c.mean - c.err)} x2={x + 4} y2={yOf(c.mean - c.err)} stroke={col} strokeWidth="1.3" />); } els.push(<circle key={`m${gi}_${i}`} cx={x} cy={y} r="3" fill={col} />); });
    if (d) els.push(<path key={"ml" + gi} d={d} fill="none" stroke={col} strokeWidth="2.4" />);
  });
  if (means.length > 1) means.forEach((m, gi) => (els.push(
    <g key={"lg" + gi} transform={`translate(${W - padR + 8}, ${padT + 4 + gi * 18})`}>
      <line x1="0" y1="-3" x2="16" y2="-3" stroke={groupColor(gi)} strokeWidth="2.4" /><text x="22" y="1" fontSize="11" fill="#222">{m.group}</text>
    </g>)));
  if (yLabel) els.push(<text key="ylb" x={14} y={padT + plotH / 2} textAnchor="middle" fontSize="11" fill="#333" transform={`rotate(-90 14 ${padT + plotH / 2})`}>{yLabel}</text>);
  if (xLabel) els.push(<text key="xlb" x={padL + plotW / 2} y={H - 6} textAnchor="middle" fontSize="11" fill="#333">{xLabel}</text>);
  if (title) els.push(<text key="ti" x={W / 2} y={17} textAnchor="middle" fontSize="12" fontWeight="bold" fill="#222">{title}</text>);
  return <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: W, fontFamily: FONT }} xmlns="http://www.w3.org/2000/svg">{els}</svg>;
}

function renderAnovaGraph(a, colById, compactById, rows, kind, onPlotResize, onPlotLegend) {
  const m = anovaModel(a, colById, compactById, rows);
  if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
  const cfg = { ...DEFAULT_ANOVA_CFG, ...a.cfg };
  const effects = effectList(m.factors, cfg.effects);
  if (effects.length === 0) return <div style={{ color: "#999" }}>No factors to plot.</div>;
  const showErr = cfg.errorBars !== "none";
  const plot = { ...DEFAULT_PLOT, ...(a.plot || {}) };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <ConfigLine a={a} depName={m.depName} />
      {effects.map((E, ei) => {
        const rws = cellMeans(m.long, E, cfg.alpha, cfg.errorBars);
        const gm = graphModel(E, rws, m.levelOrder, a.swap);
        return (
          <div key={ei}>
            <div style={{ fontWeight: "bold", marginBottom: 2 }}>{E.join(" × ")}{gm.seriesFac ? ` — x: ${gm.xFac}, series: ${gm.seriesFac}` : ` — x: ${gm.xFac}`}</div>
            {gm.facets.map((facet, fi) => (
              <div key={fi} style={{ marginBottom: 8 }}>
                {facet.label && <div style={{ fontSize: 10, color: "#456" }}>{facet.label}</div>}
                <ChartWithExport name={`${m.depName}_${E.join("-")}${facet.label ? "_" + facet.label : ""}`} data={facet.data} seriesKeys={facet.seriesKeys} kind={kind} showErr={showErr} plot={plot} depName={m.depName} hasSeries={!!gm.seriesFac} onCommitSize={onPlotResize ? (w, h) => onPlotResize(a.id, w, h) : undefined} onCommitLegend={onPlotLegend ? (xy) => onPlotLegend(a.id, xy) : undefined} />
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
}

function renderAnovaPosthoc(a, colById, compactById, rows) {
  const m = anovaModel(a, colById, compactById, rows);
  if (m.error) return <div style={{ color: "#999", fontStyle: "italic" }}>{m.error}</div>;
  const cfg = { ...DEFAULT_ANOVA_CFG, ...a.cfg };
  const method = a.method || "tukey";
  const effect = (a.phEffect ? a.phEffect.split("\u0001") : [m.factors[m.factors.length - 1]]).filter((f) => m.factors.includes(f));
  if (effect.length === 0) return <div style={{ color: "#a00" }}>Choose an effect for post-hoc comparisons.</div>;
  const nCells = effect.reduce((n, f) => n * (m.levelOrder[f] || []).length, 1);
  if (nCells < 2) return <div style={{ color: "#a00" }}>This effect has fewer than 2 cells.</div>;
  let res; try { res = posthoc(m.long, effect, m.withinNames, m.betweenNames, m.levelOrder, method); } catch (e) { return <div style={{ color: "#a00" }}>Could not compute: {String(e.message || e)}</div>; }
  const cellLabel = (arr) => effect.length === 1 ? `${effect[0]} ${arr[0]}` : `(${arr.join(", ")})`;
  const showQ = method === "tukey" && res.family === "pooled";
  const sig = (p) => p != null && p < cfg.alpha;
  const kind = res.mixed ? "mixed within×between" : (m.withinNames.includes(effect[0]) && res.family === "pooled" && !res.mixed && effect.every((f) => m.withinNames.includes(f)) ? "within" : "between");
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <div style={{ fontSize: 10, color: "#456" }}>{PH_METHODS[method]} · effect <b>{effect.join(" × ")}</b> ({kind}, {res.K} cells) · {res.m} pairwise comparisons{res.family === "pooled" ? ` · error df ${res.commonDf}` : " · per-comparison error"} · α = {cfg.alpha}</div>
      {res.fwerNA && <div style={{ color: "#a40", fontSize: 10 }}>⚠ {PH_METHODS[method]} needs a single homogeneous error term and isn't defined for mixed within×between effects. Showing uncorrected p-values — use Bonferroni, Holm, or Benjamini–Hochberg here.</div>}
      {(() => {
        const phHead = ["Comparison", "Mean diff.", "Std. err", showQ ? "q" : "t", "df", (method === "lsd" || res.fwerNA) ? "P" : "P (adj.)"];
        const phRows = res.pairs.map((p) => [`${cellLabel(p.a)} − ${cellLabel(p.b)}`, fmt(p.diff, 2), fmt(p.se, 3), fmt(showQ ? Math.SQRT2 * Math.abs(p.t) : p.t, 3), p.df, fmtP(p.padj)]);
        return (
      <ExportWrap head={phHead} rows={phRows} name="posthoc">
      <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
        <thead><tr>{["Comparison", "Mean diff.", "Std. err", showQ ? "q" : "t", "df", (method === "lsd" || res.fwerNA) ? "P" : "P (adj.)", ""].map((h, i) => (
          <th key={i} style={{ border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", textAlign: i === 0 ? "left" : "right", fontWeight: "bold" }}>{h}</th>
        ))}</tr></thead>
        <tbody>
          {res.pairs.map((p, i) => {
            const stat = showQ ? Math.SQRT2 * Math.abs(p.t) : p.t;
            const s = sig(p.padj);
            return (
              <tr key={i} style={{ background: s ? "#eef7ee" : "#fff" }}>
                <td style={{ ...anvCell, textAlign: "left" }}>{cellLabel(p.a)} − {cellLabel(p.b)}</td>
                <td style={anvCell}>{fmt(p.diff, 2)}</td>
                <td style={anvCell}>{fmt(p.se, 3)}</td>
                <td style={anvCell}>{fmt(stat, 3)}</td>
                <td style={anvCell}>{p.df}</td>
                <td style={{ ...anvCell, fontWeight: s ? "bold" : "normal", color: s ? "#070" : "#000" }}>{fmtP(p.padj)}</td>
                <td style={{ ...anvCell, color: "#070", fontWeight: "bold" }}>{s ? "✓" : ""}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      </ExportWrap>
        );
      })()}
      <div style={{ fontSize: 10, color: "#555" }}>
        {res.family === "pooled"
          ? (kind === "within"
            ? "All-within effect: comparisons use the (effect × Subject) within-groups error term, matching the ANOVA stratum (assumes sphericity and complete within-cell data)."
            : "All-between effect: pooled within-cell error with per-pair standard errors √(MSE·(1/nᵢ+1/nⱼ)) — Tukey–Kramer, valid for unequal group sizes.")
          : "Mixed effect: each pair is tested directly — a paired t when the two cells share the same between level(s), a two-sample t otherwise — so the error df vary by comparison."}
        {" "}✓ = significant at α = {cfg.alpha}.
      </div>
    </div>
  );
}

/* ---- Formula builder popup (visual cell picker) ---- */
function FormulaBuilderDialog({ col, numericCols, compacts, onApply, onCancel }) {
  const [text, setText] = useState(col && col.formula ? col.formula : "");
  const [sel, setSel] = useState(() => new Set());
  const [fn, setFn] = useState("AVERAGE");
  const anchor = useRef(null);
  const dragging = useRef(false);
  useEffect(() => { const up = () => { dragging.current = false; }; window.addEventListener("mouseup", up); return () => window.removeEventListener("mouseup", up); }, []);
  const span = (a, b) => { const lo = Math.min(a, b), hi = Math.max(a, b); return new Set(numericCols.slice(lo, hi + 1).map((c) => c.id)); };
  const onDown = (idx, e) => {
    e.preventDefault(); const id = numericCols[idx].id;
    if (e.shiftKey && anchor.current != null) setSel(span(anchor.current, idx));
    else if (e.altKey || e.metaKey || e.ctrlKey) { setSel((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; }); anchor.current = idx; }
    else { setSel(new Set([id])); anchor.current = idx; dragging.current = true; }
  };
  const onEnter = (idx) => { if (dragging.current && anchor.current != null) setSel(span(anchor.current, idx)); };
  const orderedSel = numericCols.filter((c) => sel.has(c.id));
  const buildExpr = () => {
    if (orderedSel.length === 0) return "";
    const ids = orderedSel.map((c) => c.id);
    for (const cp of compacts) { if (cp.leaves.length === ids.length && cp.leaves.every((id) => sel.has(id))) return fn === "(refs)" ? `[${cp.name}]` : `${fn}([${cp.name}])`; }
    const refs = orderedSel.map((c) => `[${c.label}]`).join(", ");
    return fn === "(refs)" ? refs : `${fn}(${refs})`;
  };
  const expr = buildExpr();
  const insert = () => { if (!expr) return; setText((t) => (t && t.trim() ? t + " " + expr : expr)); setSel(new Set()); anchor.current = null; };
  let lastGroup;
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 3000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ width: 470, maxHeight: 580, background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11, display: "flex", flexDirection: "column" }}>
        <div style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>Formula Builder{col ? " — " + col.name : ""}</span>
        </div>
        <div style={{ padding: 10, display: "flex", flexDirection: "column", gap: 8, minHeight: 0, flex: 1 }}>
          <div>
            <div style={{ fontWeight: "bold", marginBottom: 2 }}>Formula</div>
            <input value={text} onChange={(e) => setText(e.target.value)} placeholder="=AVERAGE([RT (ms)])" style={{ width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 11, ...bevelIn, padding: "2px 4px", background: "#fffef6" }} />
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
            <span style={{ fontWeight: "bold" }}>Wrap selection in:</span>
            <select value={fn} onChange={(e) => setFn(e.target.value)} style={{ fontFamily: FONT, fontSize: 11 }}>
              {["AVERAGE", "SUM", "MIN", "MAX", "COUNT", "MEDIAN", "STDEV", "VAR", "PRODUCT", "RANGE"].map((f) => <option key={f} value={f}>{f}</option>)}
              <option value="(refs)">(no function — refs only)</option>
            </select>
            <MacBtn onClick={insert} disabled={!expr}>Insert ▸</MacBtn>
            <span style={{ color: "#456" }}>{orderedSel.length} selected</span>
          </div>
          {expr && <div style={{ ...bevelIn, background: "#f4f4f4", padding: "2px 4px", fontFamily: "monospace", fontSize: 10, color: "#225", whiteSpace: "nowrap", overflow: "auto" }}>{expr}</div>}
          <div style={{ fontSize: 10, color: "#556" }}>click to select · shift-click for a span · ⌥/⌘-click to toggle · click-drag to sweep</div>
          <div style={{ flex: 1, minHeight: 90, overflowY: "auto", ...bevelIn, background: "#fff" }}>
            {numericCols.length === 0 && <div style={{ padding: 8, color: "#999" }}>No numeric columns in the dataset.</div>}
            {numericCols.map((c, idx) => {
              const head = c.group !== lastGroup ? c.group : null; lastGroup = c.group;
              const on = sel.has(c.id);
              return (
                <div key={c.id}>
                  {head && <div style={{ background: "#f3e7dc", color: COMPACT_CLR, fontWeight: "bold", padding: "1px 6px", fontSize: 10 }}>▣ {head}</div>}
                  <div onMouseDown={(e) => onDown(idx, e)} onMouseEnter={() => onEnter(idx)}
                    style={{ padding: c.group ? "1px 6px 1px 18px" : "1px 6px", cursor: "pointer", background: on ? PLAT.sel : "transparent", userSelect: "none" }}>
                    {on ? "☑" : "☐"} {c.label} <span style={{ color: "#aab", fontSize: 9 }}>[{c.id}]</span>
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
            <MacBtn onClick={() => { setText(""); setSel(new Set()); anchor.current = null; }}>Clear</MacBtn>
            <span style={{ display: "flex", gap: 8 }}>
              <MacBtn onClick={onCancel}>Cancel</MacBtn>
              <MacBtn onClick={() => onApply(text)} primary>OK</MacBtn>
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---- Plot settings popup (editable graph elements) ---- */
// Lets a centered modal be dragged by its title bar (offset from center via transform).
function useModalDrag() {
  const [d, setD] = useState({ x: 0, y: 0 });
  const onHeaderDown = (e) => {
    if (e.button !== 0) return;
    if (e.target.closest && e.target.closest("input,select,button,textarea,a")) return;
    e.preventDefault();
    const sx = e.clientX, sy = e.clientY, ox = d.x, oy = d.y;
    const move = (ev) => setD({ x: ox + ev.clientX - sx, y: oy + ev.clientY - sy });
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
  };
  return { dragStyle: { transform: `translate(${d.x}px, ${d.y}px)` }, onHeaderDown };
}
function PlotSettingsDialog({ dlg, onApply, onCancel }) {
  const [p, setP] = useState(() => ({ ...DEFAULT_PLOT, ...dlg.plot, series: { ...(dlg.plot.series || {}) } }));
  const { dragStyle, onHeaderDown } = useModalDrag();
  const set = (k, v) => setP((s) => ({ ...s, [k]: v }));
  const setSeries = (key, field, v) => setP((s) => ({ ...s, series: { ...s.series, [key]: { ...seriesStyleOf(s, key, dlg.seriesKeys.indexOf(key)), [field]: v } } }));
  const num = { width: 56, boxSizing: "border-box", fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 3px" };
  const txt = { fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 3px" };
  const lbl = { display: "flex", alignItems: "center", gap: 5, margin: "2px 0" };
  const sec = { fontWeight: "bold", margin: "8px 0 3px", borderBottom: `1px solid ${PLAT.dark}`, paddingBottom: 1 };
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 3000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ width: 440, maxHeight: 600, background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11, display: "flex", flexDirection: "column", ...dragStyle }}>
        <div onMouseDown={onHeaderDown} title="drag to move" style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center", cursor: "move" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>Plot Settings</span>
        </div>
        <div style={{ padding: 10, overflowY: "auto" }}>
          <div style={sec}>Titles &amp; labels</div>
          <label style={lbl}>Title <input style={{ ...txt, flex: 1 }} value={p.title} onChange={(e) => set("title", e.target.value)} /></label>
          <label style={lbl}>X-axis label <input style={{ ...txt, flex: 1 }} value={p.xLabel} onChange={(e) => set("xLabel", e.target.value)} /></label>
          <label style={lbl}>Y-axis label <input style={{ ...txt, flex: 1 }} value={p.yLabel} onChange={(e) => set("yLabel", e.target.value)} /></label>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <label style={lbl}>Label size <input style={num} type="number" min="6" max="28" value={p.labelSize} onChange={(e) => set("labelSize", e.target.value)} /></label>
            <label style={lbl}><input type="checkbox" checked={!!p.labelBold} onChange={(e) => set("labelBold", e.target.checked)} /> bold</label>
            <label style={lbl}><input type="checkbox" checked={!!p.labelItalic} onChange={(e) => set("labelItalic", e.target.checked)} /> italic</label>
            <label style={lbl}>Tick size <input style={num} type="number" min="6" max="24" value={p.tickSize} onChange={(e) => set("tickSize", e.target.value)} /></label>
          </div>

          <div style={sec}>Axes &amp; range</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <label style={lbl}>Y min <input style={num} value={p.yMin} onChange={(e) => set("yMin", e.target.value)} placeholder="auto" /></label>
            <label style={lbl}>Y max <input style={num} value={p.yMax} onChange={(e) => set("yMax", e.target.value)} placeholder="auto" /></label>
            <label style={lbl}>Y tick step <input style={num} value={p.yTickInterval} onChange={(e) => set("yTickInterval", e.target.value)} placeholder="auto" /></label>
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <label style={lbl}>X label skip <input style={num} type="number" min="0" value={p.xInterval} onChange={(e) => set("xInterval", e.target.value)} /></label>
            <label style={lbl}>X labels
              <select value={p.xAngle || 0} onChange={(e) => set("xAngle", +e.target.value)} style={{ fontFamily: FONT, fontSize: 11 }}>
                <option value={0}>horizontal</option>
                <option value={45}>45°</option>
                <option value={90}>vertical</option>
              </select>
            </label>
            <label style={lbl}>Axis line w <input style={num} type="number" min="0" step="0.5" value={p.axisWidth} onChange={(e) => set("axisWidth", e.target.value)} /></label>
            <label style={lbl}>Tick mark w <input style={num} type="number" min="0" step="0.5" value={p.tickWidth} onChange={(e) => set("tickWidth", e.target.value)} /></label>
          </div>

          <div style={sec}>Size</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <label style={lbl}>Width <input style={num} value={p.width || ""} onChange={(e) => set("width", e.target.value)} placeholder="auto" /></label>
            <label style={lbl}>Height <input style={num} value={p.height || ""} onChange={(e) => set("height", e.target.value)} placeholder="auto" /></label>
            <span style={{ fontSize: 10, color: "#667" }}>or drag the ◢ corner of a graph</span>
          </div>

          <div style={sec}>Grid &amp; frame</div>
          <label style={lbl}><input type="checkbox" checked={p.grid} onChange={(e) => set("grid", e.target.checked)} /> Show grid lines</label>
          {p.grid && <div style={{ paddingLeft: 16, display: "flex", gap: 12, flexWrap: "wrap" }}>
            <label style={lbl}><input type="checkbox" checked={p.gridH} onChange={(e) => set("gridH", e.target.checked)} /> horizontal</label>
            <label style={lbl}><input type="checkbox" checked={p.gridV} onChange={(e) => set("gridV", e.target.checked)} /> vertical</label>
            <label style={lbl}><input type="checkbox" checked={p.gridDashed} onChange={(e) => set("gridDashed", e.target.checked)} /> dashed</label>
          </div>}
          <label style={lbl}><input type="checkbox" checked={p.frame} onChange={(e) => set("frame", e.target.checked)} /> Frame (box around plot)</label>
          <label style={lbl}>Legend
            <select value={p.legend || "top"} onChange={(e) => setP((s) => ({ ...s, legend: e.target.value, legendXY: null }))} style={{ fontFamily: FONT, fontSize: 11 }}>
              <option value="top">top</option>
              <option value="tl">inside top-left</option>
              <option value="tr">inside top-right</option>
              <option value="bl">inside bottom-left</option>
              <option value="br">inside bottom-right</option>
              <option value="right">outside right</option>
              <option value="none">none</option>
            </select>
          </label>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", paddingLeft: 16 }}>
            <label style={lbl}>Legend size <input style={num} type="number" min="6" max="24" value={p.legendSize} onChange={(e) => set("legendSize", e.target.value)} /></label>
            <label style={lbl}><input type="checkbox" checked={!!p.legendBold} onChange={(e) => set("legendBold", e.target.checked)} /> bold</label>
            <label style={lbl}><input type="checkbox" checked={!!p.legendItalic} onChange={(e) => set("legendItalic", e.target.checked)} /> italic</label>
          </div>
          <div style={{ paddingLeft: 16, fontSize: 10, color: "#667", display: "flex", alignItems: "center", gap: 8, marginTop: 2 }}>
            {p.legendXY ? (<><span>Legend is custom-positioned.</span><button onClick={() => set("legendXY", null)} style={{ fontFamily: FONT, fontSize: 10, ...bevelOut, background: PLAT.faceLite, cursor: "pointer", padding: "0 7px" }}>reset position</button></>) : (<span>Tip: drag the legend on the graph to position it freely.</span>)}
          </div>

          <div style={sec}>Series — color, symbol, connector</div>
          {dlg.seriesKeys.length === 0 && <div style={{ color: "#999" }}>No series detected yet (open the graph once, then reopen settings).</div>}
          {dlg.seriesKeys.map((k, i) => { const st = seriesStyleOf(p, k, i); return (
            <div key={k} style={{ display: "flex", alignItems: "center", gap: 6, margin: "4px 0", flexWrap: "wrap" }}>
              <span style={{ minWidth: 64, fontWeight: "bold" }}>{k === "mean" ? "(value)" : k}</span>
              <span style={{ display: "flex", alignItems: "center", gap: 3, padding: "1px 4px", ...bevelIn, background: "#f4f1ea" }}>
                <span style={{ color: "#456" }}>symbol</span>
                <select title="symbol shape" value={st.symbol} onChange={(e) => setSeries(k, "symbol", e.target.value)} style={{ fontFamily: FONT, fontSize: 10 }}>{SYMBOLS.map((s) => <option key={s} value={s}>{s}</option>)}</select>
                <input title="symbol size (px)" style={{ ...num, width: 38 }} type="number" min="0" step="0.5" value={st.symbolSize} onChange={(e) => setSeries(k, "symbolSize", e.target.value)} />
                <input title="symbol color" type="color" value={st.symbolColor} onChange={(e) => setSeries(k, "symbolColor", e.target.value)} />
              </span>
              <span style={{ display: "flex", alignItems: "center", gap: 3, padding: "1px 4px", ...bevelIn, background: "#f4f1ea" }}>
                <span style={{ color: "#456" }}>connector</span>
                <input title="connector width (px)" style={{ ...num, width: 38 }} type="number" min="0" step="0.5" value={st.lineWidth} onChange={(e) => setSeries(k, "lineWidth", e.target.value)} />
                <input title="connector color" type="color" value={st.color} onChange={(e) => setSeries(k, "color", e.target.value)} />
              </span>
            </div>
          ); })}
          <div style={{ fontSize: 10, color: "#667", marginTop: 4 }}>For bar graphs the connector color is used as the bar fill; symbol settings are ignored. Set symbol to “none” for lines without points, or connector width to 0 for points only.</div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, padding: 8, borderTop: `1px solid ${PLAT.dark}` }}>
          <MacBtn onClick={() => setP({ ...DEFAULT_PLOT, series: {} })}>Reset</MacBtn>
          <MacBtn onClick={onCancel}>Cancel</MacBtn>
          <MacBtn onClick={() => onApply(p)} primary>OK</MacBtn>
        </div>
      </div>
    </div>
  );
}

/* ---- ANOVA setup popup ---- */
function AboutGraphic() {
  return (
    <svg width="150" height="120" viewBox="0 0 150 120" xmlns="http://www.w3.org/2000/svg" style={{ display: "block", ...bevelOut, borderRadius: 13 }}>
      <defs>
        <linearGradient id="abSky" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#d3ecf9" /><stop offset="1" stopColor="#8ec8ea" /></linearGradient>
        <linearGradient id="abBoard" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#eb9a4d" /><stop offset="1" stopColor="#c0653a" /></linearGradient>
        <linearGradient id="abSun" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#ffd27a" /><stop offset="0.55" stopColor="#ff8a3d" /><stop offset="1" stopColor="#ec5526" /></linearGradient>
        <clipPath id="abClip"><rect x="0" y="0" width="150" height="120" rx="13" /></clipPath>
      </defs>
      <g clipPath="url(#abClip)">
        <rect x="0" y="0" width="150" height="120" fill="url(#abSky)" />
        <circle cx="123" cy="26" r="19" fill="#ffffff" opacity="0.30" />
        <circle cx="123" cy="26" r="12" fill="#fff3c4" />
        <g stroke="#3c5d77" strokeWidth="1.4" fill="none" strokeLinecap="round">
          <path d="M28,30 q3.5,-4 7,0 q3.5,-4 7,0" />
          <path d="M46,22 q3,-3.4 6,0 q3,-3.4 6,0" />
        </g>
        <path d="M0,66 C30,52 56,52 78,60 C100,69 124,58 150,64 L150,120 L0,120 Z" fill="#8fcdeb" />
        <path d="M0,80 C28,64 54,86 80,78 C104,70 126,82 150,76 L150,120 L0,120 Z" fill="#3f8fc4" />
        <path d="M0,80 C28,64 54,86 80,78 C104,70 126,82 150,76" fill="none" stroke="#ffffff" strokeWidth="2.2" strokeLinecap="round" opacity="0.85" />
        <g stroke="#ffffff" strokeWidth="2" strokeLinecap="round" opacity="0.6"><line x1="26" y1="54" x2="42" y2="54" /><line x1="20" y1="62" x2="38" y2="62" /></g>
        <g transform="translate(82,66) rotate(-14)">
          <ellipse cx="0" cy="6" rx="36" ry="8" fill="url(#abBoard)" stroke="#9c4f24" strokeWidth="1" />
          <rect x="-34" y="4.6" width="68" height="2.5" rx="1.25" fill="#ffffff" opacity="0.55" />
          <path d="M-24,12.5 L-28,23 L-16,13.5 Z" fill="#9c4f24" />
        </g>
        <g fill="none" strokeLinejoin="round" strokeLinecap="round">
          <polyline points="74,47 69,57 71,66" stroke="#ffffff" strokeWidth="6" />
          <polyline points="86,47 92,56 90,63" stroke="#ffffff" strokeWidth="6" />
          <polyline points="74,47 69,57 71,66" stroke="#15324f" strokeWidth="4.4" />
          <polyline points="86,47 92,56 90,63" stroke="#15324f" strokeWidth="4.4" />
          <polyline points="74,47 69,57 71,66" stroke="#ff8a3d" strokeWidth="2.6" />
          <polyline points="86,47 92,56 90,63" stroke="#ff8a3d" strokeWidth="2.6" />
        </g>
        <g transform="translate(71,66) rotate(-14)"><rect x="-4.5" y="-1.8" width="9" height="3.6" rx="1.8" fill="#ff9d4a" stroke="#15324f" strokeWidth="1" /></g>
        <g transform="translate(90,63) rotate(-14)"><rect x="-4.5" y="-1.8" width="9" height="3.6" rx="1.8" fill="#ff9d4a" stroke="#15324f" strokeWidth="1" /></g>
        <text x="80" y="46" textAnchor="middle" fontFamily="Geneva, Helvetica, sans-serif" fontSize="42" fontWeight="bold" fill="none" stroke="#ffffff" strokeWidth="3.6" transform="rotate(-6 80 46)">Σ</text>
        <text x="80" y="46" textAnchor="middle" fontFamily="Geneva, Helvetica, sans-serif" fontSize="42" fontWeight="bold" fill="url(#abSun)" stroke="#15324f" strokeWidth="1.3" transform="rotate(-6 80 46)">Σ</text>
        <path d="M0,96 C34,86 56,101 88,94 C112,88 128,96 150,93 L150,120 L0,120 Z" fill="#235d8f" />
        <g fill="#ffffff" opacity="0.9"><circle cx="52" cy="70" r="2.1" /><circle cx="45" cy="66" r="1.5" /><circle cx="56" cy="63" r="1.3" /><circle cx="41" cy="72" r="1.3" /></g>
        <path d="M116,74 l1.4,3.4 3.4,1.4 -3.4,1.4 -1.4,3.4 -1.4,-3.4 -3.4,-1.4 3.4,-1.4 Z" fill="#ffffff" opacity="0.9" />
      </g>
    </svg>
  );
}
function AboutDialog({ onClose }) {
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 2000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: 372, background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11 }}>
        <div style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>About VibeStat</span>
        </div>
        <div style={{ padding: "18px 20px", background: "#f4f4f4", display: "flex", flexDirection: "column", alignItems: "center", gap: 9, textAlign: "center" }}>
          <AboutGraphic />
          <div style={{ fontSize: 18, fontWeight: "bold", letterSpacing: 0.3 }}>VibeStat</div>
          <div style={{ color: "#456" }}>Version 1.0</div>
          <div style={{ lineHeight: 1.5, color: "#222" }}>
            A desktop-style statistics workbench: descriptive statistics, frequency &amp; distribution analysis, regression and the general linear model, factorial &amp; repeated-measures ANOVA with contrasts and effect sizes, nonparametric tests, correlation &amp; reliability, contingency tables, and publication-ready graphics.
          </div>
          <div style={{ lineHeight: 1.5, color: "#222" }}>
            Every statistic is validated to high numerical precision against established reference implementations.
          </div>
          <div style={{ lineHeight: 1.5, color: "#222" }}>
            Created by <b>Stephen Maren</b>. Developed with assistance from Claude (Anthropic). Statistical results validated against SciPy, statsmodels, and pingouin.
          </div>
          <div style={{ alignSelf: "stretch", textAlign: "left", border: `1px solid ${PLAT.dark}`, background: "#fff", padding: "6px 8px" }}>
            <div style={{ fontSize: 10, fontWeight: "bold", marginBottom: 3 }}>How to cite</div>
            <div style={{ fontFamily: "Courier, monospace", fontSize: 9.5, lineHeight: 1.5, color: "#111" }}>
              Maren, S. (2026). VibeStat (Version 1.0) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.20720153
            </div>
            <div style={{ fontSize: 9, color: "#777", marginTop: 3 }}>DOI reserved on Zenodo; resolves once the first release is published.</div>
          </div>
          <div style={{ fontSize: 10, color: "#666", lineHeight: 1.45, borderTop: `1px solid ${PLAT.dark}`, paddingTop: 9, marginTop: 3 }}>
            An independent, clean-room reimplementation inspired by the classic StatView workflow. Not affiliated with, endorsed by, or derived from the commercial StatView product; all names are property of their respective owners. Released under the MIT License.
          </div>
          <div style={{ marginTop: 4 }}><MacBtn onClick={onClose} primary>OK</MacBtn></div>
        </div>
      </div>
    </div>
  );
}
function PzfxTableDialog({ pick, onPick, onCancel }) {
  const [sel, setSel] = useState(pick.activeIdx || 0);
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 2000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: 440, maxHeight: "80%", display: "flex", flexDirection: "column", background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11 }}>
        <div style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>Import Prism table</span>
        </div>
        <div style={{ padding: "12px 16px", background: "#f4f4f4", overflowY: "auto" }}>
          <div style={{ marginBottom: 8, color: "#222" }}>This Prism file contains <b>{pick.titles.length}</b> data tables. Choose one to import (the others are listed in the resulting Prism Import panel):</div>
          <div style={{ border: `1px solid ${PLAT.dark}`, background: "#fff", maxHeight: 260, overflowY: "auto" }}>
            {pick.titles.map((t, i) => (
              <label key={i} onClick={() => setSel(i)} style={{ display: "flex", alignItems: "flex-start", gap: 7, padding: "5px 8px", cursor: "pointer", borderBottom: i < pick.titles.length - 1 ? "1px solid #e3e3e3" : "none", background: sel === i ? PLAT.sel : "transparent" }}>
                <span style={{ display: "inline-block", width: 11, height: 11, borderRadius: "50%", ...bevelIn, background: sel === i ? PLAT.selBorder : "#fff", marginTop: 1, flex: "0 0 auto" }} />
                <span>
                  <span style={{ fontWeight: "bold" }}>{t || "Table " + (i + 1)}</span>
                  {i === (pick.activeIdx || 0) ? <span style={{ color: "#777", fontSize: 10 }}> · selected in Prism</span> : null}
                  <div style={{ color: "#556", fontSize: 10 }}>{pick.typeDescs[i]} · {pick.shapes[i]}</div>
                </span>
              </label>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, padding: "8px 16px", borderTop: `1px solid ${PLAT.dark}` }}>
          <MacBtn onClick={onCancel}>Cancel</MacBtn>
          <MacBtn onClick={() => onPick(sel)} primary>Import table</MacBtn>
        </div>
      </div>
    </div>
  );
}
function AnovaSetupDialog({ dlg, onApply, onCancel }) {
  const [cfg, setCfg] = useState({ ...DEFAULT_ANOVA_CFG, ...dlg.cfg });
  const { dragStyle, onHeaderDown } = useModalDrag();
  const set = (patch) => setCfg((c) => ({ ...c, ...patch }));
  const Radio = ({ name, val, label }) => (
    <label style={{ display: "block", cursor: "pointer", padding: "1px 0" }} onClick={() => set({ [name]: val })}>
      <span style={{ display: "inline-block", width: 11, height: 11, borderRadius: "50%", ...bevelIn, background: cfg[name] === val ? PLAT.selBorder : "#fff", marginRight: 6, verticalAlign: "middle" }} />
      {label}
    </label>
  );
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 2000, background: "rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ width: 380, background: PLAT.face, ...bevelOut, boxShadow: "3px 3px 0 rgba(0,0,0,0.4)", fontFamily: FONT, fontSize: 11, ...dragStyle }}>
        <div onMouseDown={onHeaderDown} title="drag to move" style={{ background: STRIPES, borderBottom: `1px solid ${PLAT.dark}`, padding: "3px 8px", fontWeight: "bold", textAlign: "center", cursor: "move" }}>
          <span style={{ background: PLAT.face, padding: "0 8px" }}>ANOVA Setup</span>
        </div>
        <div style={{ padding: 12, background: "#f4f4f4", display: "flex", flexDirection: "column", gap: 10 }}>
          <div>
            <div style={{ fontWeight: "bold" }}>Analysis</div>
            <Radio name="design" val="repeated" label="Repeated measures (use a compact variable)" />
            <Radio name="design" val="factorial" label="Factorial (between-subjects only)" />
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ fontWeight: "bold" }}>Alpha level</span>
            <input type="number" step="0.01" min="0.001" max="0.5" value={cfg.alpha} onChange={(e) => set({ alpha: Math.max(0.001, Math.min(0.5, Number(e.target.value) || 0.05)) })} style={{ width: 60, fontFamily: FONT, fontSize: 11, ...bevelIn, padding: "1px 4px" }} />
            <span style={{ color: "#678" }}>(default 0.05)</span>
          </div>
          <div>
            <div style={{ fontWeight: "bold" }}>Tables &amp; graphs show</div>
            <Radio name="effects" val="highest" label="Highest-order effect only" />
            <Radio name="effects" val="all" label="All effects" />
          </div>
          <div>
            <div style={{ fontWeight: "bold" }}>Error bars</div>
            <Radio name="errorBars" val="none" label="None" />
            <Radio name="errorBars" val="se" label="Standard error" />
            <Radio name="errorBars" val="sd" label="Standard deviation" />
            <Radio name="errorBars" val="ci" label={`${Math.round((1 - cfg.alpha) * 100)}% confidence interval`} />
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <MacBtn onClick={onCancel}>Cancel</MacBtn>
            <MacBtn onClick={() => onApply(cfg)} primary>OK</MacBtn>
          </div>
        </div>
      </div>
    </div>
  );
}
const anvCell = { border: "1px solid #ddd", padding: "1px 8px", textAlign: "right", fontFamily: "monospace" };

function Few({ n }) { return <div style={{ color: "#a00" }}>Need at least 3 complete cases (have {n ?? 0}).</div>; }

const _cellStr = (c) => String(c == null ? "" : c);
const tableToTSV = (head, rows) => [head, ...rows].map((r) => r.map(_cellStr).join("\t")).join("\n");
const tableToCSV = (head, rows) => [head, ...rows].map((r) => r.map((c) => { const s = _cellStr(c); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(",")).join("\n");
const copyTable = (head, rows) => { try { navigator.clipboard.writeText(tableToTSV(head, rows)); return true; } catch (e) { return false; } };
const downloadTableCSV = async (head, rows, name) => { const blob = new Blob([tableToCSV(head, rows)], { type: "text/csv;charset=utf-8" }); const fn = (name || "table") + ".csv"; if (!(await nativeSaveBytes(blob, fn))) downloadBlob(blob, fn); };
const tableExpBtn = { fontSize: 9, fontFamily: FONT, cursor: "pointer", border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "0 5px", borderRadius: 2, ...bevelOut };

function StatTable({ head, rows }) {
  const [hover, setHover] = useState(false);
  const [copied, setCopied] = useState(false);
  const copy = () => { if (copyTable(head, rows)) { setCopied(true); setTimeout(() => setCopied(false), 1100); } };
  return (
    <div style={{ position: "relative", display: "inline-block" }} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      {hover && (
        <div style={{ position: "absolute", top: -3, right: 0, transform: "translateY(-100%)", display: "flex", gap: 3, zIndex: 6 }} onClick={(e) => e.stopPropagation()}>
          <span style={tableExpBtn} onClick={copy} title="Copy as tab-separated (paste into Excel)">{copied ? "copied ✓" : "copy"}</span>
          <span style={tableExpBtn} onClick={() => downloadTableCSV(head, rows)} title="Download as CSV">CSV</span>
        </div>
      )}
      <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
        <thead>
          <tr>{head.map((h, i) => (
            <th key={i} style={{ border: `1px solid ${PLAT.dark}`, background: PLAT.face, padding: "1px 8px", textAlign: i === 0 ? "left" : "right", fontWeight: "bold" }}>{h}</th>
          ))}</tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => (
            <tr key={ri}>{r.map((c, ci) => (
              <td key={ci} style={{ border: "1px solid #ddd", padding: "1px 8px", textAlign: ci === 0 ? "left" : "right", fontFamily: ci === 0 ? FONT : "monospace" }}>{c}</td>
            ))}</tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* Wrap any bespoke (non-StatTable) table to add the same hover copy/CSV export, given its serialized head + rows. */
function ExportWrap({ head, rows, name, children }) {
  const [hover, setHover] = useState(false);
  const [copied, setCopied] = useState(false);
  const copy = () => { if (copyTable(head, rows)) { setCopied(true); setTimeout(() => setCopied(false), 1100); } };
  return (
    <div style={{ position: "relative", display: "inline-block" }} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      {hover && (
        <div style={{ position: "absolute", top: -3, right: 0, transform: "translateY(-100%)", display: "flex", gap: 3, zIndex: 6 }} onClick={(e) => e.stopPropagation()}>
          <span style={tableExpBtn} onClick={copy} title="Copy as tab-separated (paste into Excel)">{copied ? "copied ✓" : "copy"}</span>
          <span style={tableExpBtn} onClick={() => downloadTableCSV(head, rows, name || "table")} title="Download as CSV">CSV</span>
        </div>
      )}
      {children}
    </div>
  );
}
