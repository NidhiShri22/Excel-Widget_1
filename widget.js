/**
 * SAC Custom Widget - Excel Export Widget  (v1.2.0)
 *
 * Features:
 *  - Exports the bound SAC table to .xlsx via xlsx-js-style (SheetJS fork)
 *  - Bold, coloured, frozen header row
 *  - Optional logo / branding row at the top (user-configurable URL)
 *  - Sheet name = subtitle + timestamp + username (<=31 chars)
 *  - Grand-total and subtotal rows styled with grey background
 *  - Hierarchy / drill-down: Excel row outline groups (collapse / expand)
 *  - Rows collapsed in SAC start hidden in Excel
 *  - First-dim cells indented to match SAC hierarchy indentation
 *  - Visible / filtered data only; hidden columns excluded
 *  - Measure number-format preservation
 *  - Configurable file-name prefix and button appearance
 *
 * Dependency: xlsx-js-style loaded at runtime from jsDelivr CDN
 *   https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
 */

/* global customElements, HTMLElement, document, window, CustomEvent, Promise */

(function () {
  "use strict";

  /* ============================================================
     Constants
  ============================================================ */

  var XLSX_CDN =
    "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";

  var MAX_OUTLINE_LEVEL = 7;

  var COLOR = {
    HEADER_BG:    "FFFF00",
    HEADER_FG:    "000000",
    TOTAL_BG:     "D9D9D9",
    TOTAL_FG:     "000000",
    STRIPE_ODD:   "F5F5F5",
    STRIPE_EVEN:  "FFFFFF",
    BORDER:       "CCCCCC",
    HIER_LEAF_BG: "FAFAFA"
  };

  var THIN_BORDER = { style: "thin", color: { rgb: COLOR.BORDER } };
  var FULL_BORDER = {
    top:    THIN_BORDER,
    bottom: THIN_BORDER,
    left:   THIN_BORDER,
    right:  THIN_BORDER
  };

  /* ============================================================
     Shadow-DOM template
  ============================================================ */

  var tmpl = document.createElement("template");
  tmpl.innerHTML =
    '<link rel="stylesheet" href="https://raw.githubusercontent.com/NidhiShri22/Excel-Widget/refs/heads/main/widget.css" />' +
    '<div id="wrapper">' +
      '<button id="exportBtn" part="export-button" aria-label="Export to Excel">' +
        '<span id="btnLabel">Export to Excel</span>' +
      '</button>' +
      '<span id="statusMsg" role="status" aria-live="polite"></span>' +
    '</div>';

  /* ============================================================
     Web Component
  ============================================================ */

  var ExcelExportWidget = (function (_super) {

    function ExcelExportWidget() {
      var _this = _super.call(this) || this;
      _this._shadow = _this.attachShadow({ mode: "open" });
      _this._shadow.appendChild(tmpl.content.cloneNode(true));
      _this._props = {
        fileNamePrefix:    "SAC_Export",
        logoUrl:           "",
        buttonLabel:       "Export to Excel",
        buttonColor:       "#0070F2",
        buttonTextColor:   "#FFFFFF",
        headerBgColor:     "FFFF00",
        tableSubtitle:     "Export",
        username:          "",
        includeGrandTotal: true,
        logoRowHeight:     60
      };
      _this._xlsxReady   = false;
      _this._xlsxLoading = false;
      _this._autoUser    = "";   // auto-detected from SAC AppFrameContext
      _this.tableDataBinding = null;
      return _this;
    }

    ExcelExportWidget.prototype = Object.create(_super.prototype);
    ExcelExportWidget.prototype.constructor = ExcelExportWidget;

    /* SAC lifecycle */

    ExcelExportWidget.prototype.connectedCallback = function () {
      this._btn      = this._shadow.getElementById("exportBtn");
      this._btnLabel = this._shadow.getElementById("btnLabel");
      this._statusEl = this._shadow.getElementById("statusMsg");
      var self = this;
      this._btn.addEventListener("click", function () { self.exportToExcel(); });
      this._applyButtonStyle();
      this._autoDetectUser();
      this._ensureXLSX().catch(function () {});
    };

    ExcelExportWidget.prototype.onCustomWidgetBeforeUpdate = function () {};

    ExcelExportWidget.prototype.onCustomWidgetAfterUpdate = function (changedProps) {
      var self = this;
      Object.keys(this._props).forEach(function (k) {
        if (Object.prototype.hasOwnProperty.call(changedProps, k)) {
          self._props[k] = changedProps[k];
        }
      });
      this._applyButtonStyle();
    };

    ExcelExportWidget.prototype.onCustomWidgetDataChanged = function () {
      this._setStatus("");
    };

    /**
     * Tries to resolve the SAC logged-in user from the global AppFrameContext.
     * Called once in connectedCallback; result stored in _autoUser.
     * Falls back silently if the context is unavailable (e.g. outside SAC).
     */
    ExcelExportWidget.prototype._autoDetectUser = function () {
      try {
        var sap = window.sap;
        var ctx = sap && sap.fpa && sap.fpa.ui &&
                  sap.fpa.ui.infra && sap.fpa.ui.infra.appframe &&
                  sap.fpa.ui.infra.appframe.AppFrameContext;
        if (!ctx) return;
        var u = (ctx.getUser && ctx.getUser()) ||
                (ctx.getCurrentUser && ctx.getCurrentUser());
        if (u) {
          this._autoUser = (u.getId && u.getId()) || u.id || u.name || u.email || "";
        }
      } catch (e) { /* SAC context not available – silent fallback */ }
    };

    /* Public API */

    ExcelExportWidget.prototype.exportToExcel = function () {
      var self = this;
      if (this._btn) this._btn.disabled = true;
      this._setStatus("Preparing\u2026");
      this._ensureXLSX()
        .then(function () { return self._runExport(); })
        .catch(function (err) {
          console.error("[ExcelExportWidget]", err);
          self._setStatus("Error \u2013 see browser console");
          self._fireEvent("onExportError", { error: String(err.message || err) });
        })
        .finally(function () {
          if (self._btn) self._btn.disabled = false;
        });
    };

    /* Button styling */

    ExcelExportWidget.prototype._applyButtonStyle = function () {
      if (!this._btn) return;
      this._btnLabel.textContent      = this._props.buttonLabel     || "Export to Excel";
      this._btn.style.backgroundColor = this._props.buttonColor     || "#0070F2";
      this._btn.style.color           = this._props.buttonTextColor || "#FFFFFF";
    };

    ExcelExportWidget.prototype._setStatus = function (msg) {
      if (this._statusEl) this._statusEl.textContent = msg;
    };

    /* SheetJS loader */

    ExcelExportWidget.prototype._ensureXLSX = function () {
      var self = this;
      if (window.XLSXStyle) { this._xlsxReady = true; return Promise.resolve(); }
      if (this._xlsxLoading) {
        return new Promise(function (resolve, reject) {
          var start = Date.now();
          var id = setInterval(function () {
            if (window.XLSXStyle) { clearInterval(id); self._xlsxReady = true; resolve(); }
            else if (Date.now() - start > 20000) { clearInterval(id); reject(new Error("SheetJS load timeout.")); }
          }, 100);
        });
      }
      this._xlsxLoading = true;
      return new Promise(function (resolve, reject) {
        var s = document.createElement("script");
        s.src = XLSX_CDN;
        s.onload = function () { self._xlsxReady = true; self._xlsxLoading = false; resolve(); };
        s.onerror = function () {
          self._xlsxLoading = false;
          reject(new Error("Failed to load SheetJS. Bundle xlsx.bundle.js into the package and set XLSX_CDN to a relative path."));
        };
        document.head.appendChild(s);
      });
    };

    /* ═══════════════════════════════ CORE EXPORT ════════════════════════════ */

    ExcelExportWidget.prototype._runExport = function () {
      var self    = this;
      var binding = this.tableDataBinding;
      if (!binding) {
        return Promise.reject(new Error("No data binding configured. Connect to a data source in widget properties."));
      }

      return Promise.resolve(binding.getData ? binding.getData() : binding)
        .then(function (result) {
          if (!result || !result.data) throw new Error("Data binding returned no data.");

          var XLSX = window.XLSXStyle;

          /* 1. Metadata */
          var meta     = result.metadata || {};
          var feedDims = (meta.feeds && meta.feeds.dimensions && meta.feeds.dimensions.members) || [];
          var feedMeas = (meta.feeds && meta.feeds.measures   && meta.feeds.measures.members)   || [];
          var visDims  = feedDims.filter(function (m) { return m.visible !== false; });
          var visMeas  = feedMeas.filter(function (m) { return m.visible !== false; });
          var allCols  = visDims.concat(visMeas);
          if (allCols.length === 0) throw new Error("No visible columns found.");

          /* 1a. Auto-detect config from table binding metadata */
          // Subtitle: prefer explicit prop, fall back to metadata title/subtitle/name
          var _autoSubtitle = meta.subtitle || meta.title || meta.name || meta.dataSetName ||
                              (meta.dataSource && (meta.dataSource.description || meta.dataSource.name)) || "";
          var _effSubtitle  = (self._props.tableSubtitle && self._props.tableSubtitle !== "Export")
                                ? self._props.tableSubtitle
                                : (_autoSubtitle || "Export");

          // Username: prefer explicit prop, fall back to SAC AppFrameContext user
          var _effUser      = (self._props.username && self._props.username !== "")
                                ? self._props.username
                                : (self._autoUser || "");

          // File prefix: prefer explicit prop, fall back to metadata name cleaned for filenames
          var _rawAutoPrefix = meta.name || meta.dataSetName || meta.title ||
                               (meta.dataSource && (meta.dataSource.name || meta.dataSource.description)) || "";
          var _autoPrefix   = _rawAutoPrefix.replace(/[^a-zA-Z0-9_\-]/g, "_").replace(/_+/g,"_")
                                            .replace(/^_|_$/g,"").substring(0, 30);
          var _effPrefix    = (self._props.fileNamePrefix && self._props.fileNamePrefix !== "SAC_Export")
                                ? self._props.fileNamePrefix
                                : (_autoPrefix || "SAC_Export");

          var colMap = self._buildColIndexMap(feedDims, feedMeas);

          /* 2. Hierarchy detection */
          var rawRows      = result.data;
          var hasHierarchy = self._detectHierarchy(rawRows);

          /* 3. Build AOA + per-row flag arrays */
          var aoa        = [];
          var totalFlags = [];
          var levelFlags = [];
          var drillFlags = [];
          var hasLogo    = !!(self._props.logoUrl && self._props.logoUrl.trim());

          if (hasLogo) {
            aoa.push(new Array(allCols.length).fill(""));
            totalFlags.push(false); levelFlags.push(0); drillFlags.push("leaf");
          }

          var HEADER_ROW = aoa.length;
          aoa.push(allCols.map(function (c) { return c.description || c.label || c.id || ""; }));
          totalFlags.push(false); levelFlags.push(0); drillFlags.push("leaf");

          rawRows.forEach(function (raw) {
            var isTotal = self._isTotalRow(raw, feedMeas);
            if (isTotal && !self._props.includeGrandTotal) return;

            var lvl   = hasHierarchy ? self._getRowLevel(raw)      : 0;
            var drill = hasHierarchy ? self._getRowDrillState(raw) : "leaf";

            var row = allCols.map(function (col) {
              var idx  = colMap.get(col.id);
              if (idx === undefined) return "";
              var cell = raw[idx];
              if (cell === null || cell === undefined) return "";
              var isMeas = feedMeas.some(function (m) { return m.id === col.id; });
              if (isMeas) {
                var v = (cell.value !== undefined) ? cell.value : cell;
                var n = Number(v);
                return isNaN(n) ? (cell.formattedValue || cell.value || "") : n;
              }
              return (cell.formattedValue !== undefined) ? cell.formattedValue
                   : (cell.value         !== undefined) ? cell.value
                   : String(cell);
            });

            aoa.push(row);
            totalFlags.push(isTotal);
            levelFlags.push(lvl);
            drillFlags.push(drill);
          });

          /* 4. Worksheet */
          var ws    = XLSX.utils.aoa_to_sheet(aoa);
          var nRows = aoa.length;
          var nCols = allCols.length;
          ws["!ref"] = XLSX.utils.encode_range({ r:0, c:0 }, { r:nRows-1, c:nCols-1 });

          /* 5. Column widths */
          ws["!cols"] = allCols.map(function (c) {
            return { wch: Math.max(14, ((c.description || c.label || c.id || "").length) + 6) };
          });

          /* 6. Row heights */
          var rh = [];
          if (hasLogo) rh[0] = { hpt: Number(self._props.logoRowHeight) || 60 };
          rh[HEADER_ROW] = { hpt: 22 };
          ws["!rows"] = rh;

          /* 7. Freeze panes */
          ws["!freeze"] = { xSplit: 0, ySplit: HEADER_ROW + 1 };

          /* 8. Logo merge */
          ws["!merges"] = [];
          if (hasLogo) ws["!merges"].push({ s:{r:0,c:0}, e:{r:0,c:nCols-1} });

          /* 9. Hierarchy grouping */
          if (hasHierarchy) {
            self._applyHierarchyGrouping(ws, HEADER_ROW, levelFlags, drillFlags, totalFlags);
          }

          /* 10. Cell styles */
          for (var r = 0; r < nRows; r++) {
            var isLogoRow   = hasLogo && r === 0;
            var isHeaderRow = r === HEADER_ROW;
            var isTotalRow  = totalFlags[r] === true;
            var rowLvl      = levelFlags[r] || 0;
            for (var c = 0; c < nCols; c++) {
              var ref = XLSX.utils.encode_cell({ r:r, c:c });
              if (!ws[ref]) ws[ref] = { v:"", t:"s" };
              if      (isLogoRow)   ws[ref].s = self._logoRowStyle();
              else if (isHeaderRow) ws[ref].s = self._headerCellStyle();
              else if (isTotalRow)  ws[ref].s = self._totalCellStyle(c, visDims.length, rowLvl);
              else                  ws[ref].s = self._dataCellStyle(r, c, visDims.length, rowLvl, hasHierarchy);
            }
          }

          /* 11. Number formats */
          var mc0 = visDims.length;
          for (var dr = HEADER_ROW + 1; dr < nRows; dr++) {
            visMeas.forEach(function (meas, mi) {
              if (!meas.format) return;
              var rf = XLSX.utils.encode_cell({ r:dr, c:mc0+mi });
              if (ws[rf]) ws[rf].z = self._sacFormatToExcel(meas.format);
            });
          }

          /* 12. Logo text */
          if (hasLogo) {
            var lr = XLSX.utils.encode_cell({ r:0, c:0 });
            if (!ws[lr]) ws[lr] = {};
            ws[lr].v = "[Logo: " + self._props.logoUrl + "]";
            ws[lr].t = "s";
            ws[lr].s = self._logoRowStyle();
          }

          /* 13. Write */
          var wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, self._buildSheetName(_effSubtitle, _effUser));
          XLSX.writeFile(wb, self._buildFileName(_effPrefix), { compression: true });

          self._setStatus("Downloaded \u2713");
          setTimeout(function () { self._setStatus(""); }, 4000);
          self._fireEvent("onExportSuccess", {});
        });
    };

    /* ═══════════════════════════════ HIERARCHY ══════════════════════════════ */

    ExcelExportWidget.prototype._detectHierarchy = function (rawRows) {
      if (!Array.isArray(rawRows)) return false;
      for (var i = 0; i < rawRows.length; i++) {
        var row = rawRows[i];
        if (!Array.isArray(row)) continue;
        for (var j = 0; j < row.length; j++) {
          var cell = row[j];
          if (!cell || typeof cell !== "object") continue;
          if (typeof cell.level          === "number" && cell.level > 0)          return true;
          if (typeof cell.hierarchyLevel === "number" && cell.hierarchyLevel > 0) return true;
          if (cell.drillState === "expanded" || cell.drillState === "collapsed")   return true;
        }
      }
      return false;
    };

    ExcelExportWidget.prototype._getRowLevel = function (rawRow) {
      if (!Array.isArray(rawRow)) return 0;
      for (var j = 0; j < rawRow.length; j++) {
        var cell = rawRow[j];
        if (!cell || typeof cell !== "object") continue;
        if (typeof cell.level          === "number") return cell.level;
        if (typeof cell.hierarchyLevel === "number") return cell.hierarchyLevel;
      }
      return 0;
    };

    ExcelExportWidget.prototype._getRowDrillState = function (rawRow) {
      if (!Array.isArray(rawRow)) return "leaf";
      for (var j = 0; j < rawRow.length; j++) {
        var cell = rawRow[j];
        if (!cell || typeof cell !== "object") continue;
        if (cell.drillState) return cell.drillState;
      }
      return "leaf";
    };

    /**
     * Applies Excel row outline grouping.
     *
     * SAC level 0  = root / flat  -> no outline group
     * SAC level N  = depth N      -> Excel outline level N  (max 7)
     *
     * A row whose nearest collapsed ancestor has a lower level starts hidden,
     * mirroring the SAC story's collapsed state.
     *
     * ws['!outline'] = { above: true } places summary rows ABOVE their groups.
     */
    ExcelExportWidget.prototype._applyHierarchyGrouping = function (
      ws, HEADER_ROW, levelFlags, drillFlags, totalFlags
    ) {
      if (!ws["!rows"]) ws["!rows"] = [];

      var dataStart     = HEADER_ROW + 1;
      var ancestorStack = [];   // { sacLevel, isCollapsed }

      for (var i = dataStart; i < levelFlags.length; i++) {
        var sacLevel = levelFlags[i]  || 0;
        var drillSt  = drillFlags[i]  || "leaf";
        var isTotal  = totalFlags[i]  === true;

        while (ws["!rows"].length <= i) ws["!rows"].push(null);
        if (!ws["!rows"][i]) ws["!rows"][i] = {};

        /* pop ancestors no longer applicable */
        while (
          ancestorStack.length > 0 &&
          ancestorStack[ancestorStack.length - 1].sacLevel >= sacLevel
        ) { ancestorStack.pop(); }

        /* hidden if any ancestor is collapsed */
        var startHidden = false;
        if (sacLevel > 0) {
          for (var a = 0; a < ancestorStack.length; a++) {
            if (ancestorStack[a].isCollapsed && ancestorStack[a].sacLevel < sacLevel) {
              startHidden = true; break;
            }
          }
        }

        if (sacLevel > 0 && !isTotal) {
          ws["!rows"][i].level  = Math.min(sacLevel, MAX_OUTLINE_LEVEL);
          ws["!rows"][i].hidden = startHidden;
        }

        if (drillSt === "expanded" || drillSt === "collapsed") {
          ancestorStack.push({ sacLevel: sacLevel, isCollapsed: drillSt === "collapsed" });
        }
      }

      /* summary rows are ABOVE their detail groups (SAC default) */
      ws["!outline"] = { above: true };
    };

    /* ═══════════════════════════════ STYLES ═════════════════════════════════ */

    ExcelExportWidget.prototype._logoRowStyle = function () {
      return {
        font:      { bold: true, sz: 12, color: { rgb: "444444" } },
        fill:      { patternType: "solid", fgColor: { rgb: "FFFFFF" } },
        alignment: { horizontal: "left", vertical: "center" }
      };
    };

    ExcelExportWidget.prototype._headerCellStyle = function () {
      var bg = this._normaliseHex(this._props.headerBgColor, COLOR.HEADER_BG);
      return {
        font:      { bold: true, sz: 11, color: { rgb: COLOR.HEADER_FG } },
        fill:      { patternType: "solid", fgColor: { rgb: bg } },
        alignment: { horizontal: "center", vertical: "center", wrapText: false },
        border:    FULL_BORDER
      };
    };

    ExcelExportWidget.prototype._totalCellStyle = function (colIdx, measStartCol, rowLvl) {
      var bg = (rowLvl > 0) ? "E8E8E8" : COLOR.TOTAL_BG;
      return {
        font:      { bold: true, sz: 11, color: { rgb: COLOR.TOTAL_FG } },
        fill:      { patternType: "solid", fgColor: { rgb: bg } },
        alignment: {
          horizontal: colIdx >= measStartCol ? "right" : "left",
          vertical:   "center",
          indent:     colIdx === 0 ? Math.min(rowLvl * 2, 14) : 0
        },
        border: FULL_BORDER
      };
    };

    ExcelExportWidget.prototype._dataCellStyle = function (rowIdx, colIdx, measStartCol, rowLvl, hasHierarchy) {
      var bg = (hasHierarchy && rowLvl > 0)
        ? (rowIdx % 2 === 0 ? COLOR.HIER_LEAF_BG : COLOR.STRIPE_ODD)
        : (rowIdx % 2 === 0 ? COLOR.STRIPE_EVEN  : COLOR.STRIPE_ODD);
      var indent = (hasHierarchy && colIdx === 0 && rowLvl > 0)
        ? Math.min(rowLvl * 2, 14) : 0;
      return {
        font:      { sz: 11 },
        fill:      { patternType: "solid", fgColor: { rgb: bg } },
        alignment: {
          horizontal: colIdx >= measStartCol ? "right" : "left",
          vertical:   "center",
          indent:     indent
        },
        border: FULL_BORDER
      };
    };

    /* ═══════════════════════════════ UTILITIES ══════════════════════════════ */

    ExcelExportWidget.prototype._buildColIndexMap = function (feedDims, feedMeas) {
      var map = new Map();
      feedDims.forEach(function (m, i) { map.set(m.id, i); });
      feedMeas.forEach(function (m, i) { map.set(m.id, feedDims.length + i); });
      return map;
    };

    ExcelExportWidget.prototype._isTotalRow = function (rawRow, feedMeas) {
      if (!Array.isArray(rawRow)) return false;
      return rawRow.some(function (cell) {
        if (!cell || typeof cell !== "object") return false;
        return cell.type === "ResultCell"  || cell.type === "TOTAL"       ||
               cell.type === "SUBTOTAL"    || cell.type === "GRAND_TOTAL" ||
               cell.isGrandTotal === true  || cell.isSubTotal === true    ||
               cell.isTotal      === true;
      });
    };

    ExcelExportWidget.prototype._sacFormatToExcel = function (fmt) {
      if (!fmt) return "@";
      var m = { "0":"0","0.0":"0.0","0.00":"0.00","#,##0":"#,##0",
                "#,##0.0":"#,##0.0","#,##0.00":"#,##0.00",
                "0%":"0%","0.0%":"0.0%","0.00%":"0.00%" };
      if (m[fmt]) return m[fmt];
      if (fmt.indexOf("$")       >= 0) return '"$"#,##0.00';
      if (fmt.indexOf("\u20ac")  >= 0) return '[$\u20ac-407]#,##0.00';
      if (fmt.indexOf("\u00a3")  >= 0) return '[$\u00a3-809]#,##0.00';
      return fmt;
    };

    /**
     * @param {string} [subtitle]  Effective subtitle (auto-detected or from prop)
     * @param {string} [user]      Effective username  (auto-detected or from prop)
     */
    ExcelExportWidget.prototype._buildSheetName = function (subtitle, user) {
      var s = (subtitle !== undefined ? subtitle : (this._props.tableSubtitle || "Export"))
                .trim().replace(/[\\\/\[\]\*\?:]/g,"_");
      var t = this._formatTimestampShort(new Date());
      var u = (user    !== undefined ? user    : (this._props.username || ""))
                .trim().replace(/[\\\/\[\]\*\?:]/g,"_");
      return [s,t,u].filter(Boolean).join(" ").substring(0,31);
    };

    /**
     * @param {string} [prefix]  Effective file-name prefix (auto-detected or from prop)
     */
    ExcelExportWidget.prototype._buildFileName = function (prefix) {
      var p = (prefix !== undefined ? prefix : (this._props.fileNamePrefix || "SAC_Export"))
                .trim().replace(/[^a-zA-Z0-9_\-]/g,"_");
      return p + "_" + this._formatTimestampLong(new Date()) + ".xlsx";
    };

    ExcelExportWidget.prototype._formatTimestampShort = function (d) {
      return d.getFullYear()+"-"+this._pad(d.getMonth()+1)+"-"+this._pad(d.getDate())+
             " "+this._pad(d.getHours())+":"+this._pad(d.getMinutes());
    };

    ExcelExportWidget.prototype._formatTimestampLong = function (d) {
      return d.getFullYear()+"-"+this._pad(d.getMonth()+1)+"-"+this._pad(d.getDate())+
             "_"+this._pad(d.getHours())+this._pad(d.getMinutes());
    };

    ExcelExportWidget.prototype._pad = function (n) {
      return String(n).padStart(2,"0");
    };

    ExcelExportWidget.prototype._normaliseHex = function (hex, fallback) {
      if (!hex) return fallback;
      var c = hex.replace(/^#/,"").toUpperCase();
      return /^[0-9A-F]{6}$/.test(c) ? c : fallback;
    };

    ExcelExportWidget.prototype._fireEvent = function (name, detail) {
      this.dispatchEvent(new CustomEvent(name, { bubbles:true, detail:detail }));
    };

    return ExcelExportWidget;
  }(HTMLElement));

  customElements.define("com-custom-sac-excel-export-widget", ExcelExportWidget);

})();
