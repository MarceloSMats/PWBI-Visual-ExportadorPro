"use strict";

/**
 * Visual: Exportador Pro
 * - Permite escolher colunas (campos) e medidas como colunas (buckets: "Colunas" e "Medidas")
 * - Permite escolher uma medida separada para o cabeçalho (bucket: "Medida do Cabeçalho")
 * - Exporta CSV e XLSX com cabeçalho dinâmico
 * - Respeita filtros/segmentações do relatório
 */

import "core-js/stable";
import "./../style/visual.less";

import powerbi from "powerbi-visuals-api";
import * as XLSX from "xlsx";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;
import DataView = powerbi.DataView;

type FileType = "csv" | "xlsx";

/** Utilitário: obtém valor do painel de formatação */
function getObjectValue<T>(
  objects: powerbi.DataViewObjects | undefined,
  objectName: string,
  propertyName: string,
  defaultValue: T
): T {
  if (!objects) return defaultValue;
  const obj = (objects as any)[objectName];
  if (obj && obj[propertyName] !== undefined && obj[propertyName] !== null) {
    return obj[propertyName] as T;
  }
  return defaultValue;
}

/** Escapa valores para CSV (aspas, quebras de linha, vírgulas) */
function csvEscape(v: any): string {
  if (v === null || v === undefined) return "";
  let s = String(v);
  // Normaliza quebras de linha
  s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  // Dobra aspas
  s = s.replace(/"/g, '""');
  // Envolve em aspas se contiver vírgula, aspas ou quebra de linha
  return /[",\n]/.test(s) ? `"${s}"` : s;
}

/** Converte valor para string segura */
function toSafeString(v: any): string {
  if (v === null || v === undefined) return "";
  return String(v);
}

/** Calcula largura aproximada de colunas para XLSX */
function computeColWidths(headers: string[], rows: any[][]): { wch: number }[] {
  const widths = headers.map(h => Math.max(10, h.length + 2));
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < headers.length; c++) {
      const val = rows[r][c];
      const len = val == null ? 0 : String(val).length;
      widths[c] = Math.min(Math.max(widths[c], len + 2), 60);
    }
  }
  return widths.map(w => ({ wch: w }));
}

/** Configurações vindas do painel de formatação */
class ExportSettings {
  fileType: FileType = "csv";
  fileName: string = "export";

  public static parse(dataView: DataView | undefined): ExportSettings {
    const s = new ExportSettings();
    const objects = dataView && dataView.metadata && dataView.metadata.objects;

    s.fileType = getObjectValue<FileType>(objects, "exportSettings", "fileType", "csv");
    s.fileName = getObjectValue<string>(objects, "exportSettings", "fileName", "export");

    // Sanitiza nome do arquivo
    if (!s.fileName || !s.fileName.trim()) s.fileName = "export";
    s.fileName = s.fileName.replace(/[\\/:*?"<>|]/g, "_");

    return s;
  }
}

export class Visual implements IVisual {
  // UI
  private root: HTMLElement;
  private headerEl: HTMLDivElement;
  private toolbarEl: HTMLDivElement;
  private btnCsv: HTMLButtonElement;
  private btnXlsx: HTMLButtonElement;
  private hintEl: HTMLDivElement;

  // Dados para exportação
  private exportHeaders: string[] = [];
  private exportRows: any[][] = [];
  private headerValue: string = "";

  // Configurações
  private settings: ExportSettings = new ExportSettings();

  constructor(options: VisualConstructorOptions) {
    this.root = document.createElement("div");
    this.root.className = "visual-container";

    // Cabeçalho do visual (dinâmico conforme formato padrão)
    this.headerEl = document.createElement("div");
    this.headerEl.className = "header";
    this.headerEl.innerText = "Exportador Pro";

    // Barra de botões
    this.toolbarEl = document.createElement("div");
    this.toolbarEl.className = "controls";

    this.btnCsv = document.createElement("button");
    this.btnCsv.className = "export-btn";
    this.btnCsv.innerText = "Exportar CSV";
    this.btnCsv.onclick = () => this.exportCsv();

    this.btnXlsx = document.createElement("button");
    this.btnXlsx.className = "export-btn";
    this.btnXlsx.innerText = "Exportar XLSX";
    this.btnXlsx.onclick = () => this.exportXlsx();

    this.toolbarEl.appendChild(this.btnCsv);
    this.toolbarEl.appendChild(this.btnXlsx);

    // Dica de uso
    this.hintEl = document.createElement("div");
    this.hintEl.className = "info";
    this.hintEl.innerText =
      "Arraste colunas/medidas para os buckets 'Colunas'/'Medidas' e a medida do cabeçalho para 'Medida do Cabeçalho'.";

    // Monta UI
    this.root.appendChild(this.headerEl);
    this.root.appendChild(this.toolbarEl);
    this.root.appendChild(this.hintEl);

    // Esconde botões inicialmente se não houver configurações
    this.updateUIStatus(0, 0);

    options.element.appendChild(this.root);
  }

  private updateUIStatus(rowCount: number, colCount: number): void {
    if (rowCount > 0 && colCount > 0) {
      this.hintEl.innerText = `${rowCount} linha(s) e ${colCount} coluna(s) prontas para exportação.`;
      this.hintEl.style.color = "#2b88d8";
      this.btnCsv.disabled = false;
      this.btnXlsx.disabled = false;
      this.btnCsv.style.opacity = "1";
      this.btnXlsx.style.opacity = "1";
    } else {
      this.hintEl.innerText = "Arraste campos para os buckets abaixo para habilitar a exportação.";
      this.hintEl.style.color = "#666";
      this.btnCsv.disabled = true;
      this.btnXlsx.disabled = true;
      this.btnCsv.style.opacity = "0.5";
      this.btnXlsx.style.opacity = "0.5";
    }
  }

  public update(options: VisualUpdateOptions): void {
    const dv: DataView | undefined = options.dataViews && options.dataViews[0];

    // Atualiza configurações do painel
    this.settings = ExportSettings.parse(dv);

    if (!dv || !dv.table || !dv.table.rows || dv.table.rows.length === 0) {
      // Sem dados → limpa estado
      this.exportHeaders = [];
      this.exportRows = [];
      this.headerValue = "";
      this.headerEl.innerText = `Exportador Pro – Formato padrão: ${this.settings.fileType.toUpperCase()}`;
      this.updateUIStatus(0, 0);
      return;
    }

    const table = dv.table;
    const columns = table.columns || [];
    const rows = table.rows || [];

    // Identifica índices pelos roles definidos em capabilities.json
    const dataIdxs: number[] = [];      // "columns" + "measures"
    const headerIdxs: number[] = [];    // "headerMeasure"

    columns.forEach((col, idx) => {
      const roles = (col.roles || {}) as Record<string, boolean>;
      if (roles["columns"] || roles["measures"]) dataIdxs.push(idx);
      if (roles["headerMeasure"]) headerIdxs.push(idx);
    });

    // Filtra duplicatas ou índices inválidos (segurança)
    const uniqueDataIdxs = Array.from(new Set(dataIdxs)).sort((a, b) => a - b);

    // Define cabeçalhos
    this.exportHeaders = uniqueDataIdxs.map(i => columns[i].displayName || columns[i].queryName || `Coluna ${i + 1}`);

    // Mapeia as linhas apenas com os índices selecionados (ordem preservada)
    this.exportRows = rows.map(r => uniqueDataIdxs.map(i => r[i]));

    // Obtém valor da medida do cabeçalho (primeira linha do índice correspondente)
    if (headerIdxs.length > 0) {
      const firstHeaderIndex = headerIdxs[0];
      this.headerValue = toSafeString(rows[0][firstHeaderIndex]);
    } else {
      this.headerValue = "";
    }

    // Atualiza título e status
    this.headerEl.innerText = `Exportador Pro – Formato padrão: ${this.settings.fileType.toUpperCase()}`;
    this.updateUIStatus(this.exportRows.length, this.exportHeaders.length);
  }

  /** Exporta em CSV com cabeçalho dinâmico */
  private exportCsv(): void {
    if (!this.exportHeaders.length) {
      alert("Nenhuma coluna selecionada. Adicione campos/medidas nos buckets 'Colunas' ou 'Medidas'.");
      return;
    }

    const lines: string[] = [];

    if (this.headerValue) lines.push(`Usuário: ${this.headerValue}`);
    lines.push(`Data: ${new Date().toLocaleString()}`);
    lines.push(""); // linha em branco

    // Cabeçalho das colunas
    lines.push(this.exportHeaders.map(csvEscape).join(","));

    // Dados
    for (const r of this.exportRows) {
      const rowValues = r.map(csvEscape);
      lines.push(rowValues.join(","));
    }

    const csv = lines.join("\r\n");
    const name = (this.settings.fileName || "export") + ".csv";
    this.downloadBlob(csv, name, "text/csv;charset=utf-8;");
  }

  /** Exporta em XLSX com cabeçalho dinâmico */
  private exportXlsx(): void {
    if (!this.exportHeaders.length) {
      alert("Nenhuma coluna selecionada. Adicione campos/medidas nos buckets 'Colunas' ou 'Medidas'.");
      return;
    }

    const aoa: any[][] = [];

    if (this.headerValue) aoa.push([`Usuário: ${this.headerValue}`]);
    aoa.push([`Data: ${new Date().toLocaleString()}`]);
    aoa.push([]); // linha em branco

    // Cabeçalho
    aoa.push(this.exportHeaders);

    // Dados
    for (const r of this.exportRows) {
      aoa.push(r);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Larguras aproximadas de coluna
    const colWidths = computeColWidths(this.exportHeaders, this.exportRows);
    (ws as any)["!cols"] = colWidths;

    const name = (this.settings.fileName || "export") + ".xlsx";
    XLSX.utils.book_append_sheet(wb, ws, "Export");
    XLSX.writeFile(wb, name);
  }

  /** Faz download de um blob/texto como arquivo */
  private downloadBlob(content: string, fileName: string, mime: string): void {
    try {
      const blob = new Blob([content], { type: mime });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      a.rel = "noopener";
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      // Fallback: tenta abrir em outra aba
      try {
        const win = window.open();
        if (win) {
          const pre = win.document.createElement("pre");
          pre.textContent = content; // textContent handles escaping automatically
          win.document.body.appendChild(pre);
          win.document.close();
        } else {
          alert("Falha ao iniciar download. Verifique bloqueadores de pop-up.");
        }
      } catch {
        alert("Falha ao iniciar download. Verifique bloqueadores de pop-up.");
      }
    }
  }

  /** Exposição das propriedades no painel de formatação */
  public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
    const enumeration: VisualObjectInstanceEnumeration = [];
    if (options.objectName === "exportSettings") {
      enumeration.push({
        objectName: "exportSettings",
        properties: {
          fileType: this.settings.fileType,
          fileName: this.settings.fileName
        },
        selector: null
      });
    }
    return enumeration;
  }
}