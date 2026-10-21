// Copyright (c) 2026 Шульжицкий Владимир
//
// This file is part of NormDocExtract.
//
// NormDocExtract is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (any later version).
//
// NormDocExtract is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Word;

namespace NormDocExtractor
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
      this.Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
    }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
      this.Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
    }

    // ─── Перехват двойного клика ────────────────────────────
    private void Application_WindowBeforeDoubleClick(Word.Selection Sel, ref bool Cancel)
    {
      try
      {
        Word.Document doc = this.Application.ActiveDocument;
        if (doc == null) return;

        string sourcePath = GetReportSourcePath(doc);
        if (sourcePath == null)
        {
          System.Diagnostics.Debug.WriteLine("[DBL] EXIT: не отчёт NormExtract");
          return;
        }

        // ─── Cancel НЕ ставим — Word обновит Selection сам ───────
        // Через таймер читаем уже актуальное состояние.

        Word.Document reportDoc = doc;
        string path = sourcePath;

        var timer = new System.Windows.Forms.Timer { Interval = 200 };
        timer.Tick += (s, e) =>
        {
          timer.Stop();
          timer.Dispose();
          try
          {
            Word.Selection freshSel = this.Application.Selection;
            if (freshSel == null) return;

            if (!(bool)freshSel.Information[Word.WdInformation.wdWithInTable])
              return;

            if (freshSel.Cells.Count == 0)
              return;

            Word.Cell cell = freshSel.Cells[1];

            System.Diagnostics.Debug.WriteLine(
                "[DBL] Row=" + cell.RowIndex +
                " Col=" + cell.ColumnIndex +
                " header=\"" + SafeHeaderText(cell) + "\"" +
                " text=\"" + cell.Range.Text
                    .Replace("\r", "").Replace("\a", "").Trim() + "\"");

            // Не заголовок
            if (cell.RowIndex < 2) return;

            // Не тот столбец
            if (!IsDocumentNumberColumn(cell)) return;

            string cellText = cell.Range.Text
                .TrimEnd('\r', '\a', '\n', ' ').Trim();

            int wasIdx = cellText.IndexOf("(было:",
                                          StringComparison.OrdinalIgnoreCase);
            string docNumber = wasIdx >= 0
                ? cellText.Substring(0, wasIdx).Trim()
                : cellText;

            if (string.IsNullOrWhiteSpace(docNumber)) return;

            // Двойной клик выделил слово — снимаем выделение,
            // чтобы пользователь не видел артефакта
            try { freshSel.Collapse(); } catch { }

            ShowReplaceDialog(reportDoc, path, docNumber);
          }
          catch (Exception ex)
          {
            System.Diagnostics.Debug.WriteLine("[DBL] Timer: " + ex.Message);
          }
        };
        timer.Start();
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine("[DBL] Exception: " + ex);
      }
    }


    // ─── Безопасное чтение заголовка столбца ─────────────────
    // Используется только для диагностики. Возвращает текст
    // заголовка столбца, в котором находится ячейка, или
    // строку-маркер, если прочитать не удалось.
    private static string SafeHeaderText(Word.Cell cell)
    {
      try
      {
        Word.Table table = cell.Range.Tables[1];
        if (table == null) return "(нет таблицы)";

        return table.Cell(1, cell.ColumnIndex).Range.Text
            .Replace("\r", "")
            .Replace("\a", "")
            .Trim();
      }
      catch
      {
        return "(недоступен)";
      }
    }

    // ─── Проверка: клик в столбце «Номер документа»? ─────────
    // Устойчиво к «сбоям» Cell.ColumnIndex: сравниваем позиции
    // символов в документе, а не номер колонки. Работает и при
    // объединённых ячейках, и если пользователь вставил/удалил
    // столбец в отчёте.
    private static bool IsDocumentNumberColumn(Word.Cell clickedCell)
    {
      try
      {
        Word.Table table = clickedCell.Range.Tables[1];
        if (table == null) return false;

        int columns = table.Columns.Count;
        int targetCol = -1;

        // Ищем столбец, в заголовке которого есть «Номер документа»
        for (int c = 1; c <= columns; c++)
        {
          try
          {
            string header = table.Cell(1, c).Range.Text
                .Replace("\r", "")
                .Replace("\a", "")
                .Trim();

            if (header.IndexOf("Номер документа",
                               StringComparison.OrdinalIgnoreCase) >= 0)
            {
              targetCol = c;
              break;
            }
          }
          catch { /* объединённая/недоступная ячейка — пропускаем */ }
        }

        if (targetCol < 0) return false;

        // Сравниваем физические позиции клика и целевой ячейки
        Word.Range targetRange = table.Cell(clickedCell.RowIndex, targetCol).Range;
        Word.Range clickedRange = clickedCell.Range;

        return clickedRange.Start >= targetRange.Start
            && clickedRange.Start < targetRange.End;
      }
      catch
      {
        return false;
      }
    }


    // ─── Открытие диалога замены ───────────────────────────
    private void ShowReplaceDialog(Word.Document report, string sourcePath, string docNumber)
    {
      Word.Document sourceDoc = FindOpenDocument(this.Application, sourcePath);

      if (sourceDoc == null)
      {
        try
        {
          sourceDoc = this.Application.Documents.Open(sourcePath, ReadOnly: false);
        }
        catch (Exception ex)
        {
          System.Windows.Forms.MessageBox.Show(
              "Не удалось открыть исходный документ:\r\n" + sourcePath +
              "\r\n\r\n" + ex.Message,
              "NormExtract",
              System.Windows.Forms.MessageBoxButtons.OK,
              System.Windows.Forms.MessageBoxIcon.Warning);
          return;
        }
      }

      using (var dlg = new ReplaceDialog(sourcePath, docNumber))
      {
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK
            && dlg.ReplacementCount > 0)
        {
          UpdateReportRow(report, dlg.OldText, dlg.NewText);
        }
      }
    }

    // ─── Путь к исходнику из переменной отчёта ─────────────
    //private static string GetReportSourcePath(Word.Document doc)
    //{
    //  try
    //  {
    //    foreach (Word.Variable v in doc.Variables)
    //    {
    //      if (string.Equals(v.Name, "NormExtract_SourcePath",
    //                        StringComparison.OrdinalIgnoreCase))
    //        return v.Value as string;
    //    }
    //  }
    //  catch { }
    //  return null;
    //}
    private static string GetReportSourcePath(Word.Document doc)
    {
      try
      {
        Office.DocumentProperties props =
            (Office.DocumentProperties)doc.CustomDocumentProperties;

        foreach (Office.DocumentProperty prop in props)
        {
          if (string.Equals(prop.Name, "NormExtract_SourcePath",
                            StringComparison.OrdinalIgnoreCase))
          {
            string value = prop.Value as string;
            return string.IsNullOrWhiteSpace(value) ? null : value;
          }
        }
      }
      catch
      {
        // документ защищён, свойство отсутствует, COM-ошибка — не наша забота
      }
      return null;
    }



    // ─── Поиск уже открытого документа по пути ─────────────
    private static Word.Document FindOpenDocument(Word.Application app, string fullPath)
    {
      foreach (Word.Document d in app.Documents)
      {
        try
        {
          if (string.Equals(d.FullName, fullPath,
                            StringComparison.OrdinalIgnoreCase))
            return d;
        }
        catch { }
      }
      return null;
    }

    // ─── Обновление строки в отчёте после замены ───────────
    private static void UpdateReportRow(Word.Document report, string oldNum, string newNum)
    {
      foreach (Word.Table table in report.Tables)
      {
        for (int r = 2; r <= table.Rows.Count; r++)
        {
          try
          {
            Word.Range cellRng = table.Cell(r, 2).Range;
            string cellText = cellRng.Text
                .TrimEnd('\r', '\a', '\n', ' ').Trim();

            int wasIdx = cellText.IndexOf("(было:", StringComparison.OrdinalIgnoreCase);
            string currentNumber = wasIdx >= 0
                ? cellText.Substring(0, wasIdx).Trim()
                : cellText;

            if (!string.Equals(currentNumber, oldNum, StringComparison.OrdinalIgnoreCase))
              continue;

            // Определяем первоначальный номер — из существующей пометки, если она есть
            string originalNumber = oldNum;
            if (wasIdx >= 0)
            {
              Match m = Regex.Match(cellText, @"\(было:\s*([^)]+)\)");
              if (m.Success)
                originalNumber = m.Groups[1].Value.Trim();
            }

            // Записываем новый текст ячейки
            Word.Range editRng = table.Cell(r, 2).Range;
            editRng.MoveEnd(Word.WdUnits.wdCharacter, -1);
            editRng.Text = newNum + " (было: " + originalNumber + ")";

            // Форматируем пометку серым курсивом
            Word.Range cellFullRng = table.Cell(r, 2).Range;
            cellFullRng.MoveEnd(Word.WdUnits.wdCharacter, -1);
            string fullText = cellFullRng.Text;
            int noteIdx = fullText.IndexOf("(было:", StringComparison.OrdinalIgnoreCase);
            if (noteIdx >= 0)
            {
              Word.Range noteRng = report.Range(
                  cellFullRng.Start + noteIdx,
                  cellFullRng.End);
              noteRng.Italic = 1;
              noteRng.Font.Color = Word.WdColor.wdColorGray45;
              noteRng.Font.Size = 8;
            }
            return;
          }
          catch { }
        }
      }
    }








    #region Код, автоматически созданный VSTO

    /// <summary>
    /// Требуемый метод для поддержки конструктора — не изменяйте 
    /// содержимое этого метода с помощью редактора кода.
    /// </summary>
    private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
