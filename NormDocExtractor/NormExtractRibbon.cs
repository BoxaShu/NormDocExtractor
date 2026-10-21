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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Text.RegularExpressions;

namespace NormDocExtractor
{
  public partial class NormExtractRibbon
  {
    private void btnExtractDocs_Click(object sender, RibbonControlEventArgs control)
    {
      var app = Globals.ThisAddIn.Application;
      Word.Document source = app.ActiveDocument;

      if (source == null)
      {
        System.Windows.Forms.MessageBox.Show(
            "NormExtract: откройте документ Word и повторите попытку.",
            "NormExtract",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
        return;
      }

      try
      {
        // 1. Извлекаем список уникальных документов
        var extracted = NormativeDocExtractor.Extract(source.Content.Text);
        var docs = extracted.UniqueDocs
            .OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (docs.Count == 0)
        {
          System.Windows.Forms.MessageBox.Show(
              "Нормативные документы в тексте не найдены.",
              "NormExtract",
              System.Windows.Forms.MessageBoxButtons.OK,
              System.Windows.Forms.MessageBoxIcon.Information);
          return;
        }

        // Документ должен быть сохранён, иначе закладки и ссылки не работают
        bool canLink = !string.IsNullOrEmpty(source.Path);

        // 2. Создаём документ-отчёт
        Word.Document report = app.Documents.Add();

        // Запоминаем путь к исходнику — понадобится для двойного клика
        if (!string.IsNullOrEmpty(source.FullName))
        {
          try
          {
            report.CustomDocumentProperties.Add(
                "NormExtract_SourcePath",
                false,                                      // LinkToContent
                Office.MsoDocProperties.msoPropertyTypeString,
                source.FullName);
          }
          catch (Exception ex)
          {
            System.Windows.Forms.MessageBox.Show(
                "Не удалось сохранить путь к исходнику: " + ex.Message,
                "NormExtract",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
          }
        }


        string linkWarning = canLink
            ? "💡 Двойной клик по номеру документа в столбце \"Номер документа\" в таблице — открыть диалог замены " +
              "во всех местах исходного файла.\r\n"
            : "⚠ Исходный документ ещё не сохранён на диск — " +
              "замена по двойному клику будет недоступна.\r\n";

        report.Content.Text =
            "NormDocExtract — список нормативных документов\r\n" +
            "Исходный документ: «" + source.Name + "»\r\n" +
            "Всего упоминаний: " + extracted.TotalMentions +
            "; уникальных документов: " + extracted.UniqueDocs.Count + "\r\n" +
            "Дата обработки: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "\r\n\r\n" +
            "Нажмите на ссылку в правой колонке, чтобы открыть страницу проверки " +
            "статуса документа в браузере.\r\n" +
            linkWarning + "\r\n";

        //report.Variables.Add("NormExtract_SourcePath", source.FullName);

        // 3. Таблица: 5 колонок
        Word.Range anchor = report.Content;
        anchor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

        Word.Table table = report.Tables.Add(
            anchor,
            docs.Count + 1,   // +1 — строка заголовка
            5);

        table.Borders.Enable = 1;
        table.Range.Font.Name = "Calibri";
        table.Range.Font.Size = 10;

        // ─── Заголовки ───
        string[] headers =
        {
    "№ п.п.",
    "Номер документа",
    "Сколько упоминаний",
    "Проверить на etr-torgi.ru",
    "Проверить на protect.gost.ru"
};
        for (int c = 0; c < headers.Length; c++)
        {
          Word.Cell cell = table.Cell(1, c + 1);
          cell.Range.Text = headers[c];
          cell.Range.Bold = 1;
          cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
        }

        // ─── Строки данных ───
        for (int r = 0; r < docs.Count; r++)
        {
          string docNum = docs[r].Key;
          int mentions = docs[r].Value;

          table.Cell(r + 2, 1).Range.Text = (r + 1).ToString();
          table.Cell(r + 2, 2).Range.Text = docNum;
          table.Cell(r + 2, 3).Range.Text = mentions.ToString();

          // Ссылка на etr-torgi.ru — работает для любого типа
          InsertHyperlink(report, table, r + 2, 4,
              BuildEtrTorgiUrl(docNum), "Открыть →");

          // Ссылка на protect.gost.ru — только для ГОСТ и СП
          string protectUrl = BuildProtectGostUrl(docNum);
          if (protectUrl != null)
          {
            InsertHyperlink(report, table, r + 2, 5, protectUrl, "Открыть →");
          }
          else
          {
            table.Cell(r + 2, 5).Range.Text = "—";
          }
        }

        table.Columns.AutoFit();
        report.Activate();
      }
      catch (Exception ex)
      {
        System.Windows.Forms.MessageBox.Show(
            "NormExtract: не удалось обработать документ.\r\n\r\n" + ex.Message,
            "NormExtract — ошибка",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
      }
    }

    // ─── Вставка гиперссылки в ячейку таблицы ──────────────────
    // Важно: Cell.Range включает маркер конца ячейки (\r\a),
    // поэтому сдвигаем конец диапазона на 1 символ назад — иначе
    // Word заменяет маркер и таблица «ломается».
    private static void InsertHyperlink(
        Word.Document report, Word.Table table,
        int row, int col, string url, string displayText)
    {
      Word.Range rng = table.Cell(row, col).Range;
      rng.MoveEnd(Word.WdUnits.wdCharacter, -1);
      rng.Text = displayText;
      report.Hyperlinks.Add(rng, url, Type.Missing, Type.Missing, Type.Missing);
    }

    // ─── Формирование URL ──────────────────────────────────────
    private static string BuildEtrTorgiUrl(string docNumber)
    {
      // Проверено: параметр ?query_light= подставляет номер в форму
      return "https://etr-torgi.ru/calc/check_gost/?query_light="
             + Uri.EscapeDataString(docNumber);
    }

    // ─── Формирование URL для protect.gost.ru ──────────────────────
    // Раздел зависит от типа документа:
    //   ГОСТ / ГОСТ Р → /gost
    //   СП            → /sp
    // Для остальных типов (СНиП, СанПиН, ТУ, РД, ...) отдельного
    // раздела на protect.gost.ru нет — возвращаем null.
    private static string BuildProtectGostUrl(string docNumber)
    {
      if (string.IsNullOrWhiteSpace(docNumber))
        return null;

      string trimmed = docNumber.Trim();
      string searchEncoded = Uri.EscapeDataString(trimmed);

      // ГОСТ, ГОСТ Р, ГОСТ ИСО и т.п.
      if (Regex.IsMatch(trimmed, @"^ГОСТ", RegexOptions.IgnoreCase))
        return "https://protect.gost.ru/gost?search=" + searchEncoded;

      // СП 50.13330.2012 и т.п.
      if (Regex.IsMatch(trimmed, @"^СП\s", RegexOptions.IgnoreCase))
        return "https://protect.gost.ru/sp?search=" + searchEncoded;

      if (Regex.IsMatch(trimmed, @"^СНиП\s", RegexOptions.IgnoreCase))
        return "https://protect.gost.ru/sp?search=" + searchEncoded;

      // Остальные типы — не поддерживаются
      return null;
    }
  }
}