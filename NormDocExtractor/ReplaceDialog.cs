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
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace NormDocExtractor
{
  public partial class ReplaceDialog : Form
  {
    private readonly Word.Document _doc;

    public string OldText => txtOld.Text.Trim();
    public string NewText => txtNew.Text.Trim();
    public int ReplacementCount { get; private set; }

    public ReplaceDialog(string sourceFullName, string preselectedNumber)
    {
      InitializeComponent();

      _doc = FindOpenDoc(sourceFullName);
      txtSource.Text = sourceFullName;

      // ─── txtOld — только для чтения ────────────────────────
      txtOld.Text = preselectedNumber ?? "";
      txtOld.ReadOnly = true;
      txtOld.BackColor = System.Drawing.SystemColors.Control;  // визуально «серое»

      // ─── Предзаполнение txtNew ─────────────────────────────
      txtNew.Text = TrimLastTwo(preselectedNumber);

      if (string.IsNullOrEmpty(preselectedNumber))
        lblCount.Text = "Найдено: —";

      // ─── Обработчики ───────────────────────────────────────
      btnCount.Click += (s, e) => DoCount();
      btnReplace.Click += (s, e) => DoReplace();
      txtOld.TextChanged += (s, e) => lblCount.Text = "Найдено: —";

      // Блокируем «Заменить всё», пока год не дописан
      txtNew.TextChanged += (s, e) => UpdateReplaceButtonState();

      UpdateReplaceButtonState();   // исходное состояние
    }

    // ─── Отрезает два последних символа (уже есть) ─────────────
    private static string TrimLastTwo(string s)
    {
      if (string.IsNullOrEmpty(s)) return "";
      return s.Length <= 2 ? "" : s.Substring(0, s.Length - 2);
    }

    // ─── Проверка: год новой версии дописан? ───────────────────
    // Условия:
    //   1. Пользователь добавил к обрезанному номеру хотя бы 1 символ.
    //   2. Строка заканчивается минимум двумя цифрами (год).
    // Примеры:
    //   old="ГОСТ 12.1.005-88", new="ГОСТ 12.1.005-"     → false (ничего не дописано)
    //   old="ГОСТ 12.1.005-88", new="ГОСТ 12.1.005-2"    → false (только 1 цифра)
    //   old="ГОСТ 12.1.005-88", new="ГОСТ 12.1.005-21"   → true  (2 цифры)
    //   old="ГОСТ 12.1.005-88", new="ГОСТ 12.1.005-2024" → true  (4 цифры)
    private static bool IsYearCompleted(string newText, string oldText)
    {
      if (string.IsNullOrWhiteSpace(newText)) return false;

      string trimmedOld = TrimLastTwo(oldText);
      string n = newText.TrimEnd();

      // Пользователь должен выйти за пределы «обрезанного» номера
      if (n.Length <= trimmedOld.Length) return false;

      // В конце должно быть минимум 2 цифры подряд (год)
      int tailDigits = 0;
      for (int i = n.Length - 1; i >= 0 && char.IsDigit(n[i]); i--)
        tailDigits++;

      //return tailDigits >= 2;
      // Разрешаем замену, если пользователь явно ввёл что-то длиннее 0 символов
      // и это отличается от обрезанной версии
      return n.Length > trimmedOld.Length;
    }

    // ─── Управление состоянием кнопки «Заменить всё» ───────────
    private void UpdateReplaceButtonState()
    {
      bool yearOk = IsYearCompleted(txtNew.Text, txtOld.Text);
      btnReplace.Enabled = yearOk;

      // Подсказка в метке: пока год не дописан — «допишите год»
      lblNew.Text = yearOk ? "Заменить на:" : "Заменить на (допишите год):";
    }


    private static Word.Document FindOpenDoc(string fullPath)
    {
      var app = Globals.ThisAddIn.Application;
      foreach (Word.Document d in app.Documents)
      {
        try
        {
          if (string.Equals(d.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
            return d;
        }
        catch { }
      }
      return null;
    }

    // ─── Подсчёт вхождений ──────────────────────────────────
    private void DoCount()
    {
      if (_doc == null || string.IsNullOrWhiteSpace(OldText))
      {
        lblCount.Text = "Найдено: 0";
        return;
      }

      lblCount.Text = "Найдено: …";
      Application.DoEvents();

      int count = CountOccurrences(_doc, OldText);
      lblCount.Text = "Найдено: " + count;
    }

    // ─── Замена всех вхождений ──────────────────────────────
    private void DoReplace()
    {

      // ─── Проверка «год дописан» ────────────────────────────
      if (!IsYearCompleted(txtNew.Text, txtOld.Text))
      {
        MessageBox.Show(
            "Номер новой версии не дописан — укажите полный номер с годом.\r\n\r\n" +
            "Пример: «" + TrimLastTwo(txtOld.Text) + "2024»",
            "NormExtract",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      if (_doc == null)
      {
        MessageBox.Show("Исходный документ не открыт.", "NormExtract",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }
      if (_doc.ReadOnly)
      {
        MessageBox.Show("Исходный документ открыт в режиме ТОЛЬКО ДЛЯ ЧТЕНИЯ. Выполните \"Сохранить как\" для продолжения работы", "NormExtract",
    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

        if (string.IsNullOrWhiteSpace(OldText))
      {
        MessageBox.Show("Введите номер, который нужно найти.", "NormExtract",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      int count = CountOccurrences(_doc, OldText);
      if (count == 0)
      {
        MessageBox.Show("Указанный номер в документе не найден.", "NormExtract",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      var confirm = MessageBox.Show(
          "Будет заменено вхождений: " + count +
          "\r\n\r\n«" + OldText + "»  →  «" + NewText + "»" +
          "\r\n\r\nПродолжить?",
          "NormExtract — подтверждение",
          MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

      if (confirm != DialogResult.OK) return;

      ReplaceAll(_doc, OldText, NewText);
      ReplacementCount = count;

      MessageBox.Show(
          "Заменено вхождений: " + count + ".",
          "NormExtract",
          MessageBoxButtons.OK, MessageBoxIcon.Information);

      DialogResult = DialogResult.OK;
      Close();
    }

    // ─── Поиск всех вхождений в исходнике ───────────────────
    private static int CountOccurrences(Word.Document doc, string text)
    {
      try
      {
        Word.Range rng = doc.Content;
        Word.Find find = rng.Find;
        find.ClearFormatting();
        find.Text = text;
        find.Forward = true;
        find.Wrap = Word.WdFindWrap.wdFindStop;
        find.MatchCase = false;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        find.Format = false;

        int count = 0;
        while (find.Execute() && count < 10000)
        {
          count++;
          if (rng.End >= doc.Content.End - 1) break;
        }
        return count;
      }
      catch { return 0; }
    }

    // ─── Замена всех вхождений ──────────────────────────────
    private static void ReplaceAll(Word.Document doc, string oldText, string newText)
    {
      Word.Range rng = doc.Content;
      Word.Find find = rng.Find;
      find.ClearFormatting();
      find.Text = oldText;
      find.Replacement.ClearFormatting();
      find.Replacement.Text = newText;
      find.Forward = true;
      find.Wrap = Word.WdFindWrap.wdFindContinue;
      find.MatchCase = false;
      find.MatchWholeWord = false;
      find.MatchWildcards = false;
      find.Format = false;

      find.Execute(
          FindText: oldText, MatchCase: false, MatchWholeWord: false,
          MatchWildcards: false, MatchSoundsLike: false, MatchAllWordForms: false,
          Forward: true, Wrap: Word.WdFindWrap.wdFindContinue, Format: false,
          ReplaceWith: newText, Replace: Word.WdReplace.wdReplaceAll);
    }
  }
}