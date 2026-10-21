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
using System.Text.RegularExpressions;

namespace NormDocExtractor
{
  public static class NormativeDocExtractor
  {
    private const string Pattern =
        "(" +
        "(ГОСТ( Р)?( (ИСО|ISO|ЕН|EN|МЭК|IEC)(/\\S+)?)? \\d[\\d\\.\\-/]*([\\- ]\\d{2,4})?" +
        "|СНиП [IVX\\d][IVX\\d\\.\\-]*" +
        "|СП \\d[\\d\\.\\-]+" +
        "|СанПиН( Р)? [\\d\\.\\-]+" +
        "|ТУ [\\dА-Я][\\dА-Я\\-\\./]+" +
        "|РД [\\dА-Я][\\dА-Я\\.\\-/]+" +
        "|ВСН [\\dА-Я][\\dА-Я\\.\\-/]+" +
        "|МДС [\\dА-Я][\\dА-Я\\.\\-/]+" +
        "|СТО [\\dА-Я][\\dА-Я\\.\\-/]+" +
        "|НПБ [\\dА-Я][\\dА-Я\\-]+" +
        "|ПБ [\\dА-Я][\\dА-Я\\-]+" +
        "|ПУЭ" +
        "|ФЗ № ?\\d+\\-ФЗ" +
        "|№ ?\\d+\\-ФЗ" +
        ")\\*?)";

    private static readonly Regex Rx = new Regex(Pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public sealed class Result
    {
      public int TotalMentions { get; }
      public IReadOnlyDictionary<string, int> UniqueDocs { get; }
      public Result(int total, IReadOnlyDictionary<string, int> unique)
      { TotalMentions = total; UniqueDocs = unique; }
    }

    public static Result Extract(string rawText)
    {
      string text = rawText
          .Replace("\u00A0", " ")
          .Replace("\u00AD", "")
          .Replace("\r", " ")
          .Replace("\n", " ")
          .Replace("\t", " ")
          .Replace("\v", " ");
      text = Regex.Replace(text, " +", " ");

      var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
      int total = 0;

      foreach (Match m in Rx.Matches(text))
      {
        string s = m.Value.Trim().TrimEnd('-', '.', '/', ' ', ',', ';', ':');
        if (s.Length == 0) continue;
        total++;
        dict.TryGetValue(s, out int c);
        dict[s] = c + 1;
      }

      return new Result(total, dict);
    }
  }
}