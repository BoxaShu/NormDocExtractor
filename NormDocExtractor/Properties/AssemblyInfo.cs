// Copyright (c) 2026 Шульжицкий Владимир
//
// This file is part of NormExtract.
//
// NormExtract is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// NormExtract is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with NormExtract.  If not, see <https://www.gnu.org/licenses/>.

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// ─── Общие сведения о сборке ────────────────────────────────
[assembly: AssemblyTitle("NormDocExtract")]
[assembly: AssemblyDescription(
    "Извлечение списка нормативных документов (ГОСТ, СНиП, СП, " +
    "СанПиН и др.) из документа Word с возможностью проверки " +
    "статуса и замены устаревших номеров")]
[assembly: AssemblyConfiguration("Release")]
[assembly: AssemblyCompany("Шульжицкий Владимир")]
[assembly: AssemblyProduct("NormDocExtract")]
[assembly: AssemblyCopyright(
    "Copyright © 2026 Шульжицкий Владимир. " +
    "Licensed under the GNU General Public License v3.0.")]
[assembly: AssemblyTrademark("")]
//[assembly: AssemblyCulture("Ru-ru")]

// ─── Настройки COM-взаимодействия ───────────────────────────
[assembly: ComVisible(false)]

// Уникальный GUID для typelib, если сборка будет видна через COM
[assembly: Guid("674490d2-05bb-4a70-a082-77f89a55ace3")]

// ─── Версия ─────────────────────────────────────────────────
// Major.Minor.Build.Revision
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.14")]
[assembly: AssemblyInformationalVersion("1.0.0.14")]

// ─── Настройки безопасности для VSTO ────────────────────────
[assembly: SecurityRules(SecurityRuleSet.Level1)]

// ─── Видимость внутренних типов для тестов ──────────────────
// Раскомментируйте, если добавите unit-тесты
// [assembly: InternalsVisibleTo("NormExtract.Tests")]
