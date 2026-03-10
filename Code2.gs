/**
 * Přepočítá součet odpracovaných hodin všech pracovníků
 * v jednom měsíčním listu (formát názvu „mm-yyyy“).
 *
 * Předpokládané rozložení listu:
 * - A: den + datum (např. „Po. 02.02.2026“)
 * - B: Příchod (čas) 1. pracovník
 * - C: Odchod (čas) 1. pracovník
 * - D: Jméno 1. pracovník
 * - E: Příchod (čas) 2. pracovník
 * - F: Odchod (čas) 2. pracovník
 * - G: Jméno 2. pracovník
 *
 * Výstup:
 * - Na konci tabulky (pod posledním dnem) se nechají 3 prázdné řádky
 * - Pod ně se zapíše souhrn (jeden řádek na pracovníka):
 *   - Sloupec A: jméno (unikátní z D a G)
 *   - Sloupec B: součet hodin daného pracovníka (formát [h]:mm)
 * - No empty lines will appear under the table.
 */
function recalculateMonthSummaryForSheet(sheet) {
  const firstDataRow = 3;

  // Z názvu listu (mm-yyyy) odvodíme měsíc a rok a tím i počet dní
  const name = sheet.getName();
  const parts = name.split('-');
  if (parts.length !== 2) {
    return;
  }
  const month = parseInt(parts[0], 10);
  const year = parseInt(parts[1], 10);
  if (!month || !year) {
    return;
  }

  const numDays = new Date(year, month, 0).getDate();
  const lastDataRow = 2 + numDays; // data jsou na řádcích 3..(2+počet dní)

  const totalsMinutes = {}; // { jmeno: minuty }

  for (let row = firstDataRow; row <= lastDataRow; row++) {
    // Slot 1: B, C, D
    const name1 = sheet.getRange(row, 4).getValue(); // D
    const start1 = sheet.getRange(row, 2).getValue(); // B
    const end1 = sheet.getRange(row, 3).getValue(); // C

    if (name1 && start1 instanceof Date && end1 instanceof Date) {
      const diffMs1 = end1.getTime() - start1.getTime();
      if (diffMs1 > 0) {
        const minutes1 = Math.round(diffMs1 / (1000 * 60));
        const key1 = String(name1).trim();
        totalsMinutes[key1] = (totalsMinutes[key1] || 0) + minutes1;
      }
    }

    // Slot 2: E, F, G
    const name2 = sheet.getRange(row, 7).getValue(); // G
    const start2 = sheet.getRange(row, 5).getValue(); // E
    const end2 = sheet.getRange(row, 6).getValue(); // F

    if (name2 && start2 instanceof Date && end2 instanceof Date) {
      const diffMs2 = end2.getTime() - start2.getTime();
      if (diffMs2 > 0) {
        const minutes2 = Math.round(diffMs2 / (1000 * 60));
        const key2 = String(name2).trim();
        totalsMinutes[key2] = (totalsMinutes[key2] || 0) + minutes2;
      }
    }
  }

  const names = Object.keys(totalsMinutes);
  if (names.length === 0) {
    // Pokud nejsou žádná data, můžeme případný starý souhrn vymazat
    const maxRows = sheet.getMaxRows();
    const summaryStartRowEmpty = lastDataRow + 4;
    if (maxRows >= summaryStartRowEmpty) {
      sheet.getRange(summaryStartRowEmpty, 1, maxRows - summaryStartRowEmpty + 1, 2).clearContent();
    }
    return;
  }

  // Seřadit jména podle abecedy
  names.sort((a, b) => a.localeCompare(b, 'cs'));

  const output = [];
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const minutes = totalsMinutes[name];
    // Převod minut na časovou hodnotu (dny) pro formát [h]:mm
    const timeValue = minutes / (24 * 60);
    output.push([name, timeValue]);
  }

  // Řádek, od kterého začíná souhrn:
  // po posledním dni (lastDataRow) necháme 3 prázdné řádky
  const summaryStartRow = lastDataRow + 4;

  // Vyčistit starý souhrn v A:B od summaryStartRow dolů
  const maxRows = sheet.getMaxRows();
  if (maxRows >= summaryStartRow) {
    sheet.getRange(summaryStartRow, 1, maxRows - summaryStartRow + 1, 2).clearContent();
  }

  // Ujistit se, že máme dostatek řádků pro nový souhrn
  const requiredRows = summaryStartRow + output.length - 1;
  if (maxRows < requiredRows) {
    sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  }

  // Zapsat souhrn do sloupců A a B
  const outRange = sheet.getRange(summaryStartRow, 1, output.length, 2); // A(summaryStartRow):B(...)
  outRange.setValues(output);

  // Sloupec B = součet hodin, formát [h]:mm
  const hoursRange = sheet.getRange(summaryStartRow, 2, output.length, 1); // B(summaryStartRow):B(...)
  hoursRange.setNumberFormat('[h]:mm');
}

/**
 * Spouštěcí funkce pro přepočet součtů v aktuálním listu.
 * Lze volat ručně nebo z onEdit triggeru (instalovaného).
 */
function recalculateMonthSummaryForActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const name = sheet.getName();

  // Pouze listy ve formátu mm-yyyy (měsíční docházka)
  const monthSheetPattern = /^\d{2}-\d{4}$/;
  if (!monthSheetPattern.test(name)) {
    return;
  }

  recalculateMonthSummaryForSheet(sheet);
}

/**
 * Funkce pro instalovatelný onEdit trigger.
 * - Spouští se při jakékoliv změně.
 * - Pokud je změna v měsíčním listu (mm-yyyy), přepočítá souhrn.
 *
 * POZNÁMKA:
 * - V projektu už existuje jednoduchý onEdit v Code.gs,
 *   proto tuto funkci použijte jako INSTALOVATELNÝ trigger:
 *   1. V Apps Script klikněte na „Spouštěče“ (budík vlevo).
 *   2. Přidat spouštěč.
 *   3. Vyberte funkci „onEditMonthSummary“.
 *   4. Událost: „Při úpravě tabulky“.
 */
function onEditMonthSummary(e) {
  // Ochrana proti spuštění bez události (ručně z editoru nebo jiným typem spouštěče)
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const monthSheetPattern = /^\d{2}-\d{4}$/;
  if (!monthSheetPattern.test(name)) {
    return;
  }

  recalculateMonthSummaryForSheet(sheet);
}

