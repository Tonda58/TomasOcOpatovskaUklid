/**
 * Aktualizuje viditelnost listů a ochranu pro daný měsíční list.
 * - Zviditelní aktuální list a list "Pracovníci úklidu", ostatní skryje
 * - Ochrání list (editovat může jen vlastník skriptu)
 * - Řádek s dnešním datem (pokud je v tomto měsíci) odemkne pro všechny
 */
function updateVisibilityAndProtection(spreadsheet, sheet, month, year) {
  // Nejprve zajistíme, že aktuální list je viditelný a aktivní,
  // aby při skrývání ostatních listů nikdy nedošlo k situaci
  // „Není možné skrýt všechny listy v dokumentu“.
  sheet.showSheet();
  sheet.activate();

  // Nastavení viditelnosti listů:
  // - aktuálně zpracovávaný list je viditelný
  // - všechny ostatní listy (včetně "Pracovníci úklidu") jsou skryté
  const sheets = spreadsheet.getSheets();
  sheets.forEach(s => {
    if (s.getSheetId() === sheet.getSheetId()) {
      // aktuální měsíc necháme viditelný
      return;
    }
    s.hideSheet();
  });

  // Ochrana listu:
  // - celý list je chráněný (editovat může jen vlastník skriptu)
  // - řádek s dnešním datem (pokud je ve stejném měsíci/roce) je nechráněný
  try {
    let protection;
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (protections && protections.length > 0) {
      protection = protections[0];
    } else {
      protection = sheet.protect();
    }
    protection.setDescription('Ochrana docházky');

    // Určení řádku s dnešním datem (pokud patří do tohoto měsíce)
    const today = new Date();
    const unprotectedRanges = [];
    if (today.getFullYear() === year && today.getMonth() === month - 1) {
      const day = today.getDate();           // 1–31
      const rowIndex = 2 + day;             // data začínají na řádku 3
      const lastRow = sheet.getLastRow();
      if (rowIndex <= lastRow) {
        const editableRange = sheet.getRange(rowIndex, 1, 1, 7); // A:G
        unprotectedRanges.push(editableRange);
      }
    }

    // Nastavit (nebo vynulovat) nechráněné oblasti
    protection.setUnprotectedRanges(unprotectedRanges);

    // Povolit úpravy chráněné části vlastníkovi a adminům ze sloupce D
    const me = Session.getEffectiveUser();
    const adminEmails = getAdminEmails(spreadsheet);

    // Sada povolených emailů (vlastník + admini)
    const allowedEmails = new Set();
    if (me && me.getEmail) {
      allowedEmails.add(me.getEmail());
    }
    adminEmails.forEach(email => allowedEmails.add(email));

    // Přidat editory
    protection.addEditor(me);
    if (adminEmails.length > 0) {
      protection.addEditors(adminEmails);
    }

    // Odstranit všechny ostatní editory (kromě vlastníka a adminů)
    const editors = protection.getEditors();
    editors.forEach(editor => {
      if (!editor.getEmail) return;
      const email = editor.getEmail();
      if (email && !allowedEmails.has(email)) {
        protection.removeEditor(editor);
      }
    });

    Logger.log('✅ Ochrana listu nastavena/aktualizována, řádek s dnešním datem je nechráněný (pokud existuje v tomto měsíci).');
  } catch (error) {
    Logger.log(`⚠️ Chyba při nastavování ochrany listu: ${error}`);
  }

  // Aktualizovat i ochranu listu "Pracovníci úklidu" podle aktuálního seznamu adminů
  protectWorkersSheet(spreadsheet);
}

/**
 * Skript pro kontrolu a vytvoření listu ve formátu mm-yyyy
 * v tabulce "Úkolid OC Opatovská"
 * 
 * DŮLEŽITÉ PRO AUTORIZACI:
 * 1. Otevřete tabulku "Úkolid OC Opatovská" v Google Sheets
 * 2. Otevřete Apps Script editor: Rozšíření → Apps Script
 * 3. Při prvním spuštění klikněte na "Povolit" a autorizujte skript
 * 4. Pokud se zobrazí varování, klikněte na "Pokročilé" → "Přejít na [název projektu] (nebezpečné)"
 */

function checkAndCreateSheet(monthYearStr) {
  // Pokud není zadán měsíc-rok, použijeme aktuální
  if (!monthYearStr || typeof monthYearStr !== 'string') {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    monthYearStr = `${month}-${year}`;
  }
  
  // Parsování mm-yyyy
  const parts = monthYearStr.split('-');
  if (parts.length !== 2) {
    Logger.log('Chyba: Neplatný formát mm-yyyy. Použijte například "02-2026"');
    return false;
  }
  
  const month = parseInt(parts[0]);
  const year = parseInt(parts[1]);
  
  if (month < 1 || month > 12) {
    Logger.log('Chyba: Měsíc musí být mezi 1 a 12');
    return false;
  }
  
  // Otevření tabulky - používáme pouze aktivní tabulku
  let spreadsheet;
  
  try {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      Logger.log('❌ Chyba: Není otevřena žádná tabulka.');
      Logger.log('');
      Logger.log('📋 ŘEŠENÍ:');
      Logger.log('1. Otevřete tabulku "Úkolid OC Opatovská" v Google Sheets');
      Logger.log('2. Otevřete Apps Script editor: Rozšíření → Apps Script');
      Logger.log('3. Ujistěte se, že skript je vázán k této tabulce (ne script.google.com)');
      Logger.log('4. Při prvním spuštění klikněte na "Povolit" a autorizujte skript');
      return false;
    }
    Logger.log(`✅ Použita aktivní tabulka: "${spreadsheet.getName()}"`);
  } catch (error) {
    Logger.log(`❌ Chyba při otevírání tabulky: ${error}`);
    Logger.log('');
    Logger.log('📋 ŘEŠENÍ - AUTORIZACE:');
    Logger.log('1. Otevřete tabulku "Úkolid OC Opatovská" v Google Sheets');
    Logger.log('2. Otevřete Apps Script editor: Rozšíření → Apps Script');
    Logger.log('3. Spusťte funkci main() znovu');
    Logger.log('4. Když se zobrazí dialog "Autorizace požadována", klikněte na "Povolit"');
    Logger.log('5. Pokud se zobrazí varování, klikněte na "Pokročilé" → "Přejít na [název] (nebezpečné)"');
    Logger.log('6. Vyberte svůj Google účet a klikněte na "Povolit"');
    return false;
  }
  
  // Kontrola existence listu
  const sheet = spreadsheet.getSheetByName(monthYearStr);
  if (sheet) {
    Logger.log(`List "${monthYearStr}" již existuje.`);
    // Aktualizace viditelnosti a ochrany (odemčený řádek pro dnešní datum)
    updateVisibilityAndProtection(spreadsheet, sheet, month, year);
    return true;
  }
  
  // Vytvoření nového listu
  const newSheet = spreadsheet.insertSheet(monthYearStr);
  createSheetTemplate(newSheet, spreadsheet, month, year);
  
  Logger.log(`List "${monthYearStr}" byl úspěšně vytvořen.`);
  return true;
}

/**
 * Načte seznam pracovníků z listu "Pracovníci úklidu" ze sloupce A
 * @param {Spreadsheet} spreadsheet - Tabulka
 * @return {Array<string>} - Pole s jmény pracovníků
 */
function getWorkersList(spreadsheet) {
  try {
    const workersSheet = spreadsheet.getSheetByName('Pracovníci úklidu');
    if (!workersSheet) {
      Logger.log('⚠️ Varování: List "Pracovníci úklidu" nebyl nalezen. Dropdown nebude vytvořen.');
      return [];
    }
    
    // Načtení hodnot ze sloupce A od řádku 1
    const lastRow = workersSheet.getLastRow();
    if (lastRow < 1) {
      Logger.log('⚠️ Varování: List "Pracovníci úklidu" je prázdný. Dropdown nebude vytvořen.');
      return [];
    }
    
    const values = workersSheet.getRange(1, 1, lastRow, 1).getValues();
    const workers = [];
    
    // Filtrování prázdných hodnot
    for (let i = 0; i < values.length; i++) {
      const value = values[i][0];
      if (value && String(value).trim() !== '') {
        workers.push(String(value).trim());
      }
    }
    
    if (workers.length === 0) {
      Logger.log('⚠️ Varování: V listu "Pracovníci úklidu" nebyly nalezeny žádné hodnoty. Dropdown nebude vytvořen.');
      return [];
    }
    
    Logger.log(`✅ Načteno ${workers.length} pracovníků z listu "Pracovníci úklidu"`);
    return workers;
  } catch (error) {
    Logger.log(`⚠️ Chyba při načítání pracovníků: ${error}`);
    return [];
  }
}

/**
 * Načte emailové adresy adminů z listu "Pracovníci úklidu" ze sloupce D.
 * Tito lidé mají plný přístup ke všem listům (všechna chráněná data).
 * @param {Spreadsheet} spreadsheet
 * @return {Array<string>} - Pole emailových adres
 */
function getAdminEmails(spreadsheet) {
  try {
    const workersSheet = spreadsheet.getSheetByName('Pracovníci úklidu');
    if (!workersSheet) {
      Logger.log('⚠️ Varování: List "Pracovníci úklidu" nebyl nalezen. Nebudou nastaveni žádní admini.');
      return [];
    }

    const lastRow = workersSheet.getLastRow();
    if (lastRow < 1) {
      return [];
    }

    // Sloupec D = 4
    const values = workersSheet.getRange(1, 4, lastRow, 1).getValues();
    const emails = [];

    for (let i = 0; i < values.length; i++) {
      const value = values[i][0];
      if (value && String(value).trim() !== '') {
        emails.push(String(value).trim());
      }
    }

    Logger.log(`✅ Načteno ${emails.length} admin emailů z listu "Pracovníci úklidu"`);
    return emails;
  } catch (error) {
    Logger.log(`⚠️ Chyba při načítání admin emailů: ${error}`);
    return [];
  }
}

/**
 * Ochrání list "Pracovníci úklidu" tak, aby ho mohli upravovat jen:
 * - vlastník skriptu
 * - emaily uvedené ve sloupci D (admini)
 */
function protectWorkersSheet(spreadsheet) {
  try {
    const workersSheet = spreadsheet.getSheetByName('Pracovníci úklidu');
    if (!workersSheet) {
      Logger.log('⚠️ List "Pracovníci úklidu" nebyl nalezen, nelze nastavit ochranu.');
      return;
    }

    let protection;
    const protections = workersSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (protections && protections.length > 0) {
      protection = protections[0];
    } else {
      protection = workersSheet.protect();
    }
    protection.setDescription('Ochrana listu Pracovníci úklidu');

    // Všechno je chráněné (nemáme nechráněné oblasti)
    protection.setUnprotectedRanges([]);

    const me = Session.getEffectiveUser();
    const adminEmails = getAdminEmails(spreadsheet);

    const allowedEmails = new Set();
    if (me && me.getEmail) {
      allowedEmails.add(me.getEmail());
    }
    adminEmails.forEach(email => allowedEmails.add(email));

    // Přidat editory (vlastník + admini)
    protection.addEditor(me);
    if (adminEmails.length > 0) {
      protection.addEditors(adminEmails);
    }

    // Odstranit ostatní editory
    const editors = protection.getEditors();
    editors.forEach(editor => {
      if (!editor.getEmail) return;
      const email = editor.getEmail();
      if (email && !allowedEmails.has(email)) {
        protection.removeEditor(editor);
      }
    });

    // List "Pracovníci úklidu" necháváme skrytý (slouží jen jako zdroj dat)
    workersSheet.hideSheet();

    Logger.log('✅ Ochrana listu "Pracovníci úklidu" byla nastavena/aktualizována.');
  } catch (error) {
    Logger.log(`⚠️ Chyba při nastavování ochrany listu "Pracovníci úklidu": ${error}`);
  }
}

/**
 * Zkontroluje, zda je datum český státní svátek
 * @param {Date} date - Datum k ověření
 * @return {boolean} - true pokud je státní svátek
 */
function isCzechHoliday(date) {
  const month = date.getMonth() + 1; // getMonth() vrací 0-11
  const day = date.getDate();
  
  // České státní svátky (pevná data)
  const holidays = [
    { month: 1, day: 1 },   // 1. leden - Den obnovy samostatného českého státu
    { month: 5, day: 8 },   // 8. květen - Den vítězství
    { month: 7, day: 5 },   // 5. červenec - Den slovanských věrozvěstů Cyrila a Metoděje
    { month: 7, day: 6 },   // 6. červenec - Den upálení mistra Jana Husa
    { month: 9, day: 28 },  // 28. září - Den české státnosti
    { month: 10, day: 28 }, // 28. říjen - Den vzniku samostatného československého státu
    { month: 11, day: 17 }  // 17. listopad - Den boje za svobodu a demokracii
  ];
  
  return holidays.some(h => h.month === month && h.day === day);
}

function createSheetTemplate(sheet, spreadsheet, month, year) {
  // Žluté pozadí pro hlavičky
  const yellowColor = '#FFFF00';
  const boldFont = true;
  
  // Řádek 1: Nadpis "Docházka pracovníků úklidu M2C" přes sloupce A-G
  sheet.getRange('A1:G1').merge();
  sheet.getRange('A1').setValue('Docházka pracovníků úklidu M2C');
  sheet.getRange('A1').setBackground(yellowColor);
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setHorizontalAlignment('center');
  sheet.getRange('A1').setVerticalAlignment('middle');
  
  // Řádek 2: Hlavičky
  const monthNames = {
    1: 'Leden', 2: 'Únor', 3: 'Březen', 4: 'Duben',
    5: 'Květen', 6: 'Červen', 7: 'Červenec', 8: 'Srpen',
    9: 'Září', 10: 'Říjen', 11: 'Listopad', 12: 'Prosinec'
  };
  
  const headers = [
    [monthNames[month], 'Příchod (čas)', 'Odchod (čas)', 'Jméno', 
     'Příchod (čas)', 'Odchod (čas)', 'Jméno']
  ];
  
  sheet.getRange('A2:G2').setValues(headers);
  sheet.getRange('A2:G2').setBackground(yellowColor);
  sheet.getRange('A2:G2').setFontWeight('bold');
  sheet.getRange('A2:G2').setHorizontalAlignment('center');
  sheet.getRange('A2:G2').setVerticalAlignment('middle');
  
  // Nastavení šířky sloupců
  sheet.setColumnWidth(1, 130); // Sloupec A - Den v týdnu + Datum
  sheet.setColumnWidth(2, 120); // Sloupec B - Příchod
  sheet.setColumnWidth(3, 120); // Sloupec C - Odchod
  sheet.setColumnWidth(4, 150); // Sloupec D - Jméno
  sheet.setColumnWidth(5, 120); // Sloupec E - Příchod
  sheet.setColumnWidth(6, 120); // Sloupec F - Odchod
  sheet.setColumnWidth(7, 150); // Sloupec G - Jméno
  
  // Přidání dat pro všechny dny v měsíci
  const numDays = new Date(year, month, 0).getDate();
  const dayNames = ['Ne', 'Po', 'Út', 'St', 'Čt', 'Pá', 'So'];
  const data = [];
  const holidayRows = []; // Uložení řádků se svátky
  
  for (let day = 1; day <= numDays; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.getDay(); // 0 = neděle, 1 = pondělí, atd.
    const dayName = dayNames[dayOfWeek];
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    // Sloupec A: Den v týdnu s tečkou + Datum ve formátu "Ne. 01.02.2026"
    const formattedDate = `${dayName}. ${dateStr}`;
    data.push([formattedDate]);
    
    // Kontrola, zda je státní svátek
    if (isCzechHoliday(date)) {
      holidayRows.push(day + 2); // +2 protože data začínají na řádku 3
    }
  }
  
  if (data.length > 0) {
    // Nastavení dat do sloupce A
    const dataRange = sheet.getRange(3, 1, data.length, 1);
    dataRange.setValues(data);
    
    // Zarovnání celého sloupce A na střed
    const columnARange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
    columnARange.setHorizontalAlignment('center');
    columnARange.setVerticalAlignment('middle');
    
    // Rozsah pro formátování (celé řádky s daty - sloupce A-G)
    const formatRange = sheet.getRange(3, 1, data.length, 7);
    const rules = [];
    
    // Formátování víkendů (So, Ne) - světle modrá barva
    const lightBlueColor = '#ADD8E6'; // Světle modrá
    
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([formatRange])
      .whenTextContains('So.')
      .setBackground(lightBlueColor)
      .build();
    
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([formatRange])
      .whenTextContains('Ne.')
      .setBackground(lightBlueColor)
      .build();
    
    rules.push(rule1);
    rules.push(rule2);
    
    // Formátování státních svátků - růžová barva
    // Musíme použít přímé formátování, protože podmíněné formátování podle data je složitější
    if (holidayRows.length > 0) {
      const pinkColor = '#FFB6C1'; // Růžová barva
      
      holidayRows.forEach(row => {
        // Formátujeme celý řádek (sloupce A-G)
        const rowRange = sheet.getRange(row, 1, 1, 7);
        rowRange.setBackground(pinkColor);
      });
    }
    
    // Aplikování podmíněných pravidel (víkendy)
    sheet.setConditionalFormatRules(rules);
    
    // Odstranění prázdných řádků (řádky po datech) - PŘED validací
    const lastDataRow = 2 + data.length; // Poslední řádek s daty (řádek 2 je hlavička, data od řádku 3)
    const maxRows = sheet.getMaxRows();
    if (maxRows > lastDataRow) {
      const rowsToDelete = maxRows - lastDataRow;
      try {
        sheet.deleteRows(lastDataRow + 1, rowsToDelete);
        Logger.log(`✅ Odstraněno ${rowsToDelete} prázdných řádků (řádek ${lastDataRow + 1} až ${maxRows})`);
      } catch (error) {
        Logger.log(`⚠️ Chyba při odstraňování řádků: ${error}`);
      }
    }
    
    // Odstranění sloupců H až Z - PŘED validací
    const maxColumns = sheet.getMaxColumns();
    if (maxColumns > 7) { // 7 = sloupec G
      const columnsToDelete = maxColumns - 7;
      try {
        sheet.deleteColumns(8, columnsToDelete); // Začínáme od sloupce H (index 8)
        Logger.log(`✅ Odstraněno ${columnsToDelete} sloupců (H až poslední sloupec)`);
      } catch (error) {
        Logger.log(`⚠️ Chyba při odstraňování sloupců: ${error}`);
      }
    }
    
    // Vytvoření dropdown pro sloupce D a G s pracovníky
    const workers = getWorkersList(spreadsheet);
    if (workers.length > 0) {
      // Vytvoření data validation pro sloupec D (Jméno - první skupina)
      const columnDRange = sheet.getRange(3, 4, data.length, 1); // Sloupec D od řádku 3
      const validationD = SpreadsheetApp.newDataValidation()
        .requireValueInList(workers, true)
        .setAllowInvalid(false)
        .build();
      columnDRange.setDataValidation(validationD);
      
      // Vytvoření data validation pro sloupec G (Jméno - druhá skupina)
      const columnGRange = sheet.getRange(3, 7, data.length, 1); // Sloupec G od řádku 3
      const validationG = SpreadsheetApp.newDataValidation()
        .requireValueInList(workers, true)
        .setAllowInvalid(false)
        .build();
      columnGRange.setDataValidation(validationG);
      
      Logger.log(`✅ Dropdown vytvořen pro sloupce D a G s ${workers.length} pracovníky`);
    }
    
    // Formátování a validace časových sloupců B, C, E, F
    const timeColumnsRangeB = sheet.getRange(3, 2, data.length, 1); // Sloupec B
    const timeColumnsRangeC = sheet.getRange(3, 3, data.length, 1); // Sloupec C
    const timeColumnsRangeE = sheet.getRange(3, 5, data.length, 1); // Sloupec E
    const timeColumnsRangeF = sheet.getRange(3, 6, data.length, 1); // Sloupec F
    
    // Formátování na HH:MM
    timeColumnsRangeB.setNumberFormat('HH:mm');
    timeColumnsRangeC.setNumberFormat('HH:mm');
    timeColumnsRangeE.setNumberFormat('HH:mm');
    timeColumnsRangeF.setNumberFormat('HH:mm');
    
    // Jednoduchá validace formátu času pro všechny časové sloupce (0–1 = 00:00–23:59)
    const timeValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 0.999999999)
      .setAllowInvalid(false)
      .setHelpText('Zadejte čas ve formátu HH:MM (např. 08:00).')
      .build();
    
    timeColumnsRangeB.setDataValidation(timeValidation);
    timeColumnsRangeC.setDataValidation(timeValidation);
    timeColumnsRangeE.setDataValidation(timeValidation);
    timeColumnsRangeF.setDataValidation(timeValidation);
    
    // Podmíněné formátování pro kontrolu podmínek B < C a E < F
    // Červené pozadí pro chybné hodnoty (když B >= C nebo E >= F)
    const redColor = '#FFCCCC'; // Světle červená
    
    // Podmíněné formátování pro sloupec C - červená pokud B >= C (příchod >= odchod)
    const ruleC = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([timeColumnsRangeC])
      .whenFormulaSatisfied('=AND(ISNUMBER(B3), ISNUMBER(C3), B3>=C3)')
      .setBackground(redColor)
      .build();
    
    // Podmíněné formátování pro sloupec B - červená pokud B >= C (pro zvýraznění i příchodu)
    const ruleB = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([timeColumnsRangeB])
      .whenFormulaSatisfied('=AND(ISNUMBER(B3), ISNUMBER(C3), B3>=C3)')
      .setBackground(redColor)
      .build();
    
    // Podmíněné formátování pro sloupec F - červená pokud E >= F (příchod >= odchod)
    const ruleF = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([timeColumnsRangeF])
      .whenFormulaSatisfied('=AND(ISNUMBER(E3), ISNUMBER(F3), E3>=F3)')
      .setBackground(redColor)
      .build();
    
    // Podmíněné formátování pro sloupec E - červená pokud E >= F (pro zvýraznění i příchodu)
    const ruleE = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([timeColumnsRangeE])
      .whenFormulaSatisfied('=AND(ISNUMBER(E3), ISNUMBER(F3), E3>=F3)')
      .setBackground(redColor)
      .build();
    
    // Přidání pravidel do seznamu podmíněného formátování
    const existingRules = sheet.getConditionalFormatRules();
    existingRules.push(ruleB);
    existingRules.push(ruleC);
    existingRules.push(ruleE);
    existingRules.push(ruleF);
    sheet.setConditionalFormatRules(existingRules);
    
    Logger.log('✅ Formátování a validace časových sloupců B, C, E, F dokončeno');
  }
  
  // Aktualizace viditelnosti listů a ochrany pro nově vytvořený list
  updateVisibilityAndProtection(spreadsheet, sheet, month, year);
}

/**
 * Hlavní funkce - spustí se automaticky nebo ručně
 * DŮLEŽITÉ: Skript musí být spuštěn z otevřené tabulky!
 * 
 * Použití:
 * - main() - použije aktuální měsíc a aktivní tabulku
 * - main('02-2026') - použije konkrétní měsíc-rok
 */
function main(monthYearStr) {
  // Pokud je main() volán časovačem, předá se objekt události – ten ignorujeme
  if (typeof monthYearStr !== 'string') {
    monthYearStr = undefined;
  }
  // Pro aktuální měsíc a rok (nebo zadaný parametr):
  checkAndCreateSheet(monthYearStr);
  
  // Příklady použití:
  // main(); // Aktuální měsíc, aktivní tabulka
  // main('02-2026'); // Konkrétní měsíc, aktivní tabulka
}

/**
 * Obnoví dropdowny (sloupce D a G) ve všech měsíčních listech daty z "Pracovníci úklidu"
 */
function refreshWorkersDropdowns(spreadsheet) {
  const workers = getWorkersList(spreadsheet);
  if (workers.length === 0) return;
  
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(workers, true)
    .setAllowInvalid(false)
    .build();
  
  const sheets = spreadsheet.getSheets();
  const monthSheetPattern = /^\d{2}-\d{4}$/; // mm-yyyy
  
  sheets.forEach(s => {
    const name = s.getName();
    if (!monthSheetPattern.test(name)) return;
    
    const lastRow = s.getLastRow();
    if (lastRow < 3) return;

    // Počet datových řádků (od řádku 3 dolů)
    const numRows = lastRow - 2;

    // Dropdown má být pouze ve sloupcích D a G
    const rangeD = s.getRange(3, 4, numRows, 1); // D3:D(lastRow)
    const rangeG = s.getRange(3, 7, numRows, 1); // G3:G(lastRow)
    rangeD.setDataValidation(validation);
    rangeG.setDataValidation(validation);
  });
}

/**
 * onEdit trigger - kontrola, že odchod není dříve než příchod
 * Platí pro všechny listy docházky (sloupce B/C a E/F, řádky od 3 níže)
 * Práce přes půlnoc není povolena.
 * Při úpravě sloupce A v "Pracovníci úklidu" obnoví dropdowny ve všech měsíčních listech.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  
  const ss = SpreadsheetApp.getActive();
  
  // Zvláštní logika pro list "Pracovníci úklidu"
  if (sheet.getName() === 'Pracovníci úklidu') {
    // Změna ve sloupci A -> obnovit dropdowny
    if (col === 1) {
      refreshWorkersDropdowns(ss);
      ss.toast('Dropdowny v listech docházky byly aktualizovány.', 'Aktualizace', 3);
    }
    // Změna ve sloupci D -> obnovit oprávnění (admin emaily)
    if (col === 4) {
      const sheets = ss.getSheets();
      const monthSheetPattern = /^\d{2}-\d{4}$/; // mm-yyyy
      sheets.forEach(s => {
        const name = s.getName();
        if (!monthSheetPattern.test(name)) return;
        const parts = name.split('-');
        const m = parseInt(parts[0], 10);
        const y = parseInt(parts[1], 10);
        if (!m || !y) return;
        updateVisibilityAndProtection(ss, s, m, y);
      });
      ss.toast('Oprávnění pro adminy byla aktualizována.', 'Aktualizace oprávnění', 3);
    }
    // Další logiku pro docházkové listy neprovádíme, takže končíme
    return;
  }
  
  // Ignorujeme hlavičky
  if (row < 3) return;
  
  // Zajímá nás jen B, C, E, F
  if (![2, 3, 5, 6].includes(col)) return;
  
  const b = sheet.getRange(row, 2).getValue();
  const c = sheet.getRange(row, 3).getValue();
  const eVal = sheet.getRange(row, 5).getValue();
  const f = sheet.getRange(row, 6).getValue();
  
  // Kontrola první směny (B < C) – Příchod musí být dřív než Odchod, práce přes půlnoc není povolena
  if ((col === 2 || col === 3) && b instanceof Date && c instanceof Date) {
    const bMinutes = b.getHours() * 60 + b.getMinutes();
    const cMinutes = c.getHours() * 60 + c.getMinutes();
    if (bMinutes >= cMinutes) {
      range.clearContent();
      ss.toast('Odchod (sloupec C) musí být později než příchod (sloupec B). Práce přes půlnoc není povolena.', 'Neplatný čas', 5);
      return;
    }
  }
  
  // Kontrola druhé směny (E < F)
  if ((col === 5 || col === 6) && eVal instanceof Date && f instanceof Date) {
    const eMinutes = eVal.getHours() * 60 + eVal.getMinutes();
    const fMinutes = f.getHours() * 60 + f.getMinutes();
    if (eMinutes >= fMinutes) {
      range.clearContent();
      ss.toast('Odchod (sloupec F) musí být později než příchod (sloupec E). Práce přes půlnoc není povolena.', 'Neplatný czas', 5);
      return;
    }
  }
}

function debugOnEditC7() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const range = sheet.getRange('C7'); // sem si zkuste dát C4/C6 atd.
  const e = {
    range: range,
    source: ss
  };
  onEdit(e);       // zavolá vaši logiku
}