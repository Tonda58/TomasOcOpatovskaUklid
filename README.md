# Google Apps Script pro správu docházky

Skript pro automatickou kontrolu a vytváření listů ve formátu `mm-yyyy` v Google Sheets tabulce "Úkolid OC Opatovská".

## Jak použít

⚠️ **DŮLEŽITÉ:** Skript musí být spuštěn z otevřené tabulky v Google Sheets!

1. **Otevřete Google Sheets tabulku "Úkolid OC Opatovská"**
   - Tabulka musí být otevřená v prohlížeči
   - Tabulka musí již existovat (skript ji nevytváří automaticky)

2. **Otevřete Apps Script editor přímo z tabulky:**
   - V menu klikněte na `Rozšíření` → `Apps Script`
   - **NEPOUŽÍVEJTE** samostatný Apps Script editor (script.google.com)

3. **Vložte kód:**
   - Vymažte výchozí kód a vložte obsah souboru `Code.gs`

4. **Spusťte skript:**
   - Vyberte funkci `main` z rozbalovacího menu
   - Klikněte na tlačítko ▶️ Spustit
   - **Při prvním spuštění budete muset autorizovat skript:**
     - Klikněte na **"Povolit"** v dialogu "Autorizace požadována"
     - Pokud se zobrazí varování, klikněte na **"Pokročilé"** → **"Přejít na [název] (nebezpečné)"**
     - Vyberte svůj Google účet a klikněte na **"Povolit"**
   - 📖 **Podrobné instrukce najdete v souboru `AUTORIZACE.md`**

## Funkce

- **`checkAndCreateSheet(monthYearStr)`** - Hlavní funkce, která kontroluje a vytváří list
  - Parametr `monthYearStr`: formát `mm-yyyy` (např. `"02-2026"`)
  - Pokud parametr není zadán, použije aktuální měsíc a rok
  - Používá aktivní tabulku (musí být otevřená)

- **`main(monthYearStr)`** - Spouštěcí funkce

## Co skript dělá

1. Otevře nebo vytvoří tabulku "Úkolid OC Opatovská"
2. Zkontroluje, zda existuje list s názvem ve formátu `mm-yyyy`
3. Pokud list neexistuje, vytvoří ho s:
   - Nadpisem "Docházka pracovníků úklidu M2C"
   - Hlavičkami: Měsíc, Příchod (čas), Odchod (čas), Jméno (2x)
   - Daty pro všechny dny v měsíci ve sloupci A
   - Žlutým pozadím pro hlavičky
   - Nastavenými šířkami sloupců

## Příklad použití

```javascript
// Pro aktuální měsíc a rok (použije aktivní tabulku)
main();

// Pro konkrétní měsíc-rok
main('02-2026');
main('12-2025');
```

## Důležité poznámky

- ⚠️ **Skript MUSÍ být spuštěn z otevřené tabulky** (Rozšíření → Apps Script z tabulky)
- ⚠️ **Při prvním spuštění je nutná autorizace** - viz `AUTORIZACE.md` pro podrobné instrukce
- Pokud tabulka není otevřená, skript zobrazí chybovou zprávu s instrukcemi
- Pokud list s daným názvem již existuje, skript ho nepřepíše

## Řešení problémů s autorizací

Pokud vidíte chybu o oprávněních, postupujte podle instrukcí v souboru **`AUTORIZACE.md`**.
