# Instrukce pro autorizaci skriptu

## ⚠️ DŮLEŽITÉ: Autorizace je nutná při prvním spuštění

Pokud vidíte chybu o oprávněních, postupujte podle těchto kroků:

## Krok 1: Otevřete správnou tabulku
1. Otevřete Google Sheets tabulku **"Úkolid OC Opatovská"**
2. Tabulka musí být otevřená v prohlížeči

## Krok 2: Otevřete Apps Script editor Z TABULKY
1. V menu klikněte na **Rozšíření** → **Apps Script**
2. ⚠️ **NEPOUŽÍVEJTE** samostatný Apps Script editor (script.google.com)
3. Skript musí být vázán k tabulce

## Krok 3: Autorizujte skript
1. V Apps Script editoru vyberte funkci `main` z rozbalovacího menu
2. Klikněte na tlačítko ▶️ **Spustit**
3. Objeví se dialog **"Autorizace požadována"**
4. Klikněte na **"Povolit"**

## Krok 4: Dokončete autorizaci
1. Pokud se zobrazí varování o neověřené aplikaci:
   - Klikněte na **"Pokročilé"**
   - Klikněte na **"Přejít na [název vašeho projektu] (nebezpečné)"**
2. Vyberte svůj Google účet
3. Klikněte na **"Povolit"**

## Krok 5: Ověření
1. Po autorizaci by se skript měl spustit bez chyb
2. V logu uvidíte: `✅ Použita aktivní tabulka: "Úkolid OC Opatovská"`

## Problémy?

### Chyba: "You do not have permission"
- Ujistěte se, že jste otevřeli Apps Script editor **z tabulky** (Rozšíření → Apps Script)
- NEPOUŽÍVEJTE script.google.com
- Zkuste znovu autorizovat: v menu klikněte na **Spustit** → **main** → při výzvě klikněte na **"Povolit"**

### Skript není vázán k tabulce
- Zavřete Apps Script editor
- Otevřete tabulku "Úkolid OC Opatovská"
- Znovu otevřete Apps Script editor z tabulky: Rozšíření → Apps Script

### Autorizace nefunguje
- Zkuste vymazat autorizace: [myaccount.google.com/permissions](https://myaccount.google.com/permissions)
- Najděte "Google Apps Script API" a odeberte oprávnění
- Zkuste znovu spustit skript a autorizovat
