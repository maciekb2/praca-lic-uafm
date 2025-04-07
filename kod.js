/**
 * @OnlyCurrentDoc
 *
 * Skrypt do zarządzania cudzysłowami w Google Docs.
 * Wersja UNIWERSALNA: Wykrywa różne typy cudzysłowów (", “, ”, „, «, »)
 * i zamienia je na polskie cudzysłowy typograficzne („ ”).
 * Weryfikuje parzystość, zamienia i podświetla treść wewnątrz "..." lub „...”.
 */

// --- Konfiguracja ---
const POLISH_OPENING_QUOTE = '„'; // Cel: Polski otwierający (U+201E)
const POLISH_CLOSING_QUOTE = '”'; // Cel: Polski zamykający (U+201D)

// Znaki cudzysłowów do wyszukania i zamiany
const QUOTES_TO_FIND = ['"', '“', '”', '„', '«', '»'];
// Wyrażenie regularne do znalezienia któregokolwiek z powyższych znaków
// Należy "uescape'ować" znaki specjalne dla regex, jeśli takie by były. Tutaj nie ma potrzeby.
const QUOTE_REGEX_PATTERN = `[${QUOTES_TO_FIND.join('')}]`;

const HIGHLIGHT_COLOR = '#FFFF00'; // Żółty
const ERROR_HIGHLIGHT_COLOR = '#FFDDDD'; // Jasnoczerwony

// --- Menu ---

/**
 * Dodaje niestandardowe menu do interfejsu Google Docs po otwarciu dokumentu.
 * Wersja używająca funkcji opartych na JS indexOf.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Praca Lic. UAFM') // Zmieniono nazwę menu dla jasności
    // Pozycja 1 wywołuje handler, który używa verifyQuoteParityJs
    .addItem('Sprawdź parzystość cudzysłowów', 'runQuoteVerificationUniversal')
    // Pozycja 2 wywołuje handler, który używa replaceQuotesJs (poprzez verifyQuoteParityJs)
    .addItem('Zamień cudzysłowy na polskie', 'runQuoteReplacementUniversal')
    // Pozycja 3 wywołuje NOWY handler dla podświetlania
    .addItem('Podświetl treść (w akapicie)', 'runHighlightQuoteContentJs') // <<<< ZMIENIONO HANDLER
    .addItem('Wyczyść podświetlenia', 'clearHighlights')
    .addSeparator() // Dodaj linię oddzielającą
    .addItem('Wstaw twarde spacje po sierotkach', 'runPreventOrphans') // NOWA OPCJA
    .addItem('usuń podwójne spacje, spacje przed interpunkcją ...', 'runCommonCleanups') // NOWA OPCJA
    .addToUi();
}

// --- Funkcje Uruchamiające (wywoływane z menu) ---

// --- KROK 8: Funkcja zapobiegania sierotkom (1-3 litery) ---

/**
 * Handler dla menu zapobiegania sierotkom (wersja 1-2 litery).
 * Działa TYLKO na GŁÓWNYM TEKŚCIE (BODY).
 */
function runPreventOrphans() {
    Logger.log("[DEBUG runPreventOrphans - 1-2 Letters] --- Start ---");
    const ui = DocumentApp.getUi();
    let doc = null;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) { throw new Error("Nie można uzyskać dostępu do dokumentu."); }
        Logger.log(`[DEBUG runPreventOrphans - 1-2 Letters] Doc OK, ID: ${doc.getId()}`);
    } catch (e) {
        Logger.log(`[ERROR runPreventOrphans - 1-2 Letters] Błąd getActiveDocument: ${e}`);
        ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK);
        return;
    }

    // Zaktualizowany komunikat potwierdzenia
    const confirmation = ui.alert(
        'Potwierdzenie - Twarde Spacje (1-2 litery)', // Zaktualizowano tytuł
        'Czy na pewno chcesz wstawić twarde spacje po wszystkich 1- i 2-literowych słowach w GŁÓWNYM TEKŚCIE dokumentu?\nTa operacja zmodyfikuje tekst i może być trudna do cofnięcia.', // Zaktualizowano wiadomość
        ui.ButtonSet.YES_NO
    );

    if (confirmation === ui.Button.YES) {
        Logger.log("[DEBUG runPreventOrphans - 1-2 Letters] Rozpoczynanie preventOrphansRecursive dla BODY...");
        try {
            const state = { count: 0 }; // Licznik zmian

            const body = doc.getBody();
            if (body) {
                // Wywołujemy funkcję rekurencyjną (która teraz użyje nowego regexa 1-2 litery)
                preventOrphansRecursive(body, state);
                Logger.log(`[DEBUG runPreventOrphans - 1-2 Letters] Przetwarzanie BODY zakończone.`);
            } else {
                 Logger.log("[ERROR runPreventOrphans - 1-2 Letters] Nie udało się pobrać BODY dokumentu.");
                 ui.alert("Błąd", "Nie udało się przetworzyć głównego tekstu dokumentu.", ui.ButtonSet.OK);
                 return;
            }

            Logger.log("[DEBUG runPreventOrphans - 1-2 Letters] Nagłówki, stopki i przypisy zostały zignorowane.");
            Logger.log(`[DEBUG runPreventOrphans - 1-2 Letters] Zakończono. Dokonano około ${state.count} zamian w BODY.`);
            // Zaktualizowany komunikat końcowy
            ui.alert('Twarde Spacje Wstawione (1-2 litery)',
                     `Operacja zakończona. Wstawiono twarde spacje w około ${state.count} miejscach po 1- i 2-literowych słowach w głównym tekście dokumentu.`,
                     ui.ButtonSet.OK);

        } catch (e) {
            Logger.log(`[ERROR runPreventOrphans - 1-2 Letters] Błąd podczas wykonywania preventOrphansRecursive: ${e}`);
            ui.alert("Błąd", "Wystąpił błąd podczas wstawiania twardych spacji.", ui.ButtonSet.OK);
        }
    } else {
        ui.alert('Anulowano', 'Operacja wstawiania twardych spacji została anulowana.', ui.ButtonSet.OK);
    }
    Logger.log("[DEBUG runPreventOrphans - 1-2 Letters] --- Koniec ---");
}

/**
 * Rekurencyjnie przechodzi przez elementy i zamienia spację PO słowach 1- lub 2-literowych na twardą spację (\u00A0).
 * WERSJA KROK 12: Używa pętli regex.exec() i ręcznej zamiany deleteText/insertText.
 * @param {GoogleAppsScript.Document.Element} element Bieżący element.
 * @param {{count: number}} state Obiekt do przekazywania licznika FAKTYCZNYCH zmian.
 */
function preventOrphansRecursive(element, state) {
    if (!element) return;

    let elementType = "UNKNOWN";
    try { elementType = element.getType(); } catch (e) { return; }

    switch (elementType) {
        case DocumentApp.ElementType.TEXT:
            const textElement = element.asText();
            const initialText = textElement.getText(); // Pobierz tekst

            // Sprawdź, czy jest co przetwarzać
            if (initialText && initialText.length > 0) {
                const regex = /\b([a-zA-ZżźćńółęąśŻŹĆŃÓŁĘĄŚ]{1,2})\b(\s)/g; // Zmieniono regex, aby łapał spację jako grupę 2
                const nbsp = '\u00A0'; // Twarda spacja

                // Zbierz wszystkie miejsca (indeksy spacji) do modyfikacji
                const modifications = [];
                let match;
                while ((match = regex.exec(initialText)) !== null) {
                    // match[0] - całe dopasowanie (np. "i ")
                    // match[1] - słowo (np. "i")
                    // match[2] - spacja (np. " ")
                    // match.index - indeks początku całego dopasowania

                    // Indeks spacji do zamiany znajduje się po słowie
                    const wordIndex = match.index;
                    const wordLength = match[1].length;
                    const spaceIndex = wordIndex + wordLength; // Indeks początku spacji
                    const spaceLength = match[2].length; // Długość spacji (zwykle 1)

                    modifications.push({
                        spaceIndex: spaceIndex,
                        spaceLength: spaceLength // Zapisujemy długość spacji na wypadek wielokrotnych spacji
                    });

                    // Zapobieganie nieskończonej pętli dla dopasowań zerowej długości (tu nie powinno wystąpić)
                    if (match[0].length === 0) {
                        regex.lastIndex++;
                    }
                     // Logger.log(`  -> Found match: '${match[0]}' at index ${match.index}. Space index: ${spaceIndex}`); // Opcjonalny log
                }

                // Jeśli znaleziono modyfikacje, wykonaj je OD KOŃCA
                if (modifications.length > 0) {
                    Logger.log(`[DEBUG preventOrphansRecursive] Znaleziono ${modifications.length} spacji do zamiany w elemencie Text (początek: "${initialText.substring(0, 50)}...")`);
                    let successCountInElement = 0;
                    // Iteruj od końca listy modyfikacji, aby nie psuć wcześniejszych indeksów
                    for (let i = modifications.length - 1; i >= 0; i--) {
                        const mod = modifications[i];
                        try {
                            // Usuń oryginalną spację(e)
                            // deleteText(startInclusive, endInclusive)
                            textElement.deleteText(mod.spaceIndex, mod.spaceIndex + mod.spaceLength - 1);
                            // Wstaw twardą spację
                            textElement.insertText(mod.spaceIndex, nbsp);
                            successCountInElement++;
                        } catch (e) {
                            Logger.log(`[ERROR preventOrphansRecursive] Błąd podczas delete/insert na indeksie ${mod.spaceIndex}: ${e}. Stack: ${e.stack}`);
                            // Można rozważyć przerwanie przy błędzie: throw e; lub kontynuować
                        }
                    }
                    // Zaktualizuj globalny licznik o FAKTYCZNIE wykonane zamiany
                    state.count += successCountInElement;
                    Logger.log(`  -> Wykonano ${successCountInElement} zamian w tym elemencie.`);
                }
            }
            break;

        // Rekurencja dla kontenerów (bez zmian)
        case DocumentApp.ElementType.PARAGRAPH:
        case DocumentApp.ElementType.LIST_ITEM:
        case DocumentApp.ElementType.TABLE_CELL:
        case DocumentApp.ElementType.BODY_SECTION:
             if (typeof element.getNumChildren === 'function') {
                 const numChildren = element.getNumChildren();
                 for (let i = 0; i < numChildren; i++) { try { preventOrphansRecursive(element.getChild(i), state); } catch (e) {} }
             }
             break;
        case DocumentApp.ElementType.TABLE:
             try { const numRows = element.getNumRows(); for(let i=0; i < numRows; i++){ const row = element.getRow(i); const numCells = row.getNumCells(); for(let j=0; j < numCells; j++) { try { preventOrphansRecursive(row.getCell(j), state); } catch(e){} } } } catch(e) {}
             break;
        case DocumentApp.ElementType.FOOTNOTE:
             // Ignoruj zawartość przypisów
             break;
        default:
            break;
    }
}

/**
 * Główny handler uruchamiający zestaw typowych poprawek redakcyjnych.
 * Działa TYLKO na GŁÓWNYM TEKŚCIE (BODY).
 */
function runCommonCleanups() {
    Logger.log("[DEBUG runCommonCleanups] --- Start ---");
    const ui = DocumentApp.getUi();
    let doc = null;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) { throw new Error("Nie można uzyskać dostępu do dokumentu."); }
        Logger.log(`[DEBUG runCommonCleanups] Doc OK, ID: ${doc.getId()}`);
    } catch (e) {
        Logger.log(`[ERROR runCommonCleanups] Błąd getActiveDocument: ${e}`);
        ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK);
        return;
    }

    const confirmation = ui.alert(
        'Potwierdzenie - Poprawki Tekstu',
        'Czy na pewno chcesz zastosować zestaw typowych poprawek redakcyjnych?\nZostaną wykonane następujące operacje:\n' +
        '- Zamiana "..." na "…"\n' +
        '- Usunięcie podwójnych spacji\n' +
        '- Usunięcie spacji przed .,;:?!\n' +
        '- Zamiana myślnika w zakresach liczb (np. 10-20 -> 10–20)\n' +
        '- Wstawienie twardej spacji po liczbach przed jednostkami (np. zł, kg, r., w.)\n\n' +
        'Zalecane jest wykonanie kopii zapasowej dokumentu. Tej operacji nie można łatwo cofnąć.',
        ui.ButtonSet.YES_NO
    );

    if (confirmation === ui.Button.YES) {
        Logger.log("[DEBUG runCommonCleanups] Rozpoczynanie poprawek...");
        try {
            const body = doc.getBody();
            if (!body) {
                 Logger.log("[ERROR runCommonCleanups] Nie udało się pobrać BODY dokumentu.");
                 ui.alert("Błąd", "Nie udało się przetworzyć głównego tekstu dokumentu.", ui.ButtonSet.OK);
                 return;
            }

            let results = {
                ellipsis: 0,
                doubleSpaces: 0,
                spaceBeforePunct: 0,
                hyphens: 0,
                nbspAfterNum: 0
            };

            // Wywołaj poszczególne funkcje czyszczące
            results.ellipsis = replaceEllipsis(body);
            results.doubleSpaces = replaceDoubleSpaces(body); // Ta musi być po ellipsis, a przed innymi spacjami
            results.spaceBeforePunct = removeSpaceBeforePunctuation(body);
            results.hyphens = replaceHyphenWithEnDash(body);
            results.nbspAfterNum = insertNbspAfterNumbers(body);

            // Podsumowanie dla użytkownika
            let summary = `Operacja zakończona.\nDokonano przybliżonych zmian w głównym tekście:\n`;
            summary += `- Wielokropki: ${results.ellipsis}\n`;
            summary += `- Podwójne spacje: ${results.doubleSpaces}\n`;
            summary += `- Spacje przed .,;:?!: ${results.spaceBeforePunct}\n`;
            summary += `- Myślniki w zakresach: ${results.hyphens}\n`;
            summary += `- Twarde spacje po liczbach: ${results.nbspAfterNum}`;

            Logger.log(`[DEBUG runCommonCleanups] Podsumowanie zmian: ${JSON.stringify(results)}`);
            ui.alert('Poprawki Zastosowane', summary, ui.ButtonSet.OK);

        } catch (e) {
            Logger.log(`[ERROR runCommonCleanups] Błąd podczas wykonywania poprawek: ${e}\n${e.stack}`);
            ui.alert("Błąd", "Wystąpił błąd podczas stosowania poprawek.", ui.ButtonSet.OK);
        }
    } else {
        ui.alert('Anulowano', 'Operacja wprowadzania poprawek została anulowana.', ui.ButtonSet.OK);
    }
    Logger.log("[DEBUG runCommonCleanups] --- Koniec ---");
}

// --- Funkcje Pomocnicze dla Poprawek ---

/** Zamienia "..." na "…" (U+2026) */
function replaceEllipsis(element) {
    Logger.log("[DEBUG replaceEllipsis] Running...");
    let count = 0;
    try {
        // Używamy replaceText dla całego elementu (np. Body)
        // Trzeba escapować kropki
        count = element.replaceText("\\.{3}", "…");
        Logger.log(`[DEBUG replaceEllipsis] Replaced approx ${count} occurrences.`);
    } catch (e) {
        Logger.log(`[ERROR replaceEllipsis] Error: ${e}`);
    }
    return count;
}

/** Usuwa podwójne (i wielokrotne) spacje, zamieniając je na pojedyncze. */
function replaceDoubleSpaces(element) {
    Logger.log("[DEBUG replaceDoubleSpaces] Running...");
    let text = element.editAsText(); // Potrzebujemy edycji tekstu
    let totalReplacements = 0;
    let replacementsInPass = -1;
    let safetyCounter = 0;
    const MAX_PASSES = 10; // Limit pętli dla bezpieczeństwa

    while(replacementsInPass !== 0 && safetyCounter < MAX_PASSES) {
         try {
              // Zamień dwa spacje na jedną
              replacementsInPass = text.replaceText("  ", " ");
              if (replacementsInPass > 0) {
                  totalReplacements += replacementsInPass;
                  Logger.log(`  -> Pass ${safetyCounter+1}: Replaced ${replacementsInPass} double spaces.`);
              } else {
                   Logger.log(`  -> Pass ${safetyCounter+1}: No more double spaces found.`);
              }
         } catch (e) {
              Logger.log(`  [ERROR] Error during replaceText in replaceDoubleSpaces pass ${safetyCounter+1}: ${e}`);
              break; // Przerwij pętlę w razie błędu
         }
         safetyCounter++;
    }
    if (safetyCounter >= MAX_PASSES) {
         Logger.log("[WARNING replaceDoubleSpaces] Reached maximum passes limit. There might still be multiple spaces if initial count was very high.");
    }
    Logger.log(`[DEBUG replaceDoubleSpaces] Finished. Total replacements: ${totalReplacements}`);
    return totalReplacements;
}

/** Usuwa spacje bezpośrednio przed .,;:?! */
function removeSpaceBeforePunctuation(element) {
    Logger.log("[DEBUG removeSpaceBeforePunctuation] Running...");
    let count = 0;
    try {
        // \s+ : Jedna lub więcej spacji/znaków białych
        // ([.,;?!:]) : Grupa 1: Przechwyć jeden ze znaków interpunkcyjnych
        // Zamień na zawartość grupy 1 (sam znak interpunkcyjny)
        count = element.replaceText("\\s+([.,;?!:])", "$1");
        Logger.log(`[DEBUG removeSpaceBeforePunctuation] Removed spaces before punctuation in approx ${count} places.`);
    } catch (e) {
        Logger.log(`[ERROR removeSpaceBeforePunctuation] Error: ${e}`);
    }
    return count;
}

/** Zamienia łącznik na półpauzę (en dash) w zakresach liczbowych (np. 10-20, 5 - 15) */
function replaceHyphenWithEnDash(element) {
    Logger.log("[DEBUG replaceHyphenWithEnDash] Running...");
    let count = 0;
    try {
        // (\d) : Grupa 1: Cyfra
        // \s*-\s* : Spacja lub jej brak, myślnik, spacja lub jej brak
        // (\d) : Grupa 2: Cyfra
        // Zamień na: Grupa 1, półpauza (U+2013), Grupa 2
        count = element.replaceText("(\\d)\\s*-\\s*(\\d)", "$1–$2"); // Użyj bezpośrednio znaku półpauzy
        Logger.log(`[DEBUG replaceHyphenWithEnDash] Replaced hyphens in approx ${count} number ranges.`);
    } catch (e) {
        Logger.log(`[ERROR replaceHyphenWithEnDash] Error: ${e}`);
    }
    return count;
}

/** Wstawia twardą spację między liczbą a wybranymi jednostkami/skrótami. */
function insertNbspAfterNumbers(element) {
    Logger.log("[DEBUG insertNbspAfterNumbers] Running...");
    let count = 0;
    try {
        // (\d+) : Grupa 1: Jedna lub więcej cyfr
        // \s+ : Jedna lub więcej spacji (którą zamienimy)
        // (zł|PLN|kg|g|m|cm|mm|km|r\.|w\.|%|tys\.|mln|mld|s|min|godz\.) : Grupa 2: Jednostka lub skrót (z escapowanymi kropkami)
        // \b : Granica słowa po jednostce/skrócie
        const regex = /(\d+)\s+(zł|PLN|kg|g|m|cm|mm|km|r\.|w\.|%|tys\.|mln|mld|s|min|godz\.)\b/gi; // Dodano 'i' dla case-insensitive np. PLN
        const replacement = "$1\u00A0$2"; // Grupa 1 + NBSP + Grupa 2

        // Używamy obiektu RegExp bezpośrednio w replaceText
        count = element.replaceText(regex, replacement);
        Logger.log(`[DEBUG insertNbspAfterNumbers] Inserted NBSP in approx ${count} places.`);
    } catch (e) {
        Logger.log(`[ERROR insertNbspAfterNumbers] Error: ${e}`);
    }
    return count;
}

/**
 * Uruchamia proces weryfikacji parzystości WSZYSTKICH zdefiniowanych cudzysłowów.
 * Handler dla menu '1. Sprawdź parzystość (Wszystkie Typy)'
 * WERSJA KROK 5: Wywołuje verifyQuoteParityJs (metoda indexOf)
 */
function runQuoteVerificationUniversal() {
  // Zaczynamy od logowania w handlerze
  Logger.log("[DEBUG runQuoteVerificationUniversal KROK 5] --- Start ---");
  const ui = DocumentApp.getUi();
  let doc = null;
  try {
      doc = DocumentApp.getActiveDocument();
      if (!doc) { Logger.log("[ERROR] DocumentApp.getActiveDocument() zwróciło null!"); ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK); return; }
      Logger.log(`[DEBUG] DocumentApp.getActiveDocument() OK. ID: ${doc.getId()}`);
  } catch (e) { Logger.log(`[ERROR] Błąd DocumentApp.getActiveDocument(): ${e}`); ui.alert("Błąd Krytyczny", `Błąd dostępu do dokumentu: ${e.message}.`, ui.ButtonSet.OK); return; }

  Logger.log("[DEBUG runQuoteVerificationUniversal KROK 5] Czyszczenie podświetleń...");
  try { clearHighlights(); Logger.log("[DEBUG] Czyszczenie OK."); }
  catch (e) { Logger.log(`[ERROR] Błąd clearHighlights(): ${e}. Kontynuuję.`); }

  Logger.log("[DEBUG runQuoteVerificationUniversal KROK 5] Wywołuję NOWĄ funkcję verifyQuoteParityJs...");
  // --- WYWOŁANIE NOWEJ FUNKCJI WERYFIKUJĄCEJ ---
  const result = verifyQuoteParityJs(doc); // <<<< ZMIANA TUTAJ NA NOWĄ FUNKCJĘ
  // -----------------------------------------

  Logger.log(`[DEBUG runQuoteVerificationUniversal KROK 5] Wynik z verifyQuoteParityJs: isEven=${result.isEven}, count=${result.count}`);

  // --- Obsługa wyniku (pozostaje taka sama) ---
  if (result.isEven) {
    ui.alert('Sukces (Metoda JS)', `Znaleziono parzystą liczbę (${result.count}) cudzysłowów (różnych typów).`, ui.ButtonSet.OK);
  } else {
    ui.alert('Błąd Parzystości (Metoda JS)', `Znaleziono nieparzystą liczbę (${result.count}) cudzysłowów (różnych typów). Ostatni znaleziony cudzysłów został podświetlony na czerwono. Popraw dokument.`, ui.ButtonSet.OK);
    const lastRangeEl = result.lastQuoteRangeElement;
    if (lastRangeEl && typeof lastRangeEl.getElement === 'function') {
       const element = lastRangeEl.getElement();
       // Sprawdźmy, czy element istnieje i jest typu TEXT
       if (element && element.getType() === DocumentApp.ElementType.TEXT && typeof lastRangeEl.getStartOffsetInclusive === 'function' && typeof lastRangeEl.getEndOffsetExclusive === 'function') {
         try {
            // Używamy offsetów z RangeElement, który stworzyliśmy w verifyQuoteParityJs
            element.asText().setBackgroundColor(lastRangeEl.getStartOffsetInclusive(), lastRangeEl.getEndOffsetExclusive() -1 , ERROR_HIGHLIGHT_COLOR); // endOffsetExclusive jest o 1 za daleko
             Logger.log(`[DEBUG] Podświetlono błąd parzystości na pozycji ${lastRangeEl.getStartOffsetInclusive()}.`);
         } catch (e) { Logger.log(`[ERROR] Nie udało się podświetlić błędu parzystości: ${e}`); }
       } else { Logger.log("[WARNING] Ostatni element błędu parzystości nie jest typu TEXT lub brakuje metod."); }
    } else { Logger.log("[DEBUG] Nie znaleziono RangeElement dla ostatniego błędu parzystości (lastQuoteRangeElement był null)."); }
  }
   Logger.log("[DEBUG runQuoteVerificationUniversal KROK 5] --- Koniec ---");
}

// --- KROK 6: Implementacja ZAMIANY za pomocą JS indexOf ---

/**
 * Główna funkcja zamiany cudzysłowów używająca metody JS indexOf.
 * Wywołuje zbieranie lokalizacji, sortowanie (uproszczone) i zamianę od końca.
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 */
function replaceQuotesJs(doc) {
    Logger.log("[DEBUG replaceQuotesJs] --- KROK 6: Start zamiany (JS indexOf) ---");

    // 1. Zbierz wszystkie lokalizacje cudzysłowów (używamy tej samej funkcji co w Kroku 5)
    const allLocations = [];
    Logger.log("[DEBUG replaceQuotesJs] Krok 6.1: Zbieranie lokalizacji...");
    let body = null, header = null, footer = null, footnotes = null;
    try { body = doc.getBody(); if(body) collectAllQuoteLocationsJsRecursive(body, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Body: ${e}`);}
    try { header = doc.getHeader(); if(header) collectAllQuoteLocationsJsRecursive(header, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Header: ${e}`);}
    try { footer = doc.getFooter(); if(footer) collectAllQuoteLocationsJsRecursive(footer, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Footer: ${e}`);}
    try { footnotes = doc.getFootnotes(); if(footnotes) footnotes.forEach(fn => { if(fn && fn.getFootnoteContents) collectAllQuoteLocationsJsRecursive(fn.getFootnoteContents(), allLocations); }); } catch(e) { Logger.log(`ERROR getting/processing Footnotes: ${e}`);}
    Logger.log(`[DEBUG replaceQuotesJs] Zebrano ${allLocations.length} lokalizacji.`);

    if (allLocations.length === 0) {
        Logger.log("[DEBUG replaceQuotesJs] Brak cudzysłowów do zamiany. Zakończono.");
        return; // Nie ma nic do roboty
    }

    // 2. Sortowanie Globalne (Uproszczone/Heurystyka)
    // Polegamy na kolejności przetwarzania i sortowaniu wewnątrz elementów.
    // Dla zamiany od końca jest to kluczowe. Jeśli pojawią się problemy z kolejnością,
    // trzeba będzie zaimplementować bardziej zaawansowane sortowanie globalne.
    Logger.log("[DEBUG replaceQuotesJs] Krok 6.2: Sortowanie (uproszczone)...");
     try {
         allLocations.sort((a, b) => {
              // Sortowanie wewnątrz elementu jest ważne
              if (a.element === b.element) { return a.index - b.index; }
              // Między elementami polegamy na kolejności zbierania - brak łatwego sortowania globalnego
              return 0; // Nie zmieniaj kolejności między różnymi elementami (heurystyka)
         });
          Logger.log("[DEBUG replaceQuotesJs] Sortowanie wewnątrz-elementowe zakończone.");
     } catch (e) {
         Logger.log(`[WARNING replaceQuotesJs] Błąd podczas sortowania: ${e}. Kontynuuję, ale kolejność może być nieoptymalna.`);
     }


    // 3. Zamiana od końca listy lokalizacji
    Logger.log("[DEBUG replaceQuotesJs] Krok 6.3: Zamiana od końca listy...");
    let replacementsCount = 0;
    // Iterujemy od ostatniego elementu tablicy (ostatni cudzysłów w dokumencie) do pierwszego
    for (let i = allLocations.length - 1; i >= 0; i--) {
        const location = allLocations[i];

        // Sprawdzenie, czy mamy wszystkie potrzebne informacje
        if (!location || typeof location.index !== 'number' || !location.element || typeof location.element.deleteText !== 'function') {
            Logger.log(`[WARNING replaceQuotesJs] Pomijam nieprawidłową lokalizację na indeksie ${i}.`);
            continue;
        }

        const textElement = location.element; // Element Text, w którym jest cudzysłów
        const index = location.index;         // Pozycja cudzysłowu w tekście elementu
        const originalChar = location.char;   // Jaki cudzysłów tam był (np. ", “, ”)
        const originalCharLength = originalChar.length; // Długość znaku (zwykle 1)

        // Wyznacz poprawny polski cudzysłów na podstawie pozycji w sekwencji
        // Indeks 'i' w tablicy odpowiada (i+1)-temu cudzysłowowi w dokumencie (licząc od 1).
        // Parzysty numer cudzysłowu -> zamykający ('”'), Nieparzysty -> otwierający ('„')
        const replacement = ((i + 1) % 2 === 0) ? POLISH_CLOSING_QUOTE : POLISH_OPENING_QUOTE;

        Logger.log(`[DEBUG replaceQuotesJs] Zamiana [${i+1}/${allLocations.length}] @ index ${index}: '${originalChar}' -> '${replacement}'`);

        try {
             // Dodatkowe sprawdzenie, czy znak w dokumencie się nie zmienił
             const currentChar = textElement.getText().substring(index, index + originalCharLength);
             if (currentChar === originalChar) {
                 // Zamiana: Usuń stary znak, wstaw nowy
                 textElement.deleteText(index, index + originalCharLength - 1); // deleteText(startInclusive, endInclusive)
                 textElement.insertText(index, replacement);
                 replacementsCount++;
             } else {
                  Logger.log(`[WARNING replaceQuotesJs] Oczekiwano '${originalChar}' na pozycji ${index}, ale znaleziono '${currentChar}'. Możliwe przesunięcie tekstu przez wcześniejszą modyfikację (problem z sortowaniem?). Pomijam tę zamianę.`);
             }

        } catch(e) {
            Logger.log(`[ERROR replaceQuotesJs] Błąd podczas zamiany na pozycji ${index}: ${e}`);
            // Można rozważyć przerwanie pętli w razie poważnego błędu
        }
    }
    Logger.log(`[DEBUG replaceQuotesJs] --- KROK 6: Zakończono zamianę. Wykonano ${replacementsCount} zamian. ---`);
}

/**
 * Uruchamia proces zamiany WSZYSTKICH zdefiniowanych cudzysłowów na polskie.
 * WERSJA KROK 6: Wywołuje weryfikację i zamianę metodą JS indexOf.
 */
function runQuoteReplacementUniversal() {
  const ui = DocumentApp.getUi();
   Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] --- Start ---"); // Update log
  let doc = null;
  try {
      doc = DocumentApp.getActiveDocument();
      if (!doc) { Logger.log("[ERROR] DocumentApp.getActiveDocument() zwróciło null!"); ui.alert("Błąd Krytyczny", "Nie udało się uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK); return; }
       Logger.log(`[DEBUG] Dokument aktywny uzyskany. ID: ${doc.getId()}`);
  } catch (e) { Logger.log(`[ERROR] Błąd DocumentApp.getActiveDocument(): ${e}`); ui.alert("Błąd Krytyczny", `Wystąpił błąd podczas dostępu do dokumentu: ${e.message}.`, ui.ButtonSet.OK); return; }

  Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] Czyszczenie podświetleń...");
   try { clearHighlights(); } catch (e) { Logger.log(`[ERROR] Błąd clearHighlights(): ${e}`); }

  Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] Weryfikacja parzystości (JS) przed zamianą...");
  // Używamy funkcji weryfikacji z Kroku 5
  const verificationResult = verifyQuoteParityJs(doc); // <<<< Używa verifyQuoteParityJs
  Logger.log(`[DEBUG runQuoteReplacementUniversal KROK 6] Wynik weryfikacji: isEven=${verificationResult.isEven}, count=${verificationResult.count}`);

  // Obsługa nieparzystej liczby (bez zmian)
  if (!verificationResult.isEven) {
    ui.alert('Anulowano (Metoda JS)', `Nie można zamienić cudzysłowów, ponieważ ich łączna liczba (${verificationResult.count}) jest nieparzysta. Popraw dokument.`, ui.ButtonSet.OK);
     const lastRangeEl = verificationResult.lastQuoteRangeElement;
     if (lastRangeEl /* ... reszta kodu podświetlania błędu ... */) { /* ... */ }
    return;
  }

  // Obsługa zerowej liczby (bez zmian)
  if (verificationResult.count === 0) {
       ui.alert('Informacja (Metoda JS)', `Nie znaleziono żadnych cudzysłowów (${QUOTES_TO_FIND.join(', ')}) do zamiany.`, ui.ButtonSet.OK);
       return;
  }

  // Potwierdzenie od użytkownika (bez zmian)
  const confirmation = ui.alert(
    'Potwierdzenie (Metoda JS)',
    `Znaleziono ${verificationResult.count} cudzysłowów (${QUOTES_TO_FIND.join(', ')}). Czy chcesz zamienić je na polskie („ ”)? Tej operacji nie można cofnąć standardowym Ctrl+Z.`,
    ui.ButtonSet.YES_NO
  );

  // Wywołanie zamiany
  if (confirmation === ui.Button.YES) {
    Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] Rozpoczynanie NOWEJ funkcji replaceQuotesJs...");
    // --- WYWOŁANIE NOWEJ FUNKCJI ZAMIANY ---
    replaceQuotesJs(doc); // <<<< ZMIANA TUTAJ NA NOWĄ FUNKCJĘ
    // -------------------------------------
    Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] Zakończono replaceQuotesJs.");
    ui.alert('Sukces (Metoda JS)', 'Zamiana wszystkich typów cudzysłowów na polskie zakończona.', ui.ButtonSet.OK);
  } else {
    ui.alert('Anulowano (Metoda JS)', 'Operacja zamiany cudzysłowów została anulowana.', ui.ButtonSet.OK);
  }
   Logger.log("[DEBUG runQuoteReplacementUniversal KROK 6] --- Koniec ---");
}

// runHighlightQuoteContent pozostaje bez zmian, bo podświetla tylko "..." i „...”
// clearHighlights pozostaje bez zmian

// --- Logika Podstawowa (Wersje Uniwersalne) ---

// --- NOWE FUNKCJE POMOCNICZE DLA KROKU 5 ---

/**
 * Znajduje indeksy wszystkich szukanych cudzysłowów w danym elemencie Text
 * używając JavaScript `indexOf`.
 * @param {GoogleAppsScript.Document.Text} textElement Element Text do przeszukania.
 * @return {Array<{index: number, char: string, element: GoogleAppsScript.Document.Text}>} Tablica obiektów z indeksem, znalezionym znakiem i elementem.
 */
function findQuoteIndicesInTextElement(textElement) {
    const text = textElement.getText();
    const indices = [];
    if (!text) return indices; // Zwróć pustą tablicę, jeśli tekst jest pusty

    // Iteruj przez każdy szukany znak cudzysłowu
    for (const quoteChar of QUOTES_TO_FIND) {
        let fromIndex = 0;
        let index;
        // Pętla znajdująca wszystkie wystąpienia danego quoteChar
        while ((index = text.indexOf(quoteChar, fromIndex)) !== -1) {
            // Zapisz obiekt z informacjami o znalezisku
            indices.push({ index: index, char: quoteChar, element: textElement });
            fromIndex = index + 1; // Szukaj dalej od następnej pozycji
        }
    }
    // Posortuj znalezione indeksy w ramach tego elementu (ważne dla późniejszej zamiany)
    indices.sort((a, b) => a.index - b.index);
    return indices;
}

/**
 * Rekurencyjnie przechodzi przez elementy dokumentu i zbiera lokalizacje
 * wszystkich szukanych cudzysłowów do tablicy `allLocations`.
 * @param {GoogleAppsScript.Document.Element} element Bieżący element do przetworzenia.
 * @param {Array<{index: number, char: string, element: GoogleAppsScript.Document.Text}>} allLocations Tablica akumulująca wszystkie znalezione lokalizacje.
 */
function collectAllQuoteLocationsJsRecursive(element, allLocations) {
    if (!element) return; // Zabezpieczenie przed null

    let elementType = "UNKNOWN";
    try { elementType = element.getType(); } catch (e) { Logger.log(`Error getting element type: ${e}`); return; }

    switch (elementType) {
        case DocumentApp.ElementType.TEXT:
            // Jeśli to tekst, znajdź w nim cudzysłowy i dodaj do globalnej listy
            const indicesInElement = findQuoteIndicesInTextElement(element.asText());
            if (indicesInElement.length > 0) {
                allLocations.push(...indicesInElement);
                // Logger.log(`Collected ${indicesInElement.length} locations from Text: ${element.asText().getText().substring(0,30)}...`); // Opcjonalny log
            }
            break;

        case DocumentApp.ElementType.PARAGRAPH:
        case DocumentApp.ElementType.LIST_ITEM:
        case DocumentApp.ElementType.TABLE_CELL:
        case DocumentApp.ElementType.BODY_SECTION:
        case DocumentApp.ElementType.HEADER_SECTION:
        case DocumentApp.ElementType.FOOTER_SECTION:
        case DocumentApp.ElementType.FOOTNOTE_SECTION:
            // Jeśli to kontener, przetwórz jego dzieci rekurencyjnie
            if (typeof element.getNumChildren === 'function') {
                const numChildren = element.getNumChildren();
                for (let i = 0; i < numChildren; i++) {
                    try {
                        collectAllQuoteLocationsJsRecursive(element.getChild(i), allLocations);
                    } catch(e) { Logger.log(`Error processing child ${i} of ${elementType}: ${e}`); }
                }
            }
            break;

        case DocumentApp.ElementType.FOOTNOTE:
            // Jeśli to przypis, przetwórz jego zawartość
            if (typeof element.getFootnoteContents === 'function') {
                try {
                    collectAllQuoteLocationsJsRecursive(element.getFootnoteContents(), allLocations);
                } catch (e) { Logger.log(`Error processing footnote contents: ${e}`); }
            }
            break;

        case DocumentApp.ElementType.TABLE:
             // Jeśli to tabela, przetwórz komórki
             try {
                const numRows = element.getNumRows();
                for(let i=0; i < numRows; i++){
                    const row = element.getRow(i);
                    const numCells = row.getNumCells();
                    for(let j=0; j < numCells; j++) {
                       try { collectAllQuoteLocationsJsRecursive(row.getCell(j), allLocations); } catch (cellErr) { Logger.log(`Error processing cell [${i},${j}]: ${cellErr}`);}
                    }
                }
             } catch(e) { Logger.log(`Error processing table: ${e}`); }
             break;

        // Ignoruj inne typy elementów (np. INLINE_IMAGE, HORIZONTAL_RULE)
        default:
            break;
    }
}

/**
 * Weryfikuje parzystość cudzysłowów używając metody JS indexOf.
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 * @return {{count: number, isEven: boolean, lastQuoteRangeElement: GoogleAppsScript.Document.RangeElement | null}} Obiekt z wynikiem.
 */
function verifyQuoteParityJs(doc) {
    Logger.log("[DEBUG verifyQuoteParityJs] --- KROK 5: Start weryfikacji (JS indexOf) ---");
    const allLocations = []; // Tablica na obiekty {index, char, element}

    // Przejdź przez dokument i zbierz lokalizacje
    let body = null, header = null, footer = null, footnotes = null;
    try { body = doc.getBody(); if(body) collectAllQuoteLocationsJsRecursive(body, allLocations); else Logger.log("Body is null"); } catch(e) { Logger.log(`ERROR getting/processing Body: ${e}`);}
    try { header = doc.getHeader(); if(header) collectAllQuoteLocationsJsRecursive(header, allLocations); else Logger.log("Header is null");} catch(e) { Logger.log(`ERROR getting/processing Header: ${e}`);}
    try { footer = doc.getFooter(); if(footer) collectAllQuoteLocationsJsRecursive(footer, allLocations); else Logger.log("Footer is null"); } catch(e) { Logger.log(`ERROR getting/processing Footer: ${e}`);}
    try {
        footnotes = doc.getFootnotes();
        if(footnotes && footnotes.length > 0) {
            Logger.log(`Processing ${footnotes.length} footnotes...`);
            footnotes.forEach((fn, index) => {
                if(fn && typeof fn.getFootnoteContents === 'function') {
                    try {
                        const contents = fn.getFootnoteContents();
                        if (contents) collectAllQuoteLocationsJsRecursive(contents, allLocations);
                        // else Logger.log(`Footnote ${index+1} contents are null.`);
                    } catch (fnErr) { Logger.log(`Error processing footnote ${index+1}: ${fnErr}`); }
                } else { Logger.log(`Invalid footnote object at index ${index}.`);}
            });
        } else { Logger.log("No footnotes found."); }
    } catch(e) { Logger.log(`ERROR getting/processing Footnotes: ${e}`);}

    Logger.log(`[DEBUG verifyQuoteParityJs] Zebrano ${allLocations.length} lokalizacji cudzysłowów.`);

    // --- Globalne sortowanie jest trudne i niekonieczne dla samej weryfikacji parzystości ---
    // Kolejność w allLocations zależy od kolejności przetwarzania elementów.
    // Dla znalezienia ostatniego *fizycznie* w dokumencie, potrzebne byłoby bardziej złożone sortowanie.
    // Na razie dla błędu parzystości użyjemy ostatniego znalezionego w procesie zbierania.

    const count = allLocations.length;
    const isEven = count % 2 === 0;

    // Znajdź ostatnią lokalizację z tablicy (heurystyka)
    let lastLocation = count > 0 ? allLocations[count - 1] : null;
    let lastRangeElement = null; // Potrzebny do podświetlenia błędu

    if (lastLocation) {
        // Spróbuj stworzyć RangeElement dla ostatniego znalezionego znaku, aby móc go podświetlić
        try {
            // Tworzymy Range wskazujący na pojedynczy znak
             lastRangeElement = DocumentApp.getActiveDocument().newRange()
                 .addElement(lastLocation.element, lastLocation.index, lastLocation.index) // Zakres o długości 1 znaku (inclusive start, inclusive end)
                 .build()
                 .getRangeElements()[0]; // Pobierz obiekt RangeElement z Range
             if (lastRangeElement) {
                Logger.log(`[DEBUG verifyQuoteParityJs] Utworzono RangeElement dla ostatniej lokalizacji: index ${lastLocation.index} w elemencie ${lastLocation.element.getParent().getType()}.`);
             } else {
                 Logger.log(`[WARNING verifyQuoteParityJs] Nie udało się uzyskać RangeElement z zbudowanego Range.`);
             }
        } catch (e) {
            Logger.log(`[ERROR verifyQuoteParityJs] Nie udało się utworzyć RangeElement dla ostatniej lokalizacji: ${e}`);
            lastRangeElement = null;
        }
    }

    Logger.log(`[DEBUG verifyQuoteParityJs] --- KROK 5: Zakończono. Zwracam count=${count}, isEven=${isEven} ---`);
    return { count: count, isEven: isEven, lastQuoteRangeElement: lastRangeElement };
}

// --- KONIEC NOWYCH FUNKCJI POMOCNICZYCH ---


// --- KROK 7: Implementacja PODŚWIETLANIA za pomocą JS indexOf ---

/**
 * Podświetla treść pomiędzy parami cudzysłowów, używając lokalizacji znalezionych przez JS indexOf.
 * OGRANICZENIE: Podświetla tylko pary znajdujące się w tym samym elemencie Text.
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 * @return {number} Liczba pomyślnie podświetlonych fragmentów.
 */
function highlightContentBetweenQuotesJs(doc) {
    Logger.log("[DEBUG highlightContentJs] --- KROK 7: Start podświetlania (JS indexOf) ---");

    // 1. Zbierz wszystkie lokalizacje cudzysłowów
    const allLocations = [];
    Logger.log("[DEBUG highlightContentJs] Krok 7.1: Zbieranie lokalizacji...");
    // Używamy tej samej funkcji pomocniczej co poprzednio
    let body = null, header = null, footer = null, footnotes = null;
    try { body = doc.getBody(); if(body) collectAllQuoteLocationsJsRecursive(body, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Body: ${e}`);}
    try { header = doc.getHeader(); if(header) collectAllQuoteLocationsJsRecursive(header, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Header: ${e}`);}
    try { footer = doc.getFooter(); if(footer) collectAllQuoteLocationsJsRecursive(footer, allLocations); } catch(e) { Logger.log(`ERROR getting/processing Footer: ${e}`);}
    try { footnotes = doc.getFootnotes(); if(footnotes) footnotes.forEach(fn => { if(fn && fn.getFootnoteContents) collectAllQuoteLocationsJsRecursive(fn.getFootnoteContents(), allLocations); }); } catch(e) { Logger.log(`ERROR getting/processing Footnotes: ${e}`);}
    Logger.log(`[DEBUG highlightContentJs] Zebrano ${allLocations.length} lokalizacji.`);

    if (allLocations.length < 2) {
        Logger.log("[DEBUG highlightContentJs] Znaleziono mniej niż 2 cudzysłowy, brak par do podświetlenia.");
        return 0; // Zwróć 0 podświetlonych
    }

    // 2. Sprawdzenie parzystości (dla pewności logiki par)
    if (allLocations.length % 2 !== 0) {
        Logger.log(`[WARNING highlightContentJs] Znaleziono nieparzystą liczbę (${allLocations.length}) cudzysłowów. Podświetlanie może być niekompletne.`);
        // Kontynuujemy, próbując podświetlić znalezione pary
    }

    // 3. Sortowanie (Uproszczone - kluczowe, aby pary były obok siebie w tablicy, jeśli są blisko w dokumencie)
    Logger.log("[DEBUG highlightContentJs] Krok 7.2: Sortowanie (uproszczone)...");
     try {
         // Sortowanie wewnątrz elementu jest najważniejsze dla znajdowania par w tym samym elemencie
         allLocations.sort((a, b) => {
              if (a.element === b.element) { return a.index - b.index; }
              // Między elementami polegamy na kolejności zbierania
              return 0;
         });
          Logger.log("[DEBUG highlightContentJs] Sortowanie wewnątrz-elementowe zakończone.");
     } catch (e) { Logger.log(`[WARNING highlightContentJs] Błąd podczas sortowania: ${e}.`); }

    // 4. Iteracja przez PARY i podświetlanie (tylko w ramach tego samego elementu Text)
    Logger.log("[DEBUG highlightContentJs] Krok 7.3: Iteracja przez pary i podświetlanie...");
    let highlightedCount = 0;
    // Iterujemy co drugi element, traktując i oraz i+1 jako potencjalną parę
    for (let i = 0; i < allLocations.length - 1; i += 2) {
        const locOpen = allLocations[i];    // Potencjalny cudzysłów otwierający
        const locClose = allLocations[i + 1]; // Potencjalny cudzysłów zamykający

        // Sprawdź, czy obie lokalizacje istnieją i czy są w tym samym elemencie Text
        if (locOpen && locClose && locOpen.element === locClose.element && locOpen.element.getType() === DocumentApp.ElementType.TEXT) {
            const textElement = locOpen.element;
            const startIndex = locOpen.index + locOpen.char.length; // Pozycja po znaku otwierającym
            const endIndex = locClose.index - 1;                   // Pozycja przed znakiem zamykającym

            Logger.log(`[DEBUG highlightContentJs] Para [${i+1}, ${i+2}]. Element: ${textElement.getParent().getType()}, Zakres podświetlenia: [${startIndex}-${endIndex}]`);

            // Sprawdź, czy zakres jest poprawny (koniec >= początek)
            if (endIndex >= startIndex) {
                try {
                    // Podświetl tekst między cudzysłowami
                    textElement.setBackgroundColor(startIndex, endIndex, HIGHLIGHT_COLOR);
                    highlightedCount++;
                    Logger.log(`  -> Podświetlono!`);
                } catch (e) {
                    Logger.log(`  [ERROR highlightContentJs] Błąd podczas setBackgroundColor dla zakresu [${startIndex}-${endIndex}]: ${e}`);
                }
            } else {
                 Logger.log(`  -> Pusta treść między cudzysłowami [${startIndex}-${endIndex}]. Pomijam podświetlanie.`);
            }
        } else {
             // Loguj tylko co 10 nieudaną parę, aby uniknąć nadmiaru logów
             if (i % 20 === 0) { // Co 10 par = co 20 iteracji
                  let reason = "nieznany";
                  if (!locOpen || !locClose) reason = "brakująca lokalizacja";
                  else if (locOpen.element !== locClose.element) reason = "różne elementy Text";
                  else if (locOpen.element.getType() !== DocumentApp.ElementType.TEXT) reason = "element nie jest typu TEXT";
                  Logger.log(`[DEBUG highlightContentJs] Pomijam parę [${i+1}, ${i+2}], bo ${reason}. Podświetlanie między elementami nie jest wspierane.`);
             }
        }
    }

    Logger.log(`[DEBUG highlightContentJs] --- KROK 7: Zakończono podświetlanie. Podświetlono ${highlightedCount} fragmentów. ---`);
    return highlightedCount; // Zwróć liczbę podświetlonych fragmentów
}

/**
 * Nowy handler dla podświetlania treści cytatów, używający metody JS indexOf.
 */
function runHighlightQuoteContentJs() {
    Logger.log("[DEBUG runHighlightQuoteContentJs] --- KROK 7: Start handlera podświetlania (JS indexOf) ---");
    const ui = DocumentApp.getUi();
    let doc = null;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu."); // Rzuć błąd, jeśli doc jest null
        Logger.log(`[DEBUG runHighlightQuoteContentJs] Doc OK, ID: ${doc.getId()}`);
    }
    catch (e) { Logger.log(`[ERROR runHighlightQuoteContentJs] Błąd getActiveDocument: ${e}`); ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK); return; }

    Logger.log("[DEBUG runHighlightQuoteContentJs] Czyszczenie podświetleń...");
    try { clearHighlights(); Logger.log("[DEBUG runHighlightQuoteContentJs] Czyszczenie OK."); }
    catch (e) { Logger.log(`[ERROR runHighlightQuoteContentJs] Błąd clearHighlights: ${e}`); }

    Logger.log("[DEBUG runHighlightQuoteContentJs] Wywołuję NOWĄ funkcję highlightContentBetweenQuotesJs...");
    let highlightedCount = 0;
    try {
         // --- WYWOŁANIE NOWEJ FUNKCJI PODŚWIETLANIA ---
         highlightedCount = highlightContentBetweenQuotesJs(doc);
         // -------------------------------------------
         Logger.log(`[DEBUG runHighlightQuoteContentJs] Zakończono podświetlanie, wynik (liczba fragmentów): ${highlightedCount}`);

         // Informacja dla użytkownika
         if (highlightedCount > 0) {
             ui.alert('Podświetlono Treść (Metoda JS)', `Podświetlono treść wewnątrz ${highlightedCount} par cudzysłowów (tylko pary wewnątrz tego samego akapitu).`, ui.ButtonSet.OK);
         } else {
             ui.alert('Nie Znaleziono (Metoda JS)', 'Nie znaleziono par cudzysłowów wewnątrz tego samego akapitu do podświetlenia lub wystąpiły błędy.', ui.ButtonSet.OK);
         }

    } catch (e) {
         Logger.log(`[ERROR runHighlightQuoteContentJs] Błąd podczas highlightContentBetweenQuotesJs: ${e}`);
         ui.alert("Błąd", "Wystąpił błąd podczas próby podświetlenia cytatów.", ui.ButtonSet.OK);
    }
    Logger.log("[DEBUG runHighlightQuoteContentJs] --- Koniec ---");
}

/**
 * Usuwa podświetlenie tła (ustawia na null) z całego dokumentu.
 */
function clearHighlights() {
    // ... (kod tej funkcji pozostaje bez zmian z poprzedniej wersji) ...
     const doc = DocumentApp.getActiveDocument();
      function clearElementHighlight(element) {
         if (!element) return;
         if (element.getType() === DocumentApp.ElementType.TEXT) {
           try { if (element.asText().getText().length > 0) element.asText().setBackgroundColor(null); } catch (e) { Logger.log(`Info: Nie można wyczyścić tła dla Text: ${e}`); }
         } else if (typeof element.getNumChildren === 'function') {
           const numChildren = element.getNumChildren(); for (let i = 0; i < numChildren; i++) clearElementHighlight(element.getChild(i));
         } else if (element.getType() === DocumentApp.ElementType.FOOTNOTE && typeof element.getFootnoteContents === 'function') { clearElementHighlight(element.getFootnoteContents());
         } else if (element.getType() === DocumentApp.ElementType.TABLE) { const numRows = element.getNumRows(); for(let i=0; i < numRows; i++){ const row = element.getRow(i); const numCells = row.getNumCells(); for(let j=0; j < numCells; j++) clearElementHighlight(row.getCell(j)); }
         } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) { if (typeof element.getNumChildren === 'function') { const numChildren = element.getNumChildren(); for (let i = 0; i < numChildren; i++) clearElementHighlight(element.getChild(i)); } }
      }
      clearElementHighlight(doc.getBody());
      const header = doc.getHeader(); if (header) clearElementHighlight(header);
      const footer = doc.getFooter(); if (footer) clearElementHighlight(footer);
      const footnotes = doc.getFootnotes(); if (footnotes) footnotes.forEach(footnote => { if (footnote && typeof footnote.getFootnoteContents === 'function') clearElementHighlight(footnote.getFootnoteContents()); });
      Logger.log("Wyczyszczono podświetlenia.");
}
