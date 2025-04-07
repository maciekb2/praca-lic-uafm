/**
 * @OnlyCurrentDoc
 *
 * Skrypt do zarządzania cudzysłowami i wykonywania poprawek redakcyjnych w Google Docs.
 * Wyszukuje różne typy cudzysłowów i zamienia je na polskie typograficzne („ ”).
 * Oferuje również funkcje prewencji sierotek i typowych poprawek tekstu.
 */

// --- Konfiguracja ---
const POLISH_OPENING_QUOTE = '„'; // Polski cudzysłów otwierający (U+201E)
const POLISH_CLOSING_QUOTE = '”'; // Polski cudzysłów zamykający (U+201D)

// Znaki cudzysłowów do wyszukania i zamiany
const QUOTES_TO_FIND = ['"', '“', '”', '„', '«', '»'];

const HIGHLIGHT_COLOR = '#FFFF00'; // Żółty do podświetlania treści cytatów
const ERROR_HIGHLIGHT_COLOR = '#FFDDDD'; // Jasnoczerwony do błędów parzystości

// --- Menu ---

/**
 * Dodaje niestandardowe menu 'Praca Lic. UAFM' do interfejsu Google Docs po otwarciu dokumentu.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Praca Lic. UAFM')
    .addItem('Sprawdź parzystość cudzysłowów', 'runQuoteVerificationUniversal')
    .addItem('Zamień cudzysłowy na polskie', 'runQuoteReplacementUniversal')
    .addItem('Podświetl treść (w akapicie)', 'runHighlightQuoteContentJs')
    .addItem('Wyczyść podświetlenia', 'clearHighlights')
    .addSeparator()
    .addItem('Wstaw twarde spacje po sierotkach', 'runPreventOrphans')
    .addItem('Typowe Poprawki Tekstu', 'runCommonCleanups') // Zmieniono nazwę dla zwięzłości
    .addToUi();
}

// --- Funkcje Uruchamiające (Handlers) ---

/**
 * Handler dla menu wstawiania twardych spacji po 1- i 2-literowych słowach.
 * Działa TYLKO na GŁÓWNYM TEKŚCIE (BODY).
 */
function runPreventOrphans() {
    Logger.log("runPreventOrphans: Start");
    const ui = DocumentApp.getUi();
    let doc;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu.");
    } catch (e) {
        Logger.log(`runPreventOrphans: Błąd getActiveDocument: ${e}`);
        ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK);
        return;
    }

    const confirmation = ui.alert(
        'Potwierdzenie - Twarde Spacje (1-2 litery)',
        'Czy na pewno chcesz wstawić twarde spacje po wszystkich 1- i 2-literowych słowach w GŁÓWNYM TEKŚCIE dokumentu?\nTa operacja zmodyfikuje tekst i może być trudna do cofnięcia.',
        ui.ButtonSet.YES_NO
    );

    if (confirmation === ui.Button.YES) {
        try {
            const state = { count: 0 }; // Licznik zmian
            const body = doc.getBody();
            if (body) {
                preventOrphansRecursive(body, state);
            } else {
                 Logger.log("runPreventOrphans: Nie udało się pobrać BODY dokumentu.");
                 ui.alert("Błąd", "Nie udało się przetworzyć głównego tekstu dokumentu.", ui.ButtonSet.OK);
                 return;
            }
            Logger.log(`runPreventOrphans: Zakończono. Dokonano około ${state.count} zamian.`);
            ui.alert('Twarde Spacje Wstawione (1-2 litery)',
                     `Operacja zakończona. Wstawiono twarde spacje w około ${state.count} miejscach po 1- i 2-literowych słowach w głównym tekście dokumentu.`,
                     ui.ButtonSet.OK);
        } catch (e) {
            Logger.log(`runPreventOrphans: Błąd podczas wykonywania preventOrphansRecursive: ${e}`);
            ui.alert("Błąd", "Wystąpił błąd podczas wstawiania twardych spacji.", ui.ButtonSet.OK);
        }
    } else {
        ui.alert('Anulowano', 'Operacja wstawiania twardych spacji została anulowana.', ui.ButtonSet.OK);
    }
    Logger.log("runPreventOrphans: Koniec");
}

/**
 * Rekurencyjnie przechodzi przez elementy i zamienia spację PO słowach 1- lub 2-literowych na twardą spację (\u00A0).
 * @param {GoogleAppsScript.Document.Element} element Bieżący element.
 * @param {{count: number}} state Obiekt do przekazywania licznika zmian.
 */
function preventOrphansRecursive(element, state) {
    if (!element) return;

    let elementType;
    try { elementType = element.getType(); } catch (e) { return; } // Ignoruj błędy pobierania typu

    switch (elementType) {
        case DocumentApp.ElementType.TEXT:
            const textElement = element.asText();
            const initialText = textElement.getText();
            if (initialText && initialText.length > 0) {
                // Regex: Znajdź 1-2 litery na granicy słowa, po których jest spacja (grupa 2)
                const regex = /\b([a-zA-ZżźćńółęąśŻŹĆŃÓŁĘĄŚ]{1,2})\b(\s)/g;
                const nbsp = '\u00A0'; // Twarda spacja

                const modifications = [];
                let match;
                while ((match = regex.exec(initialText)) !== null) {
                    const wordIndex = match.index;
                    const wordLength = match[1].length;
                    const spaceIndex = wordIndex + wordLength;
                    const spaceLength = match[2].length;
                    modifications.push({ spaceIndex: spaceIndex, spaceLength: spaceLength });
                    if (match[0].length === 0) regex.lastIndex++; // Zapobieganie nieskończonej pętli
                }

                // Modyfikacje od końca, aby nie psuć wcześniejszych indeksów
                if (modifications.length > 0) {
                    let successCountInElement = 0;
                    for (let i = modifications.length - 1; i >= 0; i--) {
                        const mod = modifications[i];
                        try {
                            textElement.deleteText(mod.spaceIndex, mod.spaceIndex + mod.spaceLength - 1);
                            textElement.insertText(mod.spaceIndex, nbsp);
                            successCountInElement++;
                        } catch (e) {
                            Logger.log(`preventOrphansRecursive: Błąd delete/insert na indeksie ${mod.spaceIndex}: ${e}`);
                        }
                    }
                    state.count += successCountInElement;
                }
            }
            break;

        // Rekurencja dla kontenerów
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
             try {
                 const numRows = element.getNumRows();
                 for (let i = 0; i < numRows; i++) {
                     const row = element.getRow(i);
                     const numCells = row.getNumCells();
                     for (let j = 0; j < numCells; j++) { try { preventOrphansRecursive(row.getCell(j), state); } catch(e) {} }
                 }
             } catch(e) {}
             break;
        case DocumentApp.ElementType.FOOTNOTE:
             // Celowo ignoruje zawartość przypisów dla tej funkcji
             break;
        default:
            break;
    }
}

/**
 * Handler dla menu uruchamiającego zestaw typowych poprawek redakcyjnych.
 * Działa TYLKO na GŁÓWNYM TEKŚCIE (BODY).
 */
function runCommonCleanups() {
    Logger.log("runCommonCleanups: Start");
    const ui = DocumentApp.getUi();
    let doc;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu.");
    } catch (e) {
        Logger.log(`runCommonCleanups: Błąd getActiveDocument: ${e}`);
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
        try {
            const body = doc.getBody();
            if (!body) {
                 Logger.log("runCommonCleanups: Nie udało się pobrać BODY dokumentu.");
                 ui.alert("Błąd", "Nie udało się przetworzyć głównego tekstu dokumentu.", ui.ButtonSet.OK);
                 return;
            }

            let results = { ellipsis: 0, doubleSpaces: 0, spaceBeforePunct: 0, hyphens: 0, nbspAfterNum: 0 };

            // Wywołaj poszczególne funkcje czyszczące
            results.ellipsis = replaceEllipsis(body);
            results.doubleSpaces = replaceDoubleSpaces(body); // Ważna kolejność: po ellipsis, przed innymi spacjami
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

            Logger.log(`runCommonCleanups: Podsumowanie zmian: ${JSON.stringify(results)}`);
            ui.alert('Poprawki Zastosowane', summary, ui.ButtonSet.OK);

        } catch (e) {
            Logger.log(`runCommonCleanups: Błąd podczas wykonywania poprawek: ${e}\n${e.stack}`);
            ui.alert("Błąd", "Wystąpił błąd podczas stosowania poprawek.", ui.ButtonSet.OK);
        }
    } else {
        ui.alert('Anulowano', 'Operacja wprowadzania poprawek została anulowana.', ui.ButtonSet.OK);
    }
    Logger.log("runCommonCleanups: Koniec");
}

// --- Funkcje Pomocnicze dla Poprawek ---

/** Zamienia trzy kropki "..." na pojedynczy znak wielokropka "…" (U+2026). */
function replaceEllipsis(element) {
    let count = 0;
    try {
        count = element.replaceText("\\.{3}", "…"); // Escapowane kropki
    } catch (e) { Logger.log(`replaceEllipsis Error: ${e}`); }
    return count;
}

/** Usuwa podwójne (i wielokrotne) spacje, zamieniając je na pojedyncze. */
function replaceDoubleSpaces(element) {
    // Potrzebna edycja jako tekst i pętla, bo replaceText może nie złapać wszystkiego za jednym razem
    let text = element.editAsText();
    let totalReplacements = 0;
    let replacementsInPass = -1;
    let safetyCounter = 0;
    const MAX_PASSES = 10; // Zabezpieczenie przed nieskończoną pętlą

    while(replacementsInPass !== 0 && safetyCounter < MAX_PASSES) {
         try {
              replacementsInPass = text.replaceText("  ", " ");
              if (replacementsInPass > 0) totalReplacements += replacementsInPass;
         } catch (e) {
              Logger.log(`replaceDoubleSpaces Error pass ${safetyCounter+1}: ${e}`);
              break;
         }
         safetyCounter++;
    }
    if (safetyCounter >= MAX_PASSES) {
         Logger.log("replaceDoubleSpaces: Osiągnięto limit przejść.");
    }
    return totalReplacements;
}

/** Usuwa spacje bezpośrednio przed wybranymi znakami interpunkcyjnymi (.,;:?!). */
function removeSpaceBeforePunctuation(element) {
    let count = 0;
    try {
        // Regex: \s+([.,;?!:]) -> $1
        // Znajdź jedną lub więcej spacji ( \s+ ) przed jednym ze znaków w grupie ( [.,;?!:] )
        // i zastąp tylko znakiem interpunkcyjnym ( $1 - zawartość pierwszej grupy ).
        count = element.replaceText("\\s+([.,;?!:])", "$1");
    } catch (e) { Logger.log(`removeSpaceBeforePunctuation Error: ${e}`); }
    return count;
}

/** Zamienia łącznik (-) na półpauzę (–, U+2013) w zakresach liczbowych (np. 10-20, 5 - 15). */
function replaceHyphenWithEnDash(element) {
    let count = 0;
    try {
        // Regex: (\d)\s*-\s*(\d) -> $1–$2
        // Znajdź cyfrę (\d), zero lub więcej spacji (\s*), myślnik (-), zero lub więcej spacji (\s*), cyfrę (\d).
        // Zastąp: pierwsza cyfra ($1), półpauza (–), druga cyfra ($2).
        count = element.replaceText("(\\d)\\s*-\\s*(\\d)", "$1–$2");
    } catch (e) { Logger.log(`replaceHyphenWithEnDash Error: ${e}`); }
    return count;
}

/** Wstawia twardą spację (\u00A0) między liczbą a wybranymi jednostkami/skrótami. */
function insertNbspAfterNumbers(element) {
    let count = 0;
    try {
        // Regex: (\d+)\s+(zł|...)\b -> $1\u00A0$2
        // Znajdź cyfry (\d+), spacje (\s+), jedną z jednostek (zł|...), granicę słowa (\b).
        // Zastąp: cyfry ($1), twarda spacja (\u00A0), jednostka ($2).
        // Flaga 'i' dla ignorowania wielkości liter (np. PLN).
        const regex = /(\d+)\s+(zł|PLN|kg|g|m|cm|mm|km|r\.|w\.|%|tys\.|mln|mld|s|min|godz\.)\b/gi;
        const replacement = "$1\u00A0$2";
        count = element.replaceText(regex, replacement);
    } catch (e) { Logger.log(`insertNbspAfterNumbers Error: ${e}`); }
    return count;
}


/**
 * Handler dla menu sprawdzania parzystości cudzysłowów.
 */
function runQuoteVerificationUniversal() {
  Logger.log("runQuoteVerificationUniversal: Start");
  const ui = DocumentApp.getUi();
  let doc;
  try {
      doc = DocumentApp.getActiveDocument();
      if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu.");
  } catch (e) { Logger.log(`runQuoteVerificationUniversal: Błąd getActiveDocument: ${e}`); ui.alert("Błąd Krytyczny", `Błąd dostępu do dokumentu: ${e.message}.`, ui.ButtonSet.OK); return; }

  try { clearHighlights(); } catch (e) { Logger.log(`runQuoteVerificationUniversal: Błąd clearHighlights(): ${e}. Kontynuuję.`); }

  const result = verifyQuoteParityJs(doc);
  Logger.log(`runQuoteVerificationUniversal: Wynik weryfikacji: isEven=${result.isEven}, count=${result.count}`);

  if (result.isEven) {
    ui.alert('Sukces', `Znaleziono parzystą liczbę (${result.count}) cudzysłowów (różnych typów).`, ui.ButtonSet.OK);
  } else {
    ui.alert('Błąd Parzystości', `Znaleziono nieparzystą liczbę (${result.count}) cudzysłowów (różnych typów). Ostatni znaleziony cudzysłów został podświetlony na czerwono. Popraw dokument.`, ui.ButtonSet.OK);
    const lastRangeEl = result.lastQuoteRangeElement;
    // Próba podświetlenia ostatniego znalezionego cudzysłowu
    if (lastRangeEl && typeof lastRangeEl.getElement === 'function') {
       const element = lastRangeEl.getElement();
       if (element && element.getType() === DocumentApp.ElementType.TEXT && typeof lastRangeEl.getStartOffsetInclusive === 'function' && typeof lastRangeEl.getEndOffsetExclusive === 'function') {
         try {
            // Używamy offsetów z RangeElement
            element.asText().setBackgroundColor(lastRangeEl.getStartOffsetInclusive(), lastRangeEl.getEndOffsetExclusive() -1 , ERROR_HIGHLIGHT_COLOR);
         } catch (e) { Logger.log(`runQuoteVerificationUniversal: Nie udało się podświetlić błędu parzystości: ${e}`); }
       }
    }
  }
   Logger.log("runQuoteVerificationUniversal: Koniec");
}


/**
 * Główna funkcja logiki zamiany cudzysłowów (wywoływana przez handler). Używa metody JS `indexOf`.
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 */
function replaceQuotesJs(doc) {
    Logger.log("replaceQuotesJs: Start");

    // 1. Zbierz lokalizacje wszystkich cudzysłowów
    const allLocations = [];
    collectAllQuoteLocationsRecursive(doc.getBody(), allLocations); // Użyj nowej nazwy funkcji pomocniczej
    collectAllQuoteLocationsRecursive(doc.getHeader(), allLocations);
    collectAllQuoteLocationsRecursive(doc.getFooter(), allLocations);
    const footnotes = doc.getFootnotes();
    if (footnotes) footnotes.forEach(fn => { if(fn && fn.getFootnoteContents) collectAllQuoteLocationsRecursive(fn.getFootnoteContents(), allLocations); });
    Logger.log(`replaceQuotesJs: Zebrano ${allLocations.length} lokalizacji.`);

    if (allLocations.length === 0) {
        Logger.log("replaceQuotesJs: Brak cudzysłowów do zamiany.");
        return;
    }

    // 2. Sortowanie wewnątrz elementów - kluczowe dla poprawnej kolejności `indexOf`
    try {
         allLocations.sort((a, b) => {
              if (a.element === b.element) { return a.index - b.index; }
              // Brak łatwego globalnego sortowania między elementami, polegamy na kolejności przetwarzania
              return 0;
         });
     } catch (e) {
         Logger.log(`replaceQuotesJs: Błąd podczas sortowania: ${e}. Kolejność może być nieoptymalna.`);
     }

    // 3. Zamiana od końca listy lokalizacji, aby nie psuć indeksów
    Logger.log("replaceQuotesJs: Rozpoczęcie zamiany od końca...");
    let replacementsCount = 0;
    for (let i = allLocations.length - 1; i >= 0; i--) {
        const location = allLocations[i];

        if (!location || typeof location.index !== 'number' || !location.element || typeof location.element.deleteText !== 'function') {
            Logger.log(`replaceQuotesJs: Pomijam nieprawidłową lokalizację na indeksie ${i}.`);
            continue;
        }

        const textElement = location.element;
        const index = location.index;
        const originalChar = location.char;
        const originalCharLength = originalChar.length;

        // Wyznacz poprawny polski cudzysłów (nieparzysty index = otwierający, parzysty = zamykający)
        const replacement = ((i + 1) % 2 === 0) ? POLISH_CLOSING_QUOTE : POLISH_OPENING_QUOTE;

        try {
             // Sprawdź, czy znak się nie zmienił od czasu skanowania (ważne przy modyfikacjach)
             const currentChar = textElement.getText().substring(index, index + originalCharLength);
             if (currentChar === originalChar) {
                 textElement.deleteText(index, index + originalCharLength - 1);
                 textElement.insertText(index, replacement);
                 replacementsCount++;
             } else {
                  Logger.log(`replaceQuotesJs: Oczekiwano '${originalChar}' @${index}, znaleziono '${currentChar}'. Pomijam.`);
             }
        } catch(e) {
            Logger.log(`replaceQuotesJs: Błąd podczas zamiany @${index}: ${e}`);
        }
    }
    Logger.log(`replaceQuotesJs: Zakończono zamianę. Wykonano ${replacementsCount} zamian.`);
}

/**
 * Handler dla menu zamiany cudzysłowów na polskie.
 */
function runQuoteReplacementUniversal() {
  Logger.log("runQuoteReplacementUniversal: Start");
  const ui = DocumentApp.getUi();
  let doc;
  try {
      doc = DocumentApp.getActiveDocument();
      if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu.");
  } catch (e) { Logger.log(`runQuoteReplacementUniversal: Błąd getActiveDocument: ${e}`); ui.alert("Błąd Krytyczny", `Wystąpił błąd podczas dostępu do dokumentu: ${e.message}.`, ui.ButtonSet.OK); return; }

  try { clearHighlights(); } catch (e) { Logger.log(`runQuoteReplacementUniversal: Błąd clearHighlights(): ${e}`); }

  // Najpierw weryfikacja parzystości
  const verificationResult = verifyQuoteParityJs(doc);
  Logger.log(`runQuoteReplacementUniversal: Wynik weryfikacji: isEven=${verificationResult.isEven}, count=${verificationResult.count}`);

  if (!verificationResult.isEven) {
    ui.alert('Anulowano', `Nie można zamienić cudzysłowów, ponieważ ich łączna liczba (${verificationResult.count}) jest nieparzysta. Popraw dokument.`, ui.ButtonSet.OK);
    // Można by tu dodać kod podświetlania błędu jak w runQuoteVerificationUniversal, jeśli potrzebne
    return;
  }

  if (verificationResult.count === 0) {
       ui.alert('Informacja', `Nie znaleziono żadnych cudzysłowów (${QUOTES_TO_FIND.join(', ')}) do zamiany.`, ui.ButtonSet.OK);
       return;
  }

  // Potwierdzenie od użytkownika
  const confirmation = ui.alert(
    'Potwierdzenie Zamiany Cudzysłowów',
    `Znaleziono ${verificationResult.count} cudzysłowów (${QUOTES_TO_FIND.join(', ')}). Czy chcesz zamienić je na polskie („ ”)? Tej operacji nie można cofnąć standardowym Ctrl+Z.`,
    ui.ButtonSet.YES_NO
  );

  if (confirmation === ui.Button.YES) {
    replaceQuotesJs(doc); // Wywołanie głównej logiki zamiany
    ui.alert('Sukces', 'Zamiana wszystkich typów cudzysłowów na polskie zakończona.', ui.ButtonSet.OK);
  } else {
    ui.alert('Anulowano', 'Operacja zamiany cudzysłowów została anulowana.', ui.ButtonSet.OK);
  }
   Logger.log("runQuoteReplacementUniversal: Koniec");
}


// --- Funkcje Pomocnicze (Logika Podstawowa) ---

/**
 * Znajduje indeksy wszystkich szukanych cudzysłowów w danym elemencie Text używając `indexOf`.
 * @param {GoogleAppsScript.Document.Text} textElement Element Text do przeszukania.
 * @return {Array<{index: number, char: string, element: GoogleAppsScript.Document.Text}>} Tablica lokalizacji cudzysłowów.
 */
function findQuoteIndicesInTextElement(textElement) {
    const text = textElement.getText();
    const indices = [];
    if (!text) return indices;

    for (const quoteChar of QUOTES_TO_FIND) {
        let fromIndex = 0;
        let index;
        while ((index = text.indexOf(quoteChar, fromIndex)) !== -1) {
            indices.push({ index: index, char: quoteChar, element: textElement });
            fromIndex = index + 1;
        }
    }
    // Sortowanie w ramach elementu jest ważne dla poprawnej kolejności
    indices.sort((a, b) => a.index - b.index);
    return indices;
}

/**
 * Rekurencyjnie przechodzi przez elementy dokumentu (np. akapity, listy, komórki tabel)
 * i zbiera lokalizacje wszystkich szukanych cudzysłowów do tablicy `allLocations`.
 * @param {GoogleAppsScript.Document.Element} element Bieżący element do przetworzenia.
 * @param {Array<{index: number, char: string, element: GoogleAppsScript.Document.Text}>} allLocations Tablica akumulująca znalezione lokalizacje.
 */
function collectAllQuoteLocationsRecursive(element, allLocations) { // Zmieniono nazwę z ...JsRecursive
    if (!element) return;

    let elementType;
    try { elementType = element.getType(); } catch (e) { Logger.log(`collectAllQuoteLocationsRecursive: Error getting element type: ${e}`); return; }

    switch (elementType) {
        case DocumentApp.ElementType.TEXT:
            const indicesInElement = findQuoteIndicesInTextElement(element.asText());
            if (indicesInElement.length > 0) {
                allLocations.push(...indicesInElement);
            }
            break;

        // Rekurencyjne przetwarzanie kontenerów
        case DocumentApp.ElementType.PARAGRAPH:
        case DocumentApp.ElementType.LIST_ITEM:
        case DocumentApp.ElementType.TABLE_CELL:
        case DocumentApp.ElementType.BODY_SECTION:
        case DocumentApp.ElementType.HEADER_SECTION:
        case DocumentApp.ElementType.FOOTER_SECTION:
        case DocumentApp.ElementType.FOOTNOTE_SECTION: // Kontener dla treści przypisu
            if (typeof element.getNumChildren === 'function') {
                const numChildren = element.getNumChildren();
                for (let i = 0; i < numChildren; i++) {
                    try { collectAllQuoteLocationsRecursive(element.getChild(i), allLocations); }
                    catch(e) { Logger.log(`collectAllQuoteLocationsRecursive: Error processing child ${i} of ${elementType}: ${e}`); }
                }
            }
            break;

        case DocumentApp.ElementType.FOOTNOTE: // Sam obiekt przypisu
            if (typeof element.getFootnoteContents === 'function') {
                try { collectAllQuoteLocationsRecursive(element.getFootnoteContents(), allLocations); }
                catch (e) { Logger.log(`collectAllQuoteLocationsRecursive: Error processing footnote contents: ${e}`); }
            }
            break;

        case DocumentApp.ElementType.TABLE:
             try {
                const numRows = element.getNumRows();
                for(let i = 0; i < numRows; i++){
                    const row = element.getRow(i);
                    const numCells = row.getNumCells();
                    for(let j = 0; j < numCells; j++) {
                       try { collectAllQuoteLocationsRecursive(row.getCell(j), allLocations); }
                       catch (cellErr) { Logger.log(`collectAllQuoteLocationsRecursive: Error processing cell [${i},${j}]: ${cellErr}`);}
                    }
                }
             } catch(e) { Logger.log(`collectAllQuoteLocationsRecursive: Error processing table: ${e}`); }
             break;

        // Ignoruj inne, nie tekstowe lub nie zawierające tekstu typy elementów
        default:
            break;
    }
}

/**
 * Weryfikuje parzystość wszystkich znalezionych cudzysłowów w dokumencie. Używa `indexOf`.
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 * @return {{count: number, isEven: boolean, lastQuoteRangeElement: GoogleAppsScript.Document.RangeElement | null}} Obiekt z wynikiem weryfikacji.
 */
function verifyQuoteParityJs(doc) {
    Logger.log("verifyQuoteParityJs: Start");
    const allLocations = [];

    // Zbierz lokalizacje z całego dokumentu
    collectAllQuoteLocationsRecursive(doc.getBody(), allLocations); // Użyj nowej nazwy
    collectAllQuoteLocationsRecursive(doc.getHeader(), allLocations);
    collectAllQuoteLocationsRecursive(doc.getFooter(), allLocations);
    const footnotes = doc.getFootnotes();
    if (footnotes) {
        footnotes.forEach((fn, index) => {
            if(fn && typeof fn.getFootnoteContents === 'function') {
                try {
                    const contents = fn.getFootnoteContents();
                    if (contents) collectAllQuoteLocationsRecursive(contents, allLocations);
                } catch (fnErr) { Logger.log(`verifyQuoteParityJs: Error processing footnote ${index+1}: ${fnErr}`); }
            }
        });
    }
    Logger.log(`verifyQuoteParityJs: Zebrano ${allLocations.length} lokalizacji.`);

    const count = allLocations.length;
    const isEven = count % 2 === 0;

    // Dla celów podświetlenia błędu, znajdź ostatni element (heurystycznie, wg kolejności przetwarzania)
    let lastLocation = count > 0 ? allLocations[count - 1] : null;
    let lastRangeElement = null;

    if (lastLocation) {
        // Spróbuj stworzyć RangeElement wskazujący na ostatni znaleziony znak
        try {
             lastRangeElement = DocumentApp.getActiveDocument().newRange()
                 .addElement(lastLocation.element, lastLocation.index, lastLocation.index) // Zakres dł. 1 znaku
                 .build()
                 .getRangeElements()[0];
        } catch (e) {
            Logger.log(`verifyQuoteParityJs: Nie udało się utworzyć RangeElement dla ostatniej lokalizacji: ${e}`);
            lastRangeElement = null;
        }
    }

    Logger.log(`verifyQuoteParityJs: Zakończono. Zwracam count=${count}, isEven=${isEven}`);
    return { count: count, isEven: isEven, lastQuoteRangeElement: lastRangeElement };
}


/**
 * Podświetla treść pomiędzy parami cudzysłowów. Używa `indexOf`.
 * OGRANICZENIE: Działa tylko dla par znajdujących się w tym samym elemencie Text (np. akapicie).
 * @param {GoogleAppsScript.Document.Document} doc Aktywny dokument.
 * @return {number} Liczba pomyślnie podświetlonych fragmentów.
 */
function highlightContentBetweenQuotesJs(doc) {
    Logger.log("highlightContentBetweenQuotesJs: Start");

    // 1. Zbierz lokalizacje
    const allLocations = [];
    collectAllQuoteLocationsRecursive(doc.getBody(), allLocations); // Użyj nowej nazwy
    collectAllQuoteLocationsRecursive(doc.getHeader(), allLocations);
    collectAllQuoteLocationsRecursive(doc.getFooter(), allLocations);
    const footnotes = doc.getFootnotes();
    if (footnotes) footnotes.forEach(fn => { if(fn && fn.getFootnoteContents) collectAllQuoteLocationsRecursive(fn.getFootnoteContents(), allLocations); });
    Logger.log(`highlightContentBetweenQuotesJs: Zebrano ${allLocations.length} lokalizacji.`);

    if (allLocations.length < 2) {
        return 0; // Brak par do podświetlenia
    }

    // 2. Sortowanie wewnątrz elementów (kluczowe dla znajdowania par)
     try {
         allLocations.sort((a, b) => {
              if (a.element === b.element) { return a.index - b.index; }
              return 0;
         });
     } catch (e) { Logger.log(`highlightContentBetweenQuotesJs: Błąd sortowania: ${e}.`); }

    // 3. Iteracja przez pary (co drugi element) i podświetlanie
    let highlightedCount = 0;
    for (let i = 0; i < allLocations.length - 1; i += 2) {
        const locOpen = allLocations[i];
        const locClose = allLocations[i + 1];

        // Sprawdź, czy para jest w tym samym elemencie TEXT
        if (locOpen && locClose && locOpen.element === locClose.element && locOpen.element.getType() === DocumentApp.ElementType.TEXT) {
            const textElement = locOpen.element;
            const startIndex = locOpen.index + locOpen.char.length; // Pozycja PO cudzysłowie otwierającym
            const endIndex = locClose.index - 1;                   // Pozycja PRZED cudzysłowem zamykającym

            // Sprawdź, czy zakres jest poprawny
            if (endIndex >= startIndex) {
                try {
                    textElement.setBackgroundColor(startIndex, endIndex, HIGHLIGHT_COLOR);
                    highlightedCount++;
                } catch (e) {
                    Logger.log(`highlightContentBetweenQuotesJs: Błąd setBackgroundColor [${startIndex}-${endIndex}]: ${e}`);
                }
            }
        }
        // Ignoruj pary między różnymi elementami
    }

    Logger.log(`highlightContentBetweenQuotesJs: Zakończono. Podświetlono ${highlightedCount} fragmentów.`);
    return highlightedCount;
}

/**
 * Handler dla menu podświetlania treści wewnątrz cudzysłowów.
 */
function runHighlightQuoteContentJs() {
    Logger.log("runHighlightQuoteContentJs: Start");
    const ui = DocumentApp.getUi();
    let doc;
    try {
        doc = DocumentApp.getActiveDocument();
        if (!doc) throw new Error("Nie można uzyskać dostępu do dokumentu.");
    }
    catch (e) { Logger.log(`runHighlightQuoteContentJs: Błąd getActiveDocument: ${e}`); ui.alert("Błąd Krytyczny", "Nie można uzyskać dostępu do aktywnego dokumentu.", ui.ButtonSet.OK); return; }

    try { clearHighlights(); }
    catch (e) { Logger.log(`runHighlightQuoteContentJs: Błąd clearHighlights: ${e}`); }

    let highlightedCount = 0;
    try {
         highlightedCount = highlightContentBetweenQuotesJs(doc);
         Logger.log(`runHighlightQuoteContentJs: Wynik podświetlania: ${highlightedCount}`);

         if (highlightedCount > 0) {
             ui.alert('Podświetlono Treść', `Podświetlono treść wewnątrz ${highlightedCount} par cudzysłowów (tylko pary wewnątrz tego samego akapitu).`, ui.ButtonSet.OK);
         } else {
             ui.alert('Nie Znaleziono', 'Nie znaleziono par cudzysłowów wewnątrz tego samego akapitu do podświetlenia lub wystąpiły błędy.', ui.ButtonSet.OK);
         }
    } catch (e) {
         Logger.log(`runHighlightQuoteContentJs: Błąd podczas highlightContentBetweenQuotesJs: ${e}`);
         ui.alert("Błąd", "Wystąpił błąd podczas próby podświetlenia cytatów.", ui.ButtonSet.OK);
    }
    Logger.log("runHighlightQuoteContentJs: Koniec");
}

/**
 * Usuwa podświetlenie tła (ustawia na null) z całego dokumentu (body, header, footer, footnotes).
 */
function clearHighlights() {
     const doc = DocumentApp.getActiveDocument();
      // Funkcja pomocnicza do rekurencyjnego czyszczenia
      function clearElementHighlightRecursive(element) {
         if (!element) return;
         let elementType;
         try { elementType = element.getType(); } catch (e) { return; }

         switch (elementType) {
            case DocumentApp.ElementType.TEXT:
              // Usuń tło tylko jeśli element ma jakąś treść
              try { if (element.asText().getText().length > 0) element.asText().setBackgroundColor(null); }
              catch (e) { Logger.log(`clearHighlights: Nie można wyczyścić tła dla Text: ${e}`); }
              break;
            // Rekurencja dla kontenerów
            case DocumentApp.ElementType.PARAGRAPH:
            case DocumentApp.ElementType.LIST_ITEM:
            case DocumentApp.ElementType.TABLE_CELL:
            case DocumentApp.ElementType.BODY_SECTION:
            case DocumentApp.ElementType.HEADER_SECTION:
            case DocumentApp.ElementType.FOOTER_SECTION:
            case DocumentApp.ElementType.FOOTNOTE_SECTION:
              if (typeof element.getNumChildren === 'function') {
                  const numChildren = element.getNumChildren();
                  for (let i = 0; i < numChildren; i++) clearElementHighlightRecursive(element.getChild(i));
              }
              break;
            case DocumentApp.ElementType.FOOTNOTE:
               if (typeof element.getFootnoteContents === 'function') {
                   clearElementHighlightRecursive(element.getFootnoteContents());
               }
               break;
            case DocumentApp.ElementType.TABLE:
               try {
                   const numRows = element.getNumRows();
                   for(let i=0; i < numRows; i++){
                       const row = element.getRow(i);
                       const numCells = row.getNumCells();
                       for(let j=0; j < numCells; j++) clearElementHighlightRecursive(row.getCell(j));
                   }
               } catch(e) { Logger.log(`clearHighlights: Błąd przetwarzania tabeli: ${e}`); }
               break;
            default:
               break; // Ignoruj inne typy
         }
      }
      // Uruchom czyszczenie dla głównych sekcji
      clearElementHighlightRecursive(doc.getBody());
      const header = doc.getHeader(); if (header) clearElementHighlightRecursive(header);
      const footer = doc.getFooter(); if (footer) clearElementHighlightRecursive(footer);
      const footnotes = doc.getFootnotes();
      if (footnotes) footnotes.forEach(footnote => { if (footnote) clearElementHighlightRecursive(footnote); }); // Przetwarzamy obiekt footnote, który zawiera FOOTNOTE_SECTION
      Logger.log("clearHighlights: Wyczyszczono podświetlenia.");
}
