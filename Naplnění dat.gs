/**
 * Zpracuje data pro pracovníky - OPRAVENÁ VERZE
 * Opraveno pořadí formátování: střídavé obarvení → zpracování dat → červené označení chyb
 * Opraveno hledání DPČ (normalizace diakritiky)
 * Přidány detailní logy pro diagnostiku
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Vlastní nástroje')
      .addItem('📁 Nahrát výkazy', 'processDataForWorkers')
      .addItem('🔄 Aktualizuj řádek', 'reloadCurrentRow')
      .addSeparator()
      .addItem('📊 Spustit sumarizaci', 'startSumarizace')
    .addToUi();
}

// Automatická kontrola při editaci buňky
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // Ignorovat první řádek (hlavička)
  if (row === 1) return;
  
  // Kontrolovat pouze sloupce E a dále (sloupec 5+)
  if (col < 5) return;
  
  // Kontrolovat pouze listy s měsíčními názvy (ne "Adresář", "Zdrojová data", atd.)
  const sheetName = sheet.getName();
  if (sheetName === 'Adresář' || sheetName === 'Zdrojová data' || sheetName === 'Podstřediska') return;
  
  // Kontrola součtu pro tento řádek
  checkRowSum(sheet, row);
}

function checkRowSum(sheet, row) {
  const data = sheet.getDataRange().getValues();
  const totalAmountColIndex = 4; // Sloupec E (index 4)
  
  // Načíst cílovou částku ze sloupce E
  const targetAmount = Math.ceil(data[row - 1][totalAmountColIndex]);
  
  // Pokud je cílová částka 0 nebo prázdná, přeskočit kontrolu
  if (!targetAmount || targetAmount === 0) {
    Logger.log(`ℹ️ Řádek ${row}: Prázdný řádek nebo nulová částka, přeskakuji kontrolu`);
    return;
  }
  
  // Načíst částky z podstředisek
  // Sloupec H = index 7, sloupec I = index 8
  // Částky jsou v I (8), K (10), M (12), O (14)...
  let sumFromSubcenters = 0;
  let hasAnyData = false;
  
  // Začínáme od sloupce I (index 8), pak přeskakujeme 2 (8, 10, 12, 14...)
  for (let colIndex = 8; colIndex < data[row - 1].length; colIndex += 2) {
    const castkaValue = data[row - 1][colIndex];
    
    if (castkaValue !== '' && castkaValue !== null && castkaValue !== undefined) {
      hasAnyData = true;
      if (typeof castkaValue === 'number') {
        sumFromSubcenters += castkaValue;
      } else if (typeof castkaValue === 'string') {
        const parsed = parseFloat(castkaValue.replace(/[^\d.-]/g, ''));
        if (!isNaN(parsed)) {
          sumFromSubcenters += parsed;
        }
      }
    }
  }
  
  // Pokud nejsou žádná data v podstřediscích, přeskočit kontrolu
  if (!hasAnyData) {
    Logger.log(`ℹ️ Řádek ${row}: Žádná data v podstřediscích, přeskakuji kontrolu`);
    return;
  }
  
  Logger.log(`🔍 Řádek ${row}: Cíl=${targetAmount}, Součet=${sumFromSubcenters}`);
  
  const rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  
  if (sumFromSubcenters !== targetAmount) {
    // Nesoulad - obarvit červeně
    rowRange.setBackground('#f4cccc');
    Logger.log(`⚠️ Řádek ${row}: Nesoulad ${targetAmount} vs ${sumFromSubcenters}`);
  } else {
    // Shoda - odstranit červenou, obnovit formátování
    Logger.log(`✅ Řádek ${row}: Součty se shodují, obnovuji formátování`);
    
    // Odstranit červené pozadí
    rowRange.setBackground(null);
    
    // Obnovit střídavé obarvení sloupců od H dále
    const lastCol = sheet.getLastColumn();
    if (lastCol >= 8) {
      let col = 8;
      let pairIndex = 0;
      
      while (col <= lastCol) {
        if (pairIndex % 2 === 0) {
          const range = sheet.getRange(row, col, 1, 2);
          range.setBackground('#f3f3f3');
        } else {
          const range = sheet.getRange(row, col, 1, 2);
          range.setBackground(null);
        }
        col += 2;
        pairIndex++;
      }
    }
    
    // Zachovat zelený text sloupce A (pokud tam není ❌)
    const cellA = sheet.getRange(row, 1);
    const valueA = cellA.getValue();
    if (valueA && valueA.toString().includes('✅')) {
      // Pokud je to rich text s více odkazy, zachovat formátování
      const richTextValue = cellA.getRichTextValue();
      if (richTextValue) {
        // Má rich text - nechat beze změny (už má správné barvy)
        // Nedělat nic
      } else {
        // Prostý HYPERLINK - nastavit zelenou
        cellA.setFontColor('#00ff00');
      }
    }
  }
}

// Nová pomocná funkce - okamžitě se vrátí a spustí zpracování na pozadí
function startProcessingVykazy(selectedIndex) {
    Logger.log(`========== START PROCESSING VYKAZY ==========`);
    Logger.log(`Typ parametru: ${typeof selectedIndex}`);
    Logger.log(`Hodnota parametru: ${selectedIndex}`);
    Logger.log(`Je null? ${selectedIndex === null}`);
    Logger.log(`Je undefined? ${selectedIndex === undefined}`);
    Logger.log(`======================================`);
    
    if (selectedIndex === null || selectedIndex === undefined) {
        Logger.log('❌ CHYBA: selectedIndex je null nebo undefined');
        SpreadsheetApp.getActiveSpreadsheet().toast('Chyba: Nebyl vybrán žádný měsíc', 'Chyba', 5);
        return;
    }
    
    // Spustit zpracování asynchronně
    SpreadsheetApp.getActiveSpreadsheet().toast('Spouštím zpracování...', 'Připravuji', 3);
    Logger.log('✅ Spouštím processSelectedMonthVykazy...');
    
    // Pak spustit samotné zpracování
    processSelectedMonthVykazy(selectedIndex);
}

function processDataForWorkers() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ===== NAČTENÍ DAT Z LISTU "Adresář" =====
    SpreadsheetApp.getActiveSpreadsheet().toast('Načítám dostupné měsíce...', 'Načítání', -1);
    
    const adresarSheet = ss.getSheetByName('Adresář');
    
    if (!adresarSheet) {
        ui.alert('Chyba: List "Adresář" nebyl nalezen. Vytvořte list "Adresář" se sloupci A (Měsíc) a B (ID adresáře).');
        return;
    }
    
    Logger.log('Načítám data z listu "Adresář"...');
    const adresarData = adresarSheet.getDataRange().getValues();
    
    // Načtení měsíců s vyplněným ID (od řádku 2)
    const availableMonths = [];
    for (let i = 1; i < adresarData.length; i++) {
        const month = String(adresarData[i][0]).trim();
        const folderId = String(adresarData[i][1]).trim();
        
        if (month && folderId) {
            availableMonths.push({
                month: month,
                folderId: folderId,
                rowIndex: i + 1
            });
            Logger.log(`  Řádek ${i + 1}: Měsíc="${month}", ID="${folderId}"`);
        }
    }
    
    if (availableMonths.length === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast('');
        ui.alert('Chyba: V listu "Adresář" nejsou žádné měsíce s vyplněným ID adresáře.\n\nVyplňte prosím sloupec A (Měsíc) a sloupec B (ID adresáře).');
        Logger.log('❌ Žádné měsíce s ID nenalezeny');
        return;
    }
    
    Logger.log(`✅ Nalezeno ${availableMonths.length} měsíců s ID`);
    
    // ===== VYTVOŘENÍ HTML DIALOGU S RADIO BUTTONY =====
    SpreadsheetApp.getActiveSpreadsheet().toast('');

    let htmlContent = '<style>';
    htmlContent += 'body { font-family: Arial, sans-serif; padding: 10px; }';
    htmlContent += 'label { display: block; padding: 6px 0; }';
    htmlContent += '.buttons { margin-top: 16px; text-align: right; }';
    htmlContent += 'button { margin-left: 8px; padding: 6px 14px; }';
    htmlContent += '</style>';

    htmlContent += '<h3>Vyberte měsíc pro zpracování</h3>';
    htmlContent += '<div id="months">';

    availableMonths.forEach((item, idx) => {
        htmlContent += '<label>';
        htmlContent += `<input type="radio" name="month" value="${idx}"${idx === 0 ? ' checked' : ''}> `;
        htmlContent += `${item.month}`;
        htmlContent += '</label>';
    });

    htmlContent += '</div>';
    htmlContent += '<div class="buttons">';
    htmlContent += '<button onclick="google.script.host.close()">Zrušit</button>';
    htmlContent += '<button onclick="submitSelection()">Potvrdit</button>';
    htmlContent += '</div>';

    htmlContent += '<script>';
    htmlContent += 'function submitSelection() {';
    htmlContent += '  var selected = document.querySelector(\'input[name="month"]:checked\');';
    htmlContent += '  if (!selected) { return; }';
    htmlContent += '  google.script.run.withSuccessHandler(google.script.host.close).startProcessingVykazy(selected.value);';
    htmlContent += '}';
    htmlContent += '</script>';

    const html = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(400)
        .setHeight(400);

    ui.showModalDialog(html, 'Výběr měsíce');
}

function processSelectedMonthVykazy(selectedIndex) {
    Logger.log(`\n========== PROCESS SELECTED MONTH VYKAZY ==========`);
    Logger.log(`Parametr selectedIndex: ${selectedIndex}`);
    Logger.log(`Typ: ${typeof selectedIndex}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const adresarSheet = ss.getSheetByName('Adresář');
    
    Logger.log(`Adresář sheet: ${adresarSheet ? 'OK' : 'NOT FOUND'}`);
    
    const adresarData = adresarSheet.getDataRange().getValues();
    Logger.log(`Počet řádků v adresáři: ${adresarData.length}`);
    
    const availableMonths = [];
    for (let i = 1; i < adresarData.length; i++) {
        const month = String(adresarData[i][0]).trim();
        const folderId = String(adresarData[i][1]).trim();
        if (month && folderId) {
            availableMonths.push({ month: month, folderId: folderId });
            Logger.log(`  availableMonths[${availableMonths.length - 1}]: ${month}`);
        }
    }
    
    Logger.log(`Celkem dostupných měsíců: ${availableMonths.length}`);
    Logger.log(`Parsovaný index: ${parseInt(selectedIndex)}`);
    
    const selectedMonth = availableMonths[parseInt(selectedIndex)];
    Logger.log(`Vybraný měsíc: ${selectedMonth ? JSON.stringify(selectedMonth) : 'NULL'}`);
    
    if (!selectedMonth) {
        Logger.log('❌ CHYBA: selectedMonth je null');
        SpreadsheetApp.getActiveSpreadsheet().toast('Chyba: Neplatný index měsíce', 'Chyba', 5);
        return;
    }
    
    const FOLDER_ID = selectedMonth.folderId;
    let MONTH_NAME = selectedMonth.month;
    
    Logger.log(`\n========== VYBRANÝ MĚSÍC ==========`);
    Logger.log(`Měsíc: ${MONTH_NAME}`);
    Logger.log(`ID adresáře: "${FOLDER_ID}"`);
    Logger.log(`===================================\n`);
    
    // ===== KONTROLA EXISTENCE LISTU S TÍMTO NÁZVEM =====
    const sourceSheet = ss.getSheetByName('Zdrojová data');
    if (!sourceSheet) {
        SpreadsheetApp.getActiveSpreadsheet().toast('');
        ui.alert('Chyba: List "Zdrojová data" nebyl nalezen.');
        return;
    }
    
    let targetSheet = ss.getSheetByName(MONTH_NAME);
    
    if (targetSheet) {
        // List již existuje - nabídneme 3 možnosti
        const response = ui.alert(
            `List "${MONTH_NAME}" již existuje`,
            'Co chcete udělat?\n\n' +
            '• ANO = Přepsat existující list\n' +
            '• NE = Vytvořit nový list s číslem\n' +
            '• ZRUŠIT = Zrušit operaci',
            ui.ButtonSet.YES_NO_CANCEL
        );
        
        if (response === ui.Button.YES) {
            // Přepsat existující list
            SpreadsheetApp.getActiveSpreadsheet().toast(`Mažu starý list "${MONTH_NAME}"...`, 'Příprava', -1);
            ss.deleteSheet(targetSheet);
            targetSheet = sourceSheet.copyTo(ss);
            targetSheet.setName(MONTH_NAME);
            Logger.log(`List "${MONTH_NAME}" byl přepsán.`);
            
        } else if (response === ui.Button.NO) {
            // Vytvořit nový list s číslem
            let counter = 2;
            let newName = `${MONTH_NAME} (${counter})`;
            
            while (ss.getSheetByName(newName)) {
                counter++;
                newName = `${MONTH_NAME} (${counter})`;
            }
            
            SpreadsheetApp.getActiveSpreadsheet().toast(`Vytvářím list "${newName}"...`, 'Příprava', -1);
            targetSheet = sourceSheet.copyTo(ss);
            targetSheet.setName(newName);
            MONTH_NAME = newName; // Aktualizujeme název pro další použití
            Logger.log(`Vytvořen nový list: "${newName}"`);
            
        } else {
            // CANCEL - zrušit operaci
            SpreadsheetApp.getActiveSpreadsheet().toast('');
            Logger.log('Operace zrušena uživatelem.');
            return;
        }
    } else {
        // List neexistuje - vytvoříme nový
        SpreadsheetApp.getActiveSpreadsheet().toast(`Vytvářím list pro měsíc: ${MONTH_NAME}...`, 'Příprava', -1);
        targetSheet = sourceSheet.copyTo(ss);
        targetSheet.setName(MONTH_NAME);
        Logger.log(`Vytvořen nový list: "${MONTH_NAME}"`);
    }
    
    ss.setActiveSheet(targetSheet);
    
    // ===== ÚPRAVA HLAVIČKY SLOUPCE A =====
    targetSheet.getRange('A1').setValue('Odkaz');
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Připravuji data...', 'Inicializace', -1);
    
    // ===== ODSTRANĚNÍ HYPERTEXTOVÝCH ODKAZŮ A RESET FORMÁTOVÁNÍ =====
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
        const prijemniRange = targetSheet.getRange(2, 2, lastRow - 1, 1);
        const values = prijemniRange.getValues();
        prijemniRange.clearContent();
        prijemniRange.setValues(values);
        
        const dataRangeReset = targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn());
        dataRangeReset.setBackground(null);
    }
    
    // ===== ZPRACOVÁNÍ DAT =====
    processSheetData(targetSheet, FOLDER_ID, MONTH_NAME);
}

function processSheetData(outputSheet, FOLDER_ID, MONTH_NAME) {
    const ui = SpreadsheetApp.getUi();
    const data = outputSheet.getDataRange().getValues();

    const prijemniColIndex = 1;
    const jmenoColIndex = 2;
    const druhSmlouvyColIndex = 3;
    const totalAmountColIndex = 4;
    const totalAmountColSheetIndex = totalAmountColIndex + 1;
    const missingFilesNames = [];
    const processedWorkersCount = data.length - 1;

    if (processedWorkersCount < 1) {
        SpreadsheetApp.getActiveSpreadsheet().toast('');
        ui.alert('Aktivní list neobsahuje žádná data.');
        return;
    }

    const CZK_FORMAT = '#,##0 "Kč"';
    const TEXT_FORMAT = '@';
    const podstrediskoRegex = /^\d{2}-\d{2}-\d{2}(\/\d+)?$/;

    const normalizeName = (name) => {
        return name.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toUpperCase();
    };

    const roundHours = (value) => {
        return Math.round(value * 100) / 100;
    };

    const convertDateToFormat = (value) => {
        if (value instanceof Date) {
            const day = String(value.getDate()).padStart(2, '0');
            const month = String(value.getMonth() + 1).padStart(2, '0');
            const year = String(value.getFullYear()).slice(-2);
            return `${day}-${month}-${year}`;
        }
        return String(value);
    };

    const detectContractType = (contractText) => {
        const text = String(contractText).trim().toLowerCase();
        if (text.includes('dohoda o provedení práce') || text.includes('dpp')) {
            return 'DPP';
        } else if (text.includes('dohoda o pracovní činnosti') || text.includes('dpč')) {
            return 'DPČ';
        }
        return null;
    };

    try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Načítám soubory z Google Drive...', 'Indexace', -1);
        
        const folder = DriveApp.getFolderById(FOLDER_ID);
        Logger.log(`Přístup do adresáře: ${folder.getName()} (ID: ${FOLDER_ID})`);
        
        const filesIterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        const fileCache = [];

        while (filesIterator.hasNext()) {
            const file = filesIterator.next();
            fileCache.push({
                name: normalizeName(file.getName()),
                originalName: file.getName(),
                file: file,
                processed: false
            });
        }
        
        Logger.log(`Indexace dokončena. Nalezeno ${fileCache.length} souborů.`);
        Logger.log(`Všechny soubory v cache (normalizované názvy):`);
        fileCache.forEach(entry => Logger.log(`  - "${entry.name}" (original: "${entry.originalName}")`));

        // ===== STŘÍDAVÉ OBARVENÍ SLOUPCŮ (NEJDŘÍVE!) =====
        SpreadsheetApp.getActiveSpreadsheet().toast('Nastavuji formátování...', 'Příprava', -1);
        
        const lastDataRow = outputSheet.getLastRow();
        if (lastDataRow > 1) {
            let col = 8;
            let pairIndex = 0;
            while (col <= outputSheet.getLastColumn()) {
                if (pairIndex % 2 === 0) {
                    const range = outputSheet.getRange(2, col, lastDataRow - 1, 2);
                    range.setBackground('#f3f3f3');
                } else {
                    const range = outputSheet.getRange(2, col, lastDataRow - 1, 2);
                    range.setBackground(null);
                }
                col += 2;
                pairIndex++;
            }
        }

        // ===== ZPRACOVÁNÍ JEDNOTLIVÝCH ŘÁDKŮ =====
        for (let i = 1; i < data.length; i++) {
            SpreadsheetApp.getActiveSpreadsheet().toast(`Zpracovávám ${i}/${processedWorkersCount}...`, 'Zpracování', -1);
            
            processRow(outputSheet, i, data, fileCache, {
                prijemniColIndex,
                jmenoColIndex,
                druhSmlouvyColIndex,
                totalAmountColIndex,
                totalAmountColSheetIndex,
                CZK_FORMAT,
                TEXT_FORMAT,
                podstrediskoRegex,
                normalizeName,
                roundHours,
                convertDateToFormat,
                detectContractType,
                missingFilesNames
            });
        }

        // ===== KONTROLA NEZPRACOVANÝCH SOUBORŮ =====
        const unprocessedFiles = fileCache.filter(entry => !entry.processed);
        if (unprocessedFiles.length > 0) {
            Logger.log(`\n⚠️ Nezpracované soubory v adresáři:`);
            
            let htmlContent = '<style>';
            htmlContent += 'body { font-family: Arial, sans-serif; padding: 10px; }';
            htmlContent += 'ul { list-style: none; padding: 0; margin: 0; }';
            htmlContent += 'li { padding: 6px 0; border-bottom: 1px solid #eee; }';
            htmlContent += '.buttons { margin-top: 16px; text-align: right; }';
            htmlContent += '</style>';

            htmlContent += '<h3>⚠️ Nezpracované soubory</h3>';
            htmlContent += '<p>V adresáři byly nalezeny následující soubory, které nebyly zpracovány:</p>';
            htmlContent += '<ul>';

            unprocessedFiles.forEach(entry => {
                Logger.log(`  - ${entry.originalName}`);
                htmlContent += '<li>';
                htmlContent += `📄 ${entry.originalName} `;
                htmlContent += `<a href="${entry.file.getUrl()}" target="_blank" rel="noopener noreferrer">Otevřít soubor →</a>`;
                htmlContent += '</li>';
            });

            htmlContent += '</ul>';
            htmlContent += '<div class="buttons"><button onclick="google.script.host.close()">Zavřít</button></div>';

            const html = HtmlService.createHtmlOutput(htmlContent)
                .setWidth(500)
                .setHeight(400);
            
            SpreadsheetApp.getActiveSpreadsheet().toast('');
            ui.showModalDialog(html, 'Nezpracované soubory');
        }

        if (missingFilesNames.length > 0) {
            Logger.log(`\n⚠️ Chybějící soubory pro pracovníky:`);
            missingFilesNames.forEach(name => Logger.log(`  - ${name}`));
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('✅ Zpracování dokončeno!', 'Hotovo', 5);

    } catch (e) {
        Logger.log(`KRITICKÁ CHYBA: ${e.toString()}`);
        SpreadsheetApp.getActiveSpreadsheet().toast('');
        ui.alert(`Došlo k chybě: ${e.message}`);
    }
}

function processRow(outputSheet, rowIndex, data, fileCache, config) {
    const i = rowIndex;
    const prijemni = data[i][config.prijemniColIndex];
    const jmeno = data[i][config.jmenoColIndex];
    const druhSmlouvy = data[i][config.druhSmlouvyColIndex];
    const amountInColumnD = data[i][config.totalAmountColIndex];

    if (!prijemni || !jmeno) return;

    const workerName = `${prijemni} ${jmeno}`.replace(/\s\s+/g, ' ').trim();
    const contractType = config.detectContractType(druhSmlouvy);

    // OPRAVA: normalizovat contractType stejně jako názvy souborů
    const normalizedContractType = contractType ? config.normalizeName(contractType) : null;

    const normalizedWorkerName = config.normalizeName(workerName);

    Logger.log(`\nZpracovávám: ${workerName} (${contractType || 'N/A'})`);
    Logger.log(`  Normalizované jméno: "${normalizedWorkerName}"`);
    Logger.log(`  Normalizovaný typ smlouvy: "${normalizedContractType}"`);

    // Hledání souborů podle normalizovaného jména
    const foundByName = fileCache.filter(entry => entry.name.includes(normalizedWorkerName));
    Logger.log(`  Soubory obsahující jméno (${foundByName.length}): ${foundByName.map(e => `"${e.originalName}"`).join(', ') || 'žádné'}`);

    let foundEntries = [];

    if (normalizedContractType) {
        // OPRAVA: hledáme normalizovaný typ smlouvy (DPC místo DPČ)
        Logger.log(`  Hledám soubory s "${workerName}" A "${contractType}" (normalizovaně: "${normalizedContractType}") v názvu`);
        foundEntries = foundByName.filter(entry => entry.name.includes(normalizedContractType));
        
        if (foundEntries.length > 0) {
            Logger.log(`  ✅ Nalezeno ${foundEntries.length} souborů s typem ${contractType}`);
        } else {
            Logger.log(`  ❌ Žádný soubor nenalezen s typem ${contractType} (normalizovaně: ${normalizedContractType})`);
            Logger.log(`  DEBUG - normalizované názvy všech souborů v cache: ${fileCache.map(e => `"${e.name}"`).join(', ')}`);
        }
    } else {
        Logger.log(`  ⚠️ Typ smlouvy není definován v řádku`);
        
        // OPRAVA: kontrolujeme normalizované názvy (DPC místo DPČ)
        const filesWithType = foundByName.filter(entry => 
            entry.name.includes('DPP') || entry.name.includes('DPC')
        );
        
        if (filesWithType.length > 0) {
            Logger.log(`  ⚠️ VAROVÁNÍ: Nalezeny soubory s DPP/DPČ v názvu, ale řádek nemá definován typ smlouvy!`);
            filesWithType.forEach(entry => Logger.log(`    - ${entry.originalName}`));
            Logger.log(`  → Přeskakuji tento řádek, doplňte typ smlouvy do sloupce "${data[0][config.druhSmlouvyColIndex]}"`);
            foundEntries = []; // Nevybereme žádný soubor
        } else {
            // Soubory nemají typ v názvu, můžeme je použít
            foundEntries = foundByName;
        }
    }

    Logger.log(`  Celkem k zpracování: ${foundEntries.length} soubor(ů)`);
    foundEntries.forEach(entry => Logger.log(`    - ${entry.originalName}`));

    // ===== ZELENÉ ZATRŽÍTKO / ČERVENÝ KŘÍŽEK VE SLOUPCI A =====
    const cellA = outputSheet.getRange(i + 1, 1);
    
    if (foundEntries.length === 0) {
        // Červený křížek - PŘEPÍŠE střídavé obarvení
        cellA.setValue('❌');
        cellA.setFontColor('#ff0000');
        cellA.setHorizontalAlignment('center');
        config.missingFilesNames.push(`${workerName} (${contractType || 'N/A'})`);
        
        const rowRange = outputSheet.getRange(i + 1, 1, 1, outputSheet.getLastColumn());
        rowRange.setBackground('#f4cccc');
        return;
    }

    // Označení souborů jako zpracovaných
    foundEntries.forEach(entry => entry.processed = true);

    // Vytvoření hyperlinků ve sloupci A
    const urls = foundEntries.map(entry => entry.file.getUrl());
    
    if (urls.length === 1) {
        // Jeden soubor - zatržítko bez podtržení
        cellA.setFormula(`=HYPERLINK("${urls[0]}"; "✅")`);
        cellA.setFontColor('#00ff00');
        cellA.setFontStyle('normal');
        cellA.setFontLine('none');
        cellA.setBackground(null);
    } else {
        // Více souborů - zatržítko zelené + čísla černá
        const richText = SpreadsheetApp.newRichTextValue();
        let textBuilder = '✅';
        
        urls.forEach((url, idx) => {
            textBuilder += ` ${idx + 1}`;
        });
        
        richText.setText(textBuilder);
        
        // Styl pro zatržítko - zelený bez podtržení
        const greenStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor('#00ff00')
            .setUnderline(false)
            .build();
        richText.setTextStyle(0, 1, greenStyle);
        
        // Styl pro čísla - černý bez podtržení
        const blackStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor('#000000')
            .setUnderline(false)
            .build();
        
        // Odkazy POUZE na čísla + černý text bez podtržení
        let currentPos = 2; // Po "✅ "
        urls.forEach((url, idx) => {
            const numText = (idx + 1).toString();
            const startPos = textBuilder.indexOf(numText, currentPos);
            const endPos = startPos + numText.length;
            
            // Přidat odkaz
            richText.setLinkUrl(startPos, endPos, url);
            
            // Černý text bez podtržení pro tento odkaz
            richText.setTextStyle(startPos, endPos, blackStyle);
            
            currentPos = endPos + 1;
        });
        
        cellA.setRichTextValue(richText.build());
        cellA.setBackground(null);
    }
    cellA.setHorizontalAlignment('center');
    
    // Komentář do sloupce B - pouze pokud je více souborů
    const cellB = outputSheet.getRange(i + 1, 2);
    if (urls.length > 1) {
        cellB.setNote(`Zpracováno z ${urls.length} souborů:\n${foundEntries.map(e => e.originalName).join('\n')}`);
    }

    // ===== ZPRACOVÁNÍ DAT ZE SOUBORŮ =====
    const workerDataAggregated = new Map();
    const workerDataByFile = new Map();
    let foundData = false;

    for (const foundEntry of foundEntries) {
        const foundFile = foundEntry.file;
        
        try {
            const spreadsheet = SpreadsheetApp.open(foundFile);
            Logger.log(`  Otevírám soubor: "${foundFile.getName()}"`);
            const fixedRateSheet = spreadsheet.getSheets()[0];
            const fixedSazbaForPremiums = fixedRateSheet.getRange('X1').getValue();
            Logger.log(`  Sazba z X1: ${fixedSazbaForPremiums} (typ: ${typeof fixedSazbaForPremiums})`);

            if (typeof fixedSazbaForPremiums !== 'number' || fixedSazbaForPremiums <= 0) {
                Logger.log(`  ❌ Neplatná sazba v X1 (hodnota: "${fixedSazbaForPremiums}") - přeskakuji soubor`);
                continue;
            }

            const sheets = spreadsheet.getSheets();
            Logger.log(`  Počet listů v souboru: ${sheets.length} → ${sheets.map(s => `"${s.getName()}"`).join(', ')}`);

            sheets.forEach(sheet => {
                const sheetName = sheet.getName();
                const sheetData = sheet.getDataRange().getValues();
                Logger.log(`    Prohledávám list: "${sheetName}" (${sheetData.length} řádků, ${sheetData[0] ? sheetData[0].length : 0} sloupců)`);
                
                let summaryRowIndex = -1;
                let podstrediskoColIndex = -1;
                let sazbaColIndex = -1;
                let pocetHodinColIndex = -1;
                let nocColIndex = -1;
                let vikendColIndex = -1;
                let svatekColIndex = -1;

                for (let j = 0; j < sheetData.length; j++) {
                    const summaryRow = sheetData[j];
                    const summaryColIndex = summaryRow.findIndex(cell => 
                        typeof cell === 'string' && cell.trim().toUpperCase() === 'SOUHRN'
                    );
                    
                    if (summaryColIndex !== -1) {
                        summaryRowIndex = j;
                        Logger.log(`    ✅ Nalezena sekce SOUHRN na řádku ${j + 1}, sloupci ${summaryColIndex + 1}`);
                        const headersRow = sheetData[j + 1];
                        if (headersRow) {
                            Logger.log(`    Hlavičky pod SOUHRN: ${JSON.stringify(headersRow.slice(0, 15))}`);
                            sazbaColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'SAZBA'
                            );
                            podstrediskoColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'PODSTŘEDISKO'
                            );
                            pocetHodinColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'POČET HODIN'
                            );
                            nocColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'NOC'
                            );
                            vikendColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'VÍKEND'
                            );
                            svatekColIndex = headersRow.findIndex(cell => 
                                typeof cell === 'string' && cell.trim().toUpperCase() === 'SVÁTEK'
                            );
                            Logger.log(`    Indexy sloupců → PODSTŘEDISKO:${podstrediskoColIndex}, SAZBA:${sazbaColIndex}, POČET HODIN:${pocetHodinColIndex}, NOC:${nocColIndex}, VÍKEND:${vikendColIndex}, SVÁTEK:${svatekColIndex}`);
                        } else {
                            Logger.log(`    ❌ Chybí řádek s hlavičkami pod SOUHRN (řádek ${j + 2} neexistuje)`);
                        }
                        break;
                    }
                }

                if (summaryRowIndex === -1) {
                    Logger.log(`    ℹ️ Sekce SOUHRN nenalezena v listu "${sheetName}" - přeskakuji`);
                    return;
                }

                if (podstrediskoColIndex === -1 || sazbaColIndex === -1 || pocetHodinColIndex === -1) {
                    Logger.log(`    ❌ Chybí povinné sloupce v listu "${sheetName}": PODSTŘEDISKO:${podstrediskoColIndex}, SAZBA:${sazbaColIndex}, POČET HODIN:${pocetHodinColIndex}`);
                    return;
                }

                foundData = true;
                let dataRowsProcessed = 0;
                let dataRowsSkipped = 0;

                for (let j = summaryRowIndex + 2; j < sheetData.length; j++) {
                    const row = sheetData[j];
                    let cisloPodstrediska = row[podstrediskoColIndex];
                    const rawCislo = cisloPodstrediska;
                    cisloPodstrediska = config.convertDateToFormat(cisloPodstrediska);

                    const baseSazba = row[sazbaColIndex];
                    const baseHodin = config.roundHours(
                        typeof row[pocetHodinColIndex] === 'number' ? row[pocetHodinColIndex] : 0
                    );
                    const nocHodin = config.roundHours(
                        nocColIndex !== -1 && typeof row[nocColIndex] === 'number' ? row[nocColIndex] : 0
                    );
                    const vikendHodin = config.roundHours(
                        vikendColIndex !== -1 && typeof row[vikendColIndex] === 'number' ? row[vikendColIndex] : 0
                    );
                    const svatekHodin = config.roundHours(
                        svatekColIndex !== -1 && typeof row[svatekColIndex] === 'number' ? row[svatekColIndex] : 0
                    );

                    const hasHours = baseHodin > 0 || nocHodin > 0 || vikendHodin > 0 || svatekHodin > 0;
                    const isValid = typeof baseSazba === 'number' && hasHours && cisloPodstrediska.trim() !== '';

                    if (isValid) {
                        dataRowsProcessed++;
                        const baseCastka = baseSazba * baseHodin;
                        const nocPriplatek = nocHodin * fixedSazbaForPremiums * 0.10;
                        const vikendPriplatek = vikendHodin * fixedSazbaForPremiums * 0.10;
                        const svatekPriplatek = svatekHodin * fixedSazbaForPremiums * 1.00;
                        const castkaWithoutRounding = baseCastka + nocPriplatek + vikendPriplatek + svatekPriplatek;

                        Logger.log(`      ✅ Řádek ${j+1}: podstředisko="${cisloPodstrediska}", sazba=${baseSazba}, hodin=${baseHodin}, noc=${nocHodin}, víkend=${vikendHodin}, svátek=${svatekHodin} → ${castkaWithoutRounding.toFixed(2)} Kč`);

                        workerDataAggregated.set(
                            cisloPodstrediska, 
                            (workerDataAggregated.get(cisloPodstrediska) || 0) + castkaWithoutRounding
                        );
                        
                        if (!workerDataByFile.has(cisloPodstrediska)) {
                            workerDataByFile.set(cisloPodstrediska, []);
                        }
                        workerDataByFile.get(cisloPodstrediska).push({
                            fileName: foundEntry.originalName,
                            castka: castkaWithoutRounding
                        });
                    } else {
                        // Log pouze pro řádky s nějakým obsahem (přeskočit prázdné)
                        const hasAnyContent = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
                        if (hasAnyContent) {
                            dataRowsSkipped++;
                            Logger.log(`      ⏭️ Řádek ${j+1} přeskočen: podstředisko="${cisloPodstrediska}" (raw:"${rawCislo}"), sazba="${baseSazba}"(typ:${typeof baseSazba}), hodin=${baseHodin}, hasHours=${hasHours}`);
                        }
                    }
                }

                Logger.log(`    Výsledek listu "${sheetName}": ${dataRowsProcessed} zpracováno, ${dataRowsSkipped} přeskočeno`);
            });

        } catch (e) {
            Logger.log(`  ❌ Chyba při zpracování souboru ${foundFile.getName()}: ${e.toString()}`);
        }
    }

    if (!foundData) {
        Logger.log(`  ❌ Žádná data nenalezena pro ${workerName} - žádná sekce SOUHRN s validními sloupci nebyla nalezena`);
        
        // Označit řádek jako chybný - soubor existuje, ale nemá validní data
        cellA.setValue('⚠️');
        cellA.setFontColor('#ff9900');
        cellA.setHorizontalAlignment('center');
        
        const rowRange = outputSheet.getRange(i + 1, 1, 1, outputSheet.getLastColumn());
        rowRange.setBackground('#fff2cc'); // Žlutá = soubor nalezen, ale žádná data
        
        config.missingFilesNames.push(`${workerName} - soubor nalezen ale bez dat (${foundEntries.map(e => e.originalName).join(', ')})`);
        return;
    }

    Logger.log(`  Celkem agregovaných podstředisek: ${workerDataAggregated.size}`);

    // ===== ZAOKROUHLOVÁNÍ A ZÁPIS =====
    const subcenterData = [];
    let finalTotalAmount_UNROUNDED = 0;

    workerDataAggregated.forEach((castkaWithoutRounding, cislo) => {
        finalTotalAmount_UNROUNDED += castkaWithoutRounding;
        const decimalPart = castkaWithoutRounding - Math.floor(castkaWithoutRounding);
        subcenterData.push({
            cislo,
            castkaWithoutRounding,
            decimalPart,
            displayCastka: 0
        });
    });

    const targetAmount = Math.ceil(amountInColumnD);
    subcenterData.sort((a, b) => b.decimalPart - a.decimalPart);

    let currentSum = 0;
    subcenterData.forEach(item => {
        item.displayCastka = Math.floor(item.castkaWithoutRounding);
        currentSum += item.displayCastka;
    });

    let diff = targetAmount - currentSum;
    let idx = 0;
    while (diff > 0 && idx < subcenterData.length) {
        const oldValue = subcenterData[idx].displayCastka;
        subcenterData[idx].displayCastka = Math.ceil(subcenterData[idx].castkaWithoutRounding);
        const increase = subcenterData[idx].displayCastka - oldValue;
        diff -= increase;
        currentSum += increase;
        idx++;
    }

    // Zápis do tabulky - ČERVENÉ označení chyb PŘEPÍŠE střídavé obarvení
    for (let k = 0; k < subcenterData.length; k++) {
        const item = subcenterData[k];
        const colPodstredisko = 8 + k * 2;
        const colCastka = colPodstredisko + 1;

        const cellPodstredisko = outputSheet.getRange(i + 1, colPodstredisko);
        cellPodstredisko.setNumberFormat(config.TEXT_FORMAT);
        
        if (typeof item.cislo === 'string' && config.podstrediskoRegex.test(item.cislo.trim())) {
            cellPodstredisko.setValue(item.cislo.trim());
            // NEMĚNÍME pozadí - ponecháme střídavé obarvení
        } else {
            cellPodstredisko.setValue(String(item.cislo).substring(0, 25));
            cellPodstredisko.setBackground('#f4cccc'); // Chyba = červená PŘEPÍŠE střídavé
        }

        const cellCastka = outputSheet.getRange(i + 1, colCastka);
        cellCastka.setNumberFormat(config.CZK_FORMAT);
        cellCastka.setValue(item.displayCastka);
        
        // Přidat komentář k částce, pokud je součtem z více souborů
        if (foundEntries.length > 1) {
            const detailsForPodstredisko = workerDataByFile.get(item.cislo);
            if (detailsForPodstredisko && detailsForPodstredisko.length > 1) {
                let noteText = `Součet pro ${item.cislo}:\n\n`;
                detailsForPodstredisko.forEach((detail, idx) => {
                    noteText += `${idx + 1}. ${detail.fileName}\n`;
                    noteText += `   ${Math.round(detail.castka)} Kč\n\n`;
                });
                noteText += `CELKEM: ${item.displayCastka} Kč`;
                cellCastka.setNote(noteText);
            }
        }
    }

    // Kontrola součtu
    const columnERange = outputSheet.getRange(i + 1, config.totalAmountColSheetIndex);
    columnERange.setNumberFormat(config.CZK_FORMAT);

    const finalSum = subcenterData.reduce((sum, item) => sum + item.displayCastka, 0);
    Logger.log(`  Kontrola součtu: cíl=${targetAmount}, výsledek=${finalSum}, shoda=${finalSum === targetAmount}`);

    if (finalSum !== targetAmount) {
        // Obarvit CELÝ řádek červeně při neshodě částek
        const rowRange = outputSheet.getRange(i + 1, 1, 1, outputSheet.getLastColumn());
        rowRange.setBackground('#f4cccc');
        Logger.log(`  ⚠️ Nesoulad částek u ${workerName}: Cíl=${targetAmount}, Součet=${finalSum}`);
    } else {
        // Pokud se součty SHODUJÍ, můžeme odstranit červené pozadí
        const rowRange = outputSheet.getRange(i + 1, 1, 1, outputSheet.getLastColumn());
        const backgrounds = rowRange.getBackgrounds()[0];
        
        // Zkontrolovat, jestli je řádek červený
        const isRedRow = backgrounds.some(bg => bg === '#f4cccc');
        
        if (isRedRow) {
            Logger.log(`  ✅ Součty se shodují u ${workerName}, odstraňuji červené pozadí`);
            // Obnovit správné formátování
            rowRange.setBackground(null);
            
            // Obnovit střídavé obarvení sloupců od H dále
            const lastCol = outputSheet.getLastColumn();
            if (lastCol >= 8) {
                let col = 8;
                let pairIndex = 0;
                
                while (col <= lastCol) {
                    if (pairIndex % 2 === 0) {
                        const range = outputSheet.getRange(i + 1, col, 1, 2);
                        range.setBackground('#f3f3f3');
                    } else {
                        const range = outputSheet.getRange(i + 1, col, 1, 2);
                        range.setBackground(null);
                    }
                    col += 2;
                    pairIndex++;
                }
            }
        }
    }
}

// ===== FUNKCE PRO RELOAD AKTUÁLNÍHO ŘÁDKU =====
function reloadCurrentRow() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const currentRow = sheet.getActiveCell().getRow();
    
    if (currentRow === 1) {
        SpreadsheetApp.getUi().alert('Nelze aktualizovat hlavičku. Vyberte řádek s daty.');
        return;
    }
    
    // Zjistit ID adresáře z názvu listu
    let sheetName = sheet.getName();
    const adresarSheet = ss.getSheetByName('Adresář');
    
    if (!adresarSheet) {
        SpreadsheetApp.getUi().alert('List "Adresář" nebyl nalezen.');
        return;
    }
    
    const adresarData = adresarSheet.getDataRange().getValues();
    
    // Odstranit " (číslo)" z názvu listu pokud existuje
    const baseSheetName = sheetName.replace(/\s\(\d+\)$/, '');
    Logger.log(`Název listu: "${sheetName}", Základní název: "${baseSheetName}"`);
    
    let folderId = null;
    for (let i = 1; i < adresarData.length; i++) {
        const monthName = String(adresarData[i][0]).trim();
        
        // Kontrola přesné shody nebo shody se základním názvem
        if (sheetName === monthName || baseSheetName === monthName) {
            folderId = String(adresarData[i][1]).trim();
            Logger.log(`✅ Nalezena shoda: měsíc="${monthName}", ID="${folderId}"`);
            break;
        }
    }
    
    if (!folderId) {
        SpreadsheetApp.getUi().alert(
            `Pro list "${sheetName}" není nastaven adresář.\n\n` +
            `Zkontrolujte, zda v listu "Adresář" existuje řádek s názvem:\n"${baseSheetName}"`
        );
        Logger.log(`❌ ID adresáře nenalezeno pro list "${sheetName}" (základ: "${baseSheetName}")`);
        return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Aktualizuji řádek ${currentRow}...`, 'Reload', -1);
    
    // ===== VYMAZAT OBSAH A FORMÁTOVÁNÍ =====
    const lastCol = sheet.getLastColumn();
    
    // Vymazat pozadí pro celý řádek (kromě hlavičky)
    sheet.getRange(currentRow, 1, 1, lastCol).setBackground(null);
    
    // Vymazat sloupec A (odkaz) - obsah + barvu textu
    sheet.getRange(currentRow, 1).clearContent().setFontColor(null);
    
    // Vymazat od sloupce H dále - obsah
    if (lastCol >= 8) {
        sheet.getRange(currentRow, 8, 1, lastCol - 7).clearContent();
    }
    
    // ===== NEJDŘÍVE NASTAVIT STŘÍDAVÉ OBARVENÍ =====
    SpreadsheetApp.getActiveSpreadsheet().toast('Nastavuji formátování...', 'Příprava', -1);
    
    if (lastCol >= 8) {
        let col = 8;
        let pairIndex = 0;
        
        while (col <= lastCol) {
            if (pairIndex % 2 === 0) {
                // Dvojice H-I, L-M, P-Q... - šedá
                const range = sheet.getRange(currentRow, col, 1, 2);
                range.setBackground('#f3f3f3');
            } else {
                // Dvojice J-K, N-O, R-S... - bílá
                const range = sheet.getRange(currentRow, col, 1, 2);
                range.setBackground(null);
            }
            col += 2;
            pairIndex++;
        }
    }
    
    // ===== TEPRVE NYNÍ ZPRACOVAT ŘÁDEK (červené označení chyb přepíše střídavé) =====
    SpreadsheetApp.getActiveSpreadsheet().toast(`Zpracovávám data...`, 'Zpracování', -1);
    
    const data = sheet.getDataRange().getValues();
    
    const folder = DriveApp.getFolderById(folderId);
    const filesIterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const fileCache = [];
    
    const normalizeName = (name) => {
        return name.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toUpperCase();
    };
    
    while (filesIterator.hasNext()) {
        const file = filesIterator.next();
        fileCache.push({
            name: normalizeName(file.getName()),
            originalName: file.getName(),
            file: file,
            processed: false
        });
    }

    Logger.log(`Soubory v cache pro reload (${fileCache.length}):`);
    fileCache.forEach(entry => Logger.log(`  - "${entry.name}" (original: "${entry.originalName}")`));
    
    const config = {
        prijemniColIndex: 1,
        jmenoColIndex: 2,
        druhSmlouvyColIndex: 3,
        totalAmountColIndex: 4,
        totalAmountColSheetIndex: 5,
        CZK_FORMAT: '#,##0 "Kč"',
        TEXT_FORMAT: '@',
        podstrediskoRegex: /^\d{2}-\d{2}-\d{2}(\/\d+)?$/,
        normalizeName: normalizeName,
        roundHours: (value) => Math.round(value * 100) / 100,
        convertDateToFormat: (value) => {
            if (value instanceof Date) {
                const day = String(value.getDate()).padStart(2, '0');
                const month = String(value.getMonth() + 1).padStart(2, '0');
                const year = String(value.getFullYear()).slice(-2);
                return `${day}-${month}-${year}`;
            }
            return String(value);
        },
        detectContractType: (contractText) => {
            const text = String(contractText).trim().toLowerCase();
            if (text.includes('dohoda o provedení práce') || text.includes('dpp')) return 'DPP';
            if (text.includes('dohoda o pracovní činnosti') || text.includes('dpč')) return 'DPČ';
            return null;
        },
        missingFilesNames: []
    };
    
    // Zpracovat řádek - červené buňky PŘEPÍŠÍ střídavé obarvení tam, kde je chyba
    processRow(sheet, currentRow - 1, data, fileCache, config);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Řádek aktualizován!', 'Hotovo', 3);
    Logger.log(`✅ Řádek ${currentRow} byl úspěšně aktualizován.`);
}