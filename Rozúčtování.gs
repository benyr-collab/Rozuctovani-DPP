function startSumarizace() {
  showMonthSelectionDialog();
}

function showMonthSelectionDialog() {
  Logger.log('=== showMonthSelectionDialog START ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  Logger.log('Počet listů v sešitu: ' + sheets.length);
  
  // České názvy měsíců
  const czechMonths = [
    'Leden', 'Únor', 'Březen', 'Duben', 'Květen', 'Červen',
    'Červenec', 'Srpen', 'Září', 'Říjen', 'Listopad', 'Prosinec'
  ];
  
  // Najít všechny listy s měsíčními názvy
  const monthSheets = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    Logger.log('Kontroluji list: ' + name + ' - je měsíc? ' + czechMonths.includes(name));
    if (czechMonths.includes(name)) {
      monthSheets.push(name);
    }
  });
  
  Logger.log('Nalezené měsíční listy: ' + JSON.stringify(monthSheets));
  
  if (monthSheets.length === 0) {
    Logger.log('Žádné měsíční listy nenalezeny!');
    SpreadsheetApp.getUi().alert('Nebyly nalezeny žádné měsíční listy (Leden, Únor, Březen, ...).');
    return;
  }
  
  // Seřadit měsíce podle kalendářního pořadí
  monthSheets.sort((a, b) => {
    return czechMonths.indexOf(a) - czechMonths.indexOf(b);
  });
  
  Logger.log('Seřazené měsíce: ' + JSON.stringify(monthSheets));
  
  // Vytvořit HTML dialog s radio buttons
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 10px; }
          label { display: block; padding: 6px 0; }
          .buttons { margin-top: 16px; text-align: right; }
          button { margin-left: 8px; padding: 6px 14px; }
        </style>
      </head>
      <body>
        <h2>Vyberte měsíc pro sumarizaci</h2>
        <div id="months">
          ${monthSheets.map((m, idx) => `<label><input type="radio" name="month" value="${m}"${idx === 0 ? ' checked' : ''}> ${m}</label>`).join('')}
        </div>
        <div class="buttons">
          <button onclick="google.script.host.close()">Zrušit</button>
          <button onclick="submitMonth()">Spustit sumarizaci</button>
        </div>
        <script>
          function submitMonth() {
            var selected = document.querySelector('input[name="month"]:checked');
            if (!selected) { return; }
            google.script.run.withSuccessHandler(google.script.host.close).processSumarizaceWithCheck(selected.value);
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Výběr měsíce');
  
  Logger.log('=== showMonthSelectionDialog END ===');
}

function processSumarizaceWithCheck(selectedMonth) {
  Logger.log('processSumarizaceWithCheck zavoláno s: ' + selectedMonth);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rozuctovaniSheetName = `Rozúčtování ${selectedMonth}`;
  const kontrolaSheetName = `Kontrola ${selectedMonth}`;
  
  const existingRozuctovani = ss.getSheetByName(rozuctovaniSheetName);
  const existingKontrola = ss.getSheetByName(kontrolaSheetName);
  
  let shouldOverwrite = true;
  
  if (existingRozuctovani || existingKontrola) {
    const ui = SpreadsheetApp.getUi();
    let message = 'Následující listy již existují:\n';
    if (existingRozuctovani) message += `- ${rozuctovaniSheetName}\n`;
    if (existingKontrola) message += `- ${kontrolaSheetName}\n`;
    message += '\nChcete je přepsat nebo vytvořit nové?';
    
    const response = ui.alert(
      'Existující listy', 
      message, 
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      shouldOverwrite = true;
      Logger.log('Uživatel zvolil přepsání existujících listů');
    } else if (response === ui.Button.NO) {
      shouldOverwrite = false;
      Logger.log('Uživatel zvolil vytvoření nových listů');
    } else {
      Logger.log('Uživatel zrušil operaci');
      return;
    }
  }
  
  summarizeByWorkCenter(selectedMonth, shouldOverwrite);
}

function saveSelectedMonth(monthName) {
  Logger.log('saveSelectedMonth zavoláno s: ' + monthName);
  Logger.log('Typ: ' + typeof monthName);
  PropertiesService.getUserProperties().setProperty('SELECTED_MONTH', monthName);
  Logger.log('Měsíc uložen: ' + monthName);
}

function processSelectedMonth() {
  const selectedMonth = PropertiesService.getUserProperties().getProperty('SELECTED_MONTH');
  Logger.log('processSelectedMonth - načtený měsíc: ' + selectedMonth);
  
  if (!selectedMonth) {
    SpreadsheetApp.getUi().alert('Chyba: Nepodařilo se načíst vybraný měsíc.');
    return;
  }
  
  summarizeByWorkCenter(selectedMonth);
}

function summarizeByWorkCenter(selectedMonth, shouldOverwrite) {
  Logger.log('=== DEBUG START ===');
  Logger.log('Typ selectedMonth: ' + typeof selectedMonth);
  Logger.log('Hodnota selectedMonth: ' + selectedMonth);
  Logger.log('selectedMonth je undefined: ' + (selectedMonth === undefined));
  Logger.log('selectedMonth je null: ' + (selectedMonth === null));
  Logger.log('shouldOverwrite: ' + shouldOverwrite);
  
  // Pokud je selectedMonth undefined nebo null, zkusíme zobrazit všechny dostupné listy
  if (!selectedMonth) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    Logger.log('Dostupné listy v sešitu:');
    allSheets.forEach(sheet => {
      Logger.log('  - ' + sheet.getName());
    });
    SpreadsheetApp.getUi().alert('Chyba: Nebyl vybrán žádný měsíc. Zkontrolujte prosím logy.');
    return;
  }
  
  Logger.log(`Spouštění funkce summarizeByWorkCenter() pro měsíc: ${selectedMonth}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(selectedMonth);
    const workCenterSheet = ss.getSheetByName('Podstřediska');
    
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(`List s názvem "${selectedMonth}" nebyl nalezen.`);
      return;
    }

    if (!workCenterSheet) {
      SpreadsheetApp.getUi().alert('List s názvem "Podstřediska" nebyl nalezen. Zkontrolujte prosím název listu.');
      return;
    }

    workCenterSheet.getRange('A:A').setNumberFormat('@');
    Logger.log('Formát sloupce A v listu Podstřediska byl nastaven na text.');
    
    const workCenterData = workCenterSheet.getDataRange().getValues();
    const workCenterMap = new Map();
    if (workCenterData.length > 1) {
      for (let i = 1; i < workCenterData.length; i++) {
        const key = String(workCenterData[i][0]).trim();
        const value = String(workCenterData[i][1]).trim();
        if (key && value) {
          workCenterMap.set(key, value);
        }
      }
    }

    sourceSheet.getRange('A:ZZ').setNumberFormat('@');
    Logger.log('Formát na text nastaven pro sloupce s kódy podstředisek.');
    
    const data = sourceSheet.getDataRange().getValues();
    Logger.log(`Data byla úspěšně načtena z listu "${selectedMonth}". Počet řádků: ${data.length}`);
    
    if (data.length < 2) {
      Logger.log('V listu nejsou žádná data k zpracování (méně než 2 řádky).');
      SpreadsheetApp.getUi().alert(`V listu "${selectedMonth}" nejsou žádná data k zpracování.`);
      return;
    }

    const header = data[0];
    Logger.log('Záhlaví tabulky: ' + JSON.stringify(header));

    const prijemniIndex = header.indexOf('Příjmení');
    const jmenoIndex = header.indexOf('Jméno');
    const druhPpIndex = header.indexOf('Druh pracovního poměru'); 
    const hrubaMzdaIndex = header.indexOf('Hrubá mzda');
    const spIndex = header.indexOf('SP');
    const zpIndex = header.indexOf('ZP');

    let totalHrubaMzdaOriginal = 0;
    let totalSpOriginal = 0;
    let totalZpOriginal = 0;
    
    sourceSheet.getRange(2, hrubaMzdaIndex + 1, sourceSheet.getLastRow() - 1, 3).setNumberFormat('#,##0.00 "Kč"');
    Logger.log('Formát měny nastaven pro sloupce Hrubá mzda, SP, ZP ve zdrojovém listu.');

    for (let i = 1; i < data.length; i++) {
      const hrubaMzda = String(data[i][hrubaMzdaIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.');
      const sp = String(data[i][spIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.');
      const zp = String(data[i][zpIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.');
      
      totalHrubaMzdaOriginal += parseFloat(hrubaMzda) || 0;
      totalSpOriginal += parseFloat(sp) || 0;
      totalZpOriginal += parseFloat(zp) || 0;
    }
    
    Logger.log(`Celková Hrubá mzda (${selectedMonth}): ${totalHrubaMzdaOriginal}`);
    Logger.log(`Celková SP (${selectedMonth}): ${totalSpOriginal}`);
    Logger.log(`Celková ZP (${selectedMonth}): ${totalZpOriginal}`);

    // Získat měsíc a rok pro názvy
    const monthYear = getMonthYearFromSheet(data, selectedMonth);
    
    // Název listů s měsícem
    const rozuctovaniSheetName = `Rozúčtování ${selectedMonth}`;
    const kontrolaSheetName = `Kontrola ${selectedMonth}`;

    let rozuctovaniSheet = ss.getSheetByName(rozuctovaniSheetName);
    
    if (!shouldOverwrite && rozuctovaniSheet) {
      // Vytvořit nový list s časovým razítkem
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const newRozuctovaniName = `${rozuctovaniSheetName} (${timestamp})`;
      rozuctovaniSheet = ss.insertSheet(newRozuctovaniName);
      Logger.log(`List "${newRozuctovaniName}" byl vytvořen.`);
    } else if (!rozuctovaniSheet) {
      rozuctovaniSheet = ss.insertSheet(rozuctovaniSheetName);
      Logger.log(`List "${rozuctovaniSheetName}" byl vytvořen.`);
    } else {
      const existingFilter = rozuctovaniSheet.getFilter();
      if (existingFilter) {
          existingFilter.remove();
          Logger.log('Existující filtr byl odstraněn.');
      }
      rozuctovaniSheet.clear(); 
      Logger.log(`List "${rozuctovaniSheetName}" byl vyčištěn.`);
    }

    rozuctovaniSheet.getRange('A:A').setNumberFormat('@');
    Logger.log(`Formát sloupce A v listu "${rozuctovaniSheetName}" byl nastaven na text.`);
    
    const summaryMap = new Map();
    const employeeDetails = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const prijemni = String(row[prijemniIndex]).trim();
      const jmeno = String(row[jmenoIndex]).trim();
      const druhPpRaw = String(row[druhPpIndex]).trim();

      // detectContractType() je sdílená funkce (Naplnění dat.gs) - rozpozná DPP/DPČ
      // bez ohledu na diakritiku, velikost písmen nebo formu Unicode normalizace,
      // takže se nespoléhá na přesnou shodu celé fráze jako dřívější druhPpMap.
      const detectedDruhPp = detectContractType(druhPpRaw);
      let druhPp = detectedDruhPp || druhPpRaw.toLowerCase();

      const fullName = `${prijemni} ${jmeno}`.trim();

      if (!prijemni && !jmeno) {
        Logger.log(`Přeskakuji řádek ${i + 1}, protože neobsahuje jméno.`);
        continue;
      }
      
      Logger.log(`Zpracovávám řádek pro: ${fullName}`);

      const hrubaMzdaOriginalValue = parseFloat(String(row[hrubaMzdaIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.')) || 0;
      const totalSp = parseFloat(String(row[spIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.')) || 0;
      const totalZp = parseFloat(String(row[zpIndex]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.')) || 0;
      
      let partsTotal = 0;
      const partsMap = new Map();
      
      const spDistributed = [];
      const zpDistributed = [];

      const startColIndex = zpIndex + 1;

      for (let col = startColIndex; col < row.length; col += 2) { 
        const podstredisko = String(row[col]).trim();
        const castkaValue = String(row[col + 1]).replace(/Kč/g, '').replace(/\s/g, '').replace(',', '.');
        const castka = parseFloat(castkaValue);

        if (podstredisko !== '' && !isNaN(castka) && castka > 0) {
          partsMap.set(podstredisko, (partsMap.get(podstredisko) || 0) + castka); 
          partsTotal += castka;
        }
      }
      
      if (partsTotal === 0) {
        Logger.log(`Přeskakuji řádek pro ${fullName}, protože nemá přiřazenou částku k podstředisku.`);
        continue;
      }

      const hrubaMzdaRozuctovano = partsTotal;

      const hrubaMzdaDetails = [];
      let totalSpRozuctovano = 0;
      let totalZpRozuctovano = 0;
      
      partsMap.forEach((partAmount, podstredisko) => {
        const ratio = partAmount / partsTotal;
        const spContribution = Math.round(totalSp * ratio);
        const zpContribution = Math.round(totalZp * ratio);

        hrubaMzdaDetails.push(`${podstredisko}: ${partAmount} Kč`);
        spDistributed.push(`${podstredisko}: ${spContribution} Kč`);
        zpDistributed.push(`${podstredisko}: ${zpContribution} Kč`);
        
        totalSpRozuctovano += spContribution;
        totalZpRozuctovano += zpContribution;
      });

      employeeDetails.push([
        fullName, 
        hrubaMzdaOriginalValue,
        hrubaMzdaDetails.join(', '),
        hrubaMzdaRozuctovano,
        hrubaMzdaRozuctovano - hrubaMzdaOriginalValue,
        totalSp,
        spDistributed.join(', '),
        totalSpRozuctovano,
        totalSpRozuctovano - totalSp,
        totalZp,
        zpDistributed.join(', '),
        totalZpRozuctovano,
        totalZpRozuctovano - totalZp
      ]);
      
      partsMap.forEach((partAmount, podstredisko) => {
        const uniqueKey = `${podstredisko}_${druhPp}`;

        if (!summaryMap.has(uniqueKey)) {
          summaryMap.set(uniqueKey, { 
            hrubaMzda: { totalAmount: 0, names: new Set() },
            sp: { totalAmount: 0, names: new Set() },
            zp: { totalAmount: 0, names: new Set() }
          });
          Logger.log(`Vytvořena nová položka v mapě pro podstředisko/PP: ${uniqueKey}`);
        }
        
        const currentSummary = summaryMap.get(uniqueKey);
        const ratio = partAmount / partsTotal;

        const spContribution = Math.round(totalSp * ratio);
        const zpContribution = Math.round(totalZp * ratio);

        currentSummary.hrubaMzda.totalAmount += partAmount;
        currentSummary.sp.totalAmount += spContribution;
        currentSummary.zp.totalAmount += zpContribution;
        
        currentSummary.hrubaMzda.names.add(fullName);

        if (spContribution > 0) {
          currentSummary.sp.names.add(fullName);
        }

        if (zpContribution > 0) {
          currentSummary.zp.names.add(fullName);
        }
      });
    }

    Logger.log('Všechny řádky byly zpracovány. Vytváří se výstupní data.');
    
    const outputData = [];
    const headers = ['Druh PP', 'Druh', 'Částka', 'Jména', 'Podstředisko', 'Název podstřediska']; 
    
    const sortedKeys = Array.from(summaryMap.keys());
    
    sortedKeys.sort((a, b) => {
      const [podstrediskoA, druhPpA] = a.split('_');
      const [podstrediskoB, druhPpB] = b.split('_');

      if (druhPpA === 'DPČ' && druhPpB === 'DPP') return -1;
      if (druhPpA === 'DPP' && druhPpB === 'DPČ') return 1;

      if (podstrediskoA < podstrediskoB) return -1;
      if (podstrediskoA > podstrediskoB) return 1;
      
      return 0;
    });

    sortedKeys.forEach(uniqueKey => {
      const summary = summaryMap.get(uniqueKey);
      const [podstredisko, druhPp] = uniqueKey.split('_');
      const nazevPodstrediska = workCenterMap.get(String(podstredisko)) || 'N/A';
      
      const getFormattedNames = (namesSet) => {
        const prijemniCounts = new Map();
        namesSet.forEach(fullName => {
          const prijemni = fullName.split(' ')[0];
          prijemniCounts.set(prijemni, (prijemniCounts.get(prijemni) || 0) + 1);
        });

        const formattedNames = [];
        namesSet.forEach(fullName => {
          const parts = fullName.split(' ');
          const prijemni = parts[0];
          const jmeno = parts.length > 1 ? parts[1] : '';
          if (prijemniCounts.get(prijemni) > 1 && jmeno) {
            formattedNames.push(`${prijemni} ${jmeno.charAt(0)}.`);
          } else {
            formattedNames.push(prijemni);
          }
        });
        return formattedNames.join(', ');
      };
      
      if (summary.hrubaMzda.totalAmount > 0) {
        const namesString = `Mzdy ${monthYear} - ${druhPp} - ${getFormattedNames(summary.hrubaMzda.names)}`;
        outputData.push([
          druhPp, 
          '521100', 
          summary.hrubaMzda.totalAmount, 
          namesString,
          podstredisko, 
          nazevPodstrediska 
        ]);
      }
        
      if (summary.sp.totalAmount > 0) {
        const namesString = `Mzdy ${monthYear} - ${druhPp} - ${getFormattedNames(summary.sp.names)}`;
        outputData.push([
          druhPp,
          '524100',
          summary.sp.totalAmount,
          namesString,
          podstredisko,
          nazevPodstrediska
        ]);
      }
        
      if (summary.zp.totalAmount > 0) {
        const namesString = `Mzdy ${monthYear} - ${druhPp} - ${getFormattedNames(summary.zp.names)}`;
        outputData.push([
          druhPp,
          '524200',
          summary.zp.totalAmount,
          namesString,
          podstredisko,
          nazevPodstrediska
        ]);
      }
    });

    if (outputData.length > 0) {
      // Přidat nadpisový řádek s názvem měsíce
      const monthYearTitle = `${selectedMonth} - ${monthYear}`;
      rozuctovaniSheet.getRange(1, 1, 1, headers.length).merge().setValue(monthYearTitle);
      rozuctovaniSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
      
      // Přidat záhlaví na druhý řádek
      rozuctovaniSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
      
      // Data začínají od třetího řádku
      const dataRange = rozuctovaniSheet.getRange(3, 1, outputData.length, outputData[0].length);
      dataRange.setValues(outputData);
      
      const filterRange = rozuctovaniSheet.getRange(2, 1, rozuctovaniSheet.getLastRow() - 1, headers.length);
      filterRange.createFilter();

      rozuctovaniSheet.autoResizeColumns(1, headers.length);
      rozuctovaniSheet.getRange('A:G').setNumberFormat('@'); 
      rozuctovaniSheet.getRange('C:C').setNumberFormat('#,##0.00 "Kč"'); 
      
      Logger.log(`Data byla úspěšně zapsána do listu ${rozuctovaniSheet.getName()}.`);
      
      const lastRow = rozuctovaniSheet.getLastRow();
      
      const sumRow = rozuctovaniSheet.getRange(lastRow + 2, 1, 1, 6); 
      
      sumRow.setValues([['Celkem (dle filtru)', '', `=SUBTOTAL(9; C3:C${lastRow})`, '', '', '']]);
      
      sumRow.getCell(1, 1).setFontWeight('bold');
      sumRow.getCell(1, 3).setFontWeight('bold');
      sumRow.getCell(1, 3).setNumberFormat('#,##0.00 "Kč"');
      sumRow.setBackground('#e6e6e6');
      
    } else {
      Logger.log(`Žádná data nebyla nalezena pro zápis do listu ${rozuctovaniSheet.getName()}.`);
    }
    
    // Vytvoření kontrolního listu
    let kontrolaSheet = ss.getSheetByName(kontrolaSheetName);
    
    if (!shouldOverwrite && kontrolaSheet) {
      // Vytvořit nový list s časovým razítkem
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const newKontrolaName = `${kontrolaSheetName} (${timestamp})`;
      kontrolaSheet = ss.insertSheet(newKontrolaName);
      Logger.log(`List "${newKontrolaName}" byl vytvořen.`);
    } else if (!kontrolaSheet) {
      kontrolaSheet = ss.insertSheet(kontrolaSheetName);
    } else {
      kontrolaSheet.clear();
    }
    
    // Přidat nadpisový řádek s názvem měsíce
    const monthYearTitle = `${selectedMonth} - ${monthYear} - KONTROLA`;
    kontrolaSheet.getRange(1, 1, 1, 13).merge().setValue(monthYearTitle);
    kontrolaSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    
    const kontrolaHeaders = [
      'Jméno', 
      'Hrubá mzda (původní)',
      'Rozúčtování HM',
      'Součet HM (rozúčtované)',
      'Rozdíl HM',
      'SP (původní)',
      'Rozúčtování SP',
      'Součet SP (rozúčtované)',
      'Rozdíl SP',
      'ZP (původní)',
      'Rozúčtování ZP',
      'Součet ZP (rozúčtované)',
      'Rozdíl ZP'
    ];
    kontrolaSheet.getRange(2, 1, 1, kontrolaHeaders.length).setValues([kontrolaHeaders]);
    
    if (employeeDetails.length > 0) {
      kontrolaSheet.getRange(3, 1, employeeDetails.length, employeeDetails[0].length).setValues(employeeDetails);
    }

    kontrolaSheet.autoResizeColumns(1, kontrolaHeaders.length);
    kontrolaSheet.getRange('B:B').setNumberFormat('#,##0.00 "Kč"');
    kontrolaSheet.getRange('D:E').setNumberFormat('#,##0.00 "Kč"');
    kontrolaSheet.getRange('F:F').setNumberFormat('#,##0.00 "Kč"');
    kontrolaSheet.getRange('H:I').setNumberFormat('#,##0.00 "Kč"');
    kontrolaSheet.getRange('J:J').setNumberFormat('#,##0.00 "Kč"');
    kontrolaSheet.getRange('L:M').setNumberFormat('#,##0.00 "Kč"');
    
    // Podmíněné formátování
    const ruleHM = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#d9ead3')
      .setRanges([kontrolaSheet.getRange(`E3:E${kontrolaSheet.getLastRow()}`)])
      .build();

    const ruleHMerror = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(0)
      .setBackground('#f4cccc')
      .setRanges([kontrolaSheet.getRange(`E3:E${kontrolaSheet.getLastRow()}`)])
      .build();

    const ruleSP = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#d9ead3')
      .setRanges([kontrolaSheet.getRange(`I3:I${kontrolaSheet.getLastRow()}`)])
      .build();

    const ruleSPerror = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(0)
      .setBackground('#f4cccc')
      .setRanges([kontrolaSheet.getRange(`I3:I${kontrolaSheet.getLastRow()}`)])
      .build();
    
    const ruleZP = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#d9ead3')
      .setRanges([kontrolaSheet.getRange(`M3:M${kontrolaSheet.getLastRow()}`)])
      .build();

    const ruleZPerror = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(0)
      .setBackground('#f4cccc')
      .setRanges([kontrolaSheet.getRange(`M3:M${kontrolaSheet.getLastRow()}`)])
      .build();

    const rules = kontrolaSheet.getConditionalFormatRules();
    rules.push(ruleHM, ruleHMerror, ruleSP, ruleSPerror, ruleZP, ruleZPerror);
    kontrolaSheet.setConditionalFormatRules(rules);

    Logger.log('Podmíněné formátování bylo aplikováno na sloupce rozdílů.');

    // Přesun listů do správného pořadí
    arrangeSheets(ss, selectedMonth, rozuctovaniSheet.getName(), kontrolaSheet.getName());

    // Křížová kontrola celkových součtů
    let totalHrubaMzdaRozuctovani = 0;
    let totalSpRozuctovani = 0;
    let totalZpRozuctovani = 0;
    
    summaryMap.forEach(summary => {
      totalHrubaMzdaRozuctovani += summary.hrubaMzda.totalAmount;
      totalSpRozuctovani += summary.sp.totalAmount;
      totalZpRozuctovani += summary.zp.totalAmount;
    });

    const hrubaMzdaDifference = Math.abs(totalHrubaMzdaOriginal - totalHrubaMzdaRozuctovani);
    const spDifference = Math.abs(totalSpOriginal - totalSpRozuctovani);
    const zpDifference = Math.abs(totalZpOriginal - totalZpRozuctovani);
    
    Logger.log(`Celková Hrubá mzda (${selectedMonth}): ${totalHrubaMzdaOriginal}`);
    Logger.log(`Celková Hrubá mzda (Rozúčtování): ${totalHrubaMzdaRozuctovani}`);
    Logger.log(`Celková SP (${selectedMonth}): ${totalSpOriginal}`);
    Logger.log(`Celková SP (Rozúčtování): ${totalSpRozuctovani}`);
    Logger.log(`Celková ZP (${selectedMonth}): ${totalZpOriginal}`);
    Logger.log(`Celková ZP (Rozúčtování): ${totalZpRozuctovani}`);
    
    if (hrubaMzdaDifference > 0.01 || spDifference > 0.01 || zpDifference > 0.01) {
      Logger.log('CHYBA KONTROLY SOUČTŮ!');
      Logger.log(`Rozdíl Hrubá mzda: ${hrubaMzdaDifference}`);
      Logger.log(`Rozdíl SP: ${spDifference}`);
      Logger.log(`Rozdíl ZP: ${zpDifference}`);
      SpreadsheetApp.getUi().alert(`Upozornění: Součty na listu "${rozuctovaniSheet.getName()}" neodpovídají součtům ze zdrojového listu. Zkontrolujte logy a nově vytvořený list "${kontrolaSheet.getName()}" pro detaily.`);
    } else {
      SpreadsheetApp.getUi().alert(`Sumarizace pro měsíc ${selectedMonth} byla dokončena.\n\nByly vytvořeny listy:\n- ${rozuctovaniSheet.getName()}\n- ${kontrolaSheet.getName()}\n\nSoučty se shodují.`);
    }

  } catch (e) {
    Logger.log(`Chyba při spuštění skriptu: ${e.toString()}`);
    SpreadsheetApp.getUi().alert('Došlo k chybě. Zkontrolujte logy.\n\n' + e.toString());
  }
}

function arrangeSheets(ss, monthSheetName, rozuctovaniSheetName, kontrolaSheetName) {
  try {
    const monthSheet = ss.getSheetByName(monthSheetName);
    const rozuctovaniSheet = ss.getSheetByName(rozuctovaniSheetName);
    const kontrolaSheet = ss.getSheetByName(kontrolaSheetName);
    
    if (!monthSheet || !rozuctovaniSheet || !kontrolaSheet) {
      Logger.log('Některý z listů nebyl nalezen pro přesunutí.');
      return;
    }
    
    // Získat aktuální pozice všech listů
    const allSheets = ss.getSheets();
    let targetPosition = 0;
    
    // Najít pozici měsíčního listu
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName() === monthSheetName) {
        targetPosition = i;
        break;
      }
    }
    
    // Přesunout listy do pořadí: měsíc, rozúčtování, kontrola
    ss.setActiveSheet(monthSheet);
    ss.moveActiveSheet(targetPosition + 1);
    
    ss.setActiveSheet(rozuctovaniSheet);
    ss.moveActiveSheet(targetPosition + 2);
    
    ss.setActiveSheet(kontrolaSheet);
    ss.moveActiveSheet(targetPosition + 3);
    
    // Vrátit se na list rozúčtování
    ss.setActiveSheet(rozuctovaniSheet);
    
    Logger.log(`Listy byly přesunuty do pořadí: ${monthSheetName}, ${rozuctovaniSheetName}, ${kontrolaSheetName}`);
    
  } catch (e) {
    Logger.log(`Chyba při přesouvání listů: ${e.toString()}`);
  }
}

function getMonthYearFromSheet(data, monthName) {
  const monthNames = {
      'leden': 1, 'únor': 2, 'březen': 3, 'duben': 4, 'květen': 5, 'červen': 6,
      'červenec': 7, 'srpen': 8, 'září': 9, 'říjen': 10, 'listopad': 11, 'prosinec': 12
  };
  
  // Hledání indexu sloupce 'Měsíc' v hlavičce
  const header = data.length > 0 ? data[0] : [];
  let monthIndex = header.indexOf('Měsíc');
  
  if (monthIndex === -1) {
    monthIndex = 0;
  }
  
  if (data.length > 1 && data[1][monthIndex]) {
    const dateValue = data[1][monthIndex];
    
    // Zpracování Date objektu
    if (dateValue instanceof Date) {
      const month = dateValue.getMonth() + 1;
      const year = dateValue.getFullYear();
      return `${Utilities.formatString("%02d", month)}/${year}`;
    } 
    
    // Zpracování řetězce
    if (typeof dateValue === 'string') {
      const monthText = String(dateValue).trim().toLowerCase();
      const parts = monthText.split('/');
      
      // Formát DD/MM/RRRR
      if (parts.length === 3) {
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        if (!isNaN(month) && !isNaN(year)) {
          return `${Utilities.formatString("%02d", month)}/${year}`;
        }
      }
      
      // Slovní měsíc
      const monthNumber = monthNames[monthText];
      if (monthNumber) {
        const currentYear = new Date().getFullYear();
        const year = (monthText === 'prosinec' && new Date().getMonth() === 0) ? currentYear - 1 : currentYear; 
        return `${Utilities.formatString("%02d", monthNumber)}/${year}`;
      }
    }
  }
  
  // Fallback - použít název listu a aktuální rok
  const monthNameLower = monthName.toLowerCase();
  const monthNumber = monthNames[monthNameLower];
  if (monthNumber) {
    const currentYear = new Date().getFullYear();
    return `${Utilities.formatString("%02d", monthNumber)}/${currentYear}`;
  }
  
  return 'Neznámý měsíc/rok';
}