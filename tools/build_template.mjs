import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

// Development utility. Runtime users need only Python + openpyxl.
const [jsonPath, output, previewDir] = process.argv.slice(2);
const spec = JSON.parse(await fs.readFile(jsonPath, 'utf8'));
const wb = Workbook.create();
const types = Object.fromEntries(Object.entries(spec.types).map(([a,b])=>[b,a]));
const notes = {
  Settings: ['Укажите первый месяц полной истории и последний месяц отчёта.', 'Обе даты — первое число месяца. Начальные остатки нулевые.'],
  Products: ['Один код SKU — один товар. В покупках и продажах SKU может повторяться.', 'Пример: DEMO-A  |  Наименование товара  |  шт.'],
  Counterparties: ['Внесите поставщиков, покупателей, брокеров и перевозчиков по одному разу.', 'Названия выбираются в журналах из этого справочника.'],
  Purchases: ['Одна строка — один товар. Номер декларации / e-qaimə повторяйте для всех товаров документа.', 'Сумма всей строки без НДС. Документ — инвойс или e-qaimə. Код строки заполнится при расчёте.'],
  Expenses: ['Одна строка — один расход. Выберите номер поставки, укажите исполнителя и документ.', 'Расходы распределяются по стоимости товаров. Поздние расходы пересчитывают исходное поступление.'],
  Sales: ['Выберите товар и операцию. Для продажи обязательны покупатель и выручка, включая 0.', 'Для списания выручка 0 или пусто. Код строки заполнится при расчёте.'],
  Returns: ['Выберите тип возврата и код исходной покупки или продажи. Количество всегда положительное.', 'Сначала рассчитайте исходную операцию, чтобы появился её код. Для покупателя укажите уменьшение выручки.'],
};
const widths = {
  Date:14, 'Document Date':16, 'Shipment ID':30, Supplier:26, Counterparty:26,
  Document:23, SKU:23, Quantity:16, Currency:12, 'Amount FCY':19, 'FX to AZN':16,
  'Purchase Type':17, 'Line ID':17, 'Expense ID':17, 'Return ID':17,
  'Expense Type':24, Customer:26, Type:19, 'Revenue AZN':19,
  'Original Line ID':23, 'Revenue Reversal AZN':25, Product:42, Unit:16,
  Setting:42, Value:42, Name:52,
};
const col = (n) => {let s=''; for(n++;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s;};
function styleSheet(sheet, end, rows=26) {
  sheet.showGridLines = false;
  sheet.getRange('A1:'+end+rows).format.font = {name:'Arial',size:11,color:'#222222'};
  sheet.getRange('A1:'+end+rows).format.verticalAlignment='center';
  sheet.getRange('A2').format.font={name:'Arial',size:15,bold:true,color:'#222222'};
  sheet.getRange('A2:'+end+'2').format.rowHeight=26;
  sheet.getRange('A3:'+end+'4').format.font={name:'Arial',size:10,color:'#555555'};
  sheet.getRange('A3:'+end+'4').format.rowHeight=22;
  sheet.getRange('A6:'+end+'6').format={
    fill:'#F2F2F2',font:{name:'Arial',size:11,bold:true,color:'#222222'},
    horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,rowHeight:48,
    borders:{bottom:{style:'thin',color:'#999999'}},
  };
}
for(const [name, columns] of Object.entries(spec.columns)) {
  const sheet = wb.worksheets.add(spec.sheets[name]);
  const records=spec.data[name] || [];
  const last=Math.max(26,records.length+6);
  const end=col(columns.length-1);
  styleSheet(sheet,end,last);
  sheet.getRange('A2').values=[[spec.sheets[name]]];
  sheet.getRange('A3').values=[[notes[name][0]]];
  sheet.getRange('A4').values=[[notes[name][1]]];
  sheet.getRange('A6:'+end+'6').values=[columns.map(h=>spec.fields[h])];
  if(records.length) {
    const values=records.map(r=>columns.map(h=>{
      let v=r[h] ?? null;
      if(h==='Setting') v=spec.settings[v] || v;
      if(h==='Type'||h==='Purchase Type') v=types[v] || v;
      if((h==='Date'||h==='Document Date'||(name==='Settings'&&h==='Value'&&r.Setting!=='Company'))&&v) v=new Date(v);
      return v;
    }));
    sheet.getRange('A7:'+end+(6+records.length)).values=values;
  }
  const table=sheet.tables.add('A6:'+end+last,true,'Input'+name);
  table.style='TableStyleLight1';
  table.showFilterButton=true;
  sheet.getRange('A7:'+end+last).format.fill='#FFFFFF';
  sheet.getRange('A7:'+end+last).format.rowHeight=23;
  sheet.getRange('A7:'+end+last).format.borders={preset:'all',style:'hair',color:'#DDDDDD'};
  sheet.freezePanes.freezeRows(6);
  for(let i=0;i<columns.length;i++) {
    const h=columns[i], c=col(i);
    sheet.getRange(c+'1:'+c+last).format.columnWidth=widths[h] || 20;
    const cells=sheet.getRange(c+'7:'+c+'10006');
    if(['SKU','Shipment ID','Line ID','Expense ID','Return ID','Original Line ID','Document'].includes(h)) cells.setNumberFormat('@');
    if(['Date','Document Date'].includes(h)) cells.setNumberFormat('yyyy-mm-dd');
    if(['Amount FCY','Revenue AZN','Revenue Reversal AZN'].includes(h)) cells.setNumberFormat('#,##0.00;[Red](#,##0.00)');
    if(h==='Quantity') cells.setNumberFormat('#,##0.000');
    if(h==='FX to AZN') cells.setNumberFormat('0.000000');
    if(['Line ID','Expense ID','Return ID'].includes(h)) sheet.getRange(c+'7:'+c+last).format.font={name:'Arial',size:10,color:'#777777'};
    let rule=null;
    if(h==='SKU'&&name!=='Products') rule={type:'list',formula1:'INDIRECT("InputProducts[Код товара (SKU) *]")'};
    if(['Supplier','Customer','Counterparty'].includes(h)) rule={type:'list',formula1:'INDIRECT("InputCounterparties[Наименование *]")'};
    if(h==='Shipment ID'&&name==='Expenses') rule={type:'list',formula1:'INDIRECT("InputPurchases[Номер декларации / e-qaimə *]")'};
    if(h==='Purchase Type') rule={type:'list',values:['Импорт','Местная']};
    if(h==='Type') rule={type:'list',values:name==='Sales'?['Продажа','Списание']:['Покупатель','Поставщик']};
    if(h==='Original Line ID') rule={type:'list',formula1:'IF($B7="Покупатель",INDIRECT("InputSales[Код строки]"),INDIRECT("InputPurchases[Код строки]"))'};
    if(h==='Quantity'||h==='FX to AZN') rule={type:'decimal',operator:'greaterThan',formula1:0};
    if(rule) cells.dataValidation={rule};
    if((name==='Products'&&h==='SKU')||(name==='Counterparties'&&h==='Name')||['Line ID','Expense ID','Return ID'].includes(h)) {
      cells.conditionalFormats.addCustom('AND('+c+'7<>"",COUNTIF($'+c+'$7:$'+c+'$10006,'+c+'7)>1)',{fill:'#FCE6E4',font:{color:'#992222'}});
    }
  }
  if(name==='Settings') records.forEach((r,i)=>{
    if(r.Setting!=='Company') sheet.getRange('B'+(i+7)).setNumberFormat('yyyy-mm-dd');
  });
}
const guide=wb.worksheets.add(spec.sheets.Guide);
styleSheet(guide,'B',25);
guide.getRange('A1:A25').format.columnWidth=28;
guide.getRange('B1:B25').format.columnWidth=110;
guide.getRange('A2').values=[['Как вводить данные']];
guide.getRange('A6:B6').values=[['Действие','Правило']];
const instructions=[
 ['1. Справочники','Заполните Товары и Контрагенты. Код товара и название контрагента в справочнике не должны повторяться.'],
 ['2. Покупки','Импорт: номер декларации. Местная покупка: номер e-qaimə. Повторяйте номер для каждого товара документа.'],
 ['3. Расходы','Каждый сбор или услуга — отдельная строка. Привязка по номеру декларации / e-qaimə. Отрицательная сумма — корректировка.'],
 ['4. Продажи','Выберите Продажа или Списание. Количество положительное. У продажи обязательны покупатель и выручка, в том числе 0.'],
 ['5. Расчёт','Сохраните и закройте обе книги. Запустите Recalculate.bat. Проверьте Checks и время формирования отчёта.'],
 ['Автоматические коды','При расчёте программа заполняет пустые коды только у введённых операций и сохраняет исходник. Перед изменением создаётся копия в backups.'],
 ['6. Возвраты','После первого расчёта исходной операции откройте ввод заново. Выберите тип возврата и код исходной строки из списка. Код виден в конце строки покупки или продажи.'],
 ['Копирование строк','Для новой операции очистите скопированный код строки. Для исправления существующей операции сохраните её код. Коды не переименовывайте.'],
 ['Сортировка','Сортируйте всю таблицу, а не отдельный столбец. Постоянные коды сохраняют связь возвратов с исходными строками.'],
 ['Дубликаты','Повторы SKU допустимы в движениях, номера декларации — в товарах и расходах. Повтор кода строки блокируется. Похожие операции дают WARNING в Checks.'],
 ['Суммы и курсы','Все суммы без НДС. Сумма в валюте относится ко всей строке. Курс — AZN за единицу валюты; для AZN всегда 1.'],
 ['Поздние расходы','Все известные расходы пересчитывают исходные поступления и следующие месяцы, даже если дата расхода позже периода отчёта.'],
 ['Отчёты по поставкам','Shipment Summary — итоги; Shipment Cost — товары с полной стоимостью; Shipment Expenses — состав расходов. Фильтруйте по номеру поставки.'],
 ['Расширение таблицы','Добавляйте строки клавишей Tab из последней ячейки таблицы. Выпадающие списки используют умные таблицы справочников. Подсказки ввода настроены до строки 10006.'],
 ['Без формул','Вводите готовые значения. Названия и единицы товара находятся в справочнике и итоговом отчёте. Изменение ввода требует нового расчёта.'],
 ['Обязательные поля','Звёздочка * обозначает обязательное поле. Для возврата покупателя обязательна сумма уменьшения выручки, включая 0.'],
];
guide.getRange('A7:B'+(6+instructions.length)).values=instructions;
guide.getRange('A7:B'+(6+instructions.length)).format.wrapText=true;
guide.getRange('A7:B'+(6+instructions.length)).format.rowHeight=46;
guide.getRange('A7:A'+(6+instructions.length)).format.font={name:'Arial',size:11,bold:true};
wb.recalculate();
await fs.mkdir(previewDir,{recursive:true});
for(const name of Object.values(spec.sheets)){
  const sheet=wb.worksheets.getItem(name);
  const range=name===spec.sheets.Guide?'A1:B22':name===spec.sheets.Purchases?'A1:K12':name===spec.sheets.Expenses?'A1:I12':name===spec.sheets.Sales?'A1:H12':name===spec.sheets.Returns?'A1:F12':name===spec.sheets.Products?'A1:C12':name===spec.sheets.Counterparties?'A1:D12':'A1:D12';
  const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
  await fs.writeFile(previewDir+'/'+name+'.png',new Uint8Array(await blob.arrayBuffer()));
}
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:20},maxChars:1500})).ndjson);
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
console.log('Saved '+output);
