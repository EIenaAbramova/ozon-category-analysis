import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const d=JSON.parse(await fs.readFile(path.join(root,'outputs','workbook_data.json'),'utf8'));
const wb=Workbook.create();
const summary=wb.worksheets.add('Итоги');
const cats=wb.worksheets.add('Категории');
const products=wb.worksheets.add('Товары');
const source=wb.worksheets.add('Источник');
const blue='#005BFF'; const dark='#20334D';
function base(s,range){s.showGridLines=false;s.getRange(range).format.font={name:'Arial',size:10,color:dark};s.getRange(range).format.rowHeight=23;s.getRange(range).format.verticalAlignment='center';}
function title(s,text,last){s.getRange('A2').values=[[text]];s.getRange('A2').format.font={name:'Arial',size:16,bold:true,color:dark};s.getRange(`A3:${last}3`).format.borders={bottom:{style:'thin',color:'#CFD9E5'}};}
function header(s,range){s.getRange(range).format={fill:dark,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',rowHeight:46};}
base(summary,'A1:H25');summary.tabColor=blue; title(summary,'Ozon  Выбор категорий','H');
summary.getRange('A4').values=[['Период: 03.07.2024–30.07.2024. Сумма заказов используется как показатель выручки.']];
summary.getRange('A5:F5').values=[['ДРР ≤',.03,'Доля >1 млн ≥',.10,'Товаров ≥',30]];
for(const addr of ['B5','D5']){summary.getRange(addr).format.numberFormat='0.0%';summary.getRange(addr).format.fill='#FFF2CC';}
summary.getRange('F5').format.fill='#FFF2CC';
summary.getRange('A7:H7').values=[['Категория 4 уровня','Товаров','>1 млн, шт.','Доля >1 млн','ДРР взвеш.','Доля лидера','Отбор','Строка категорий']];header(summary,'A7:H7');
const catHeaders=['Категория 1 уровня','Категория 2 уровня','Категория 3 уровня','Категория 4 уровня','Товаров, шт.','>1 млн, шт.','Заказы, ₽','Расходы расчетные, ₽','ДРР средняя','ДРР ноль, шт.','Продавцов','Медианная цена, ₽','Доля лидера','ДРР взвеш.','Доля >1 млн','Отбор'];
base(cats,`A1:P${d.categories.length+5}`); title(cats,'Категории и расчёты','P');
cats.getRange('A4').values=[['Агрегаты из products CSV. Ключ категории — полный путь из четырёх уровней.']];
cats.getRange('A5:P5').values=[catHeaders];header(cats,'A5:P5');
cats.getRange(`A6:M${d.categories.length+5}`).values=d.categories;
const last=d.categories.length+5;
cats.getRange('N6:P6').formulas=[['=H6/G6','=F6/E6',`=IF(AND(E6>='Итоги'!$F$5,N6<='Итоги'!$B$5,O6>='Итоги'!$D$5),"Да","Нет")`]];
cats.getRange(`N6:P${last}`).fillDown();
cats.getRange(`G6:H${last}`).format.numberFormat='#,##0';cats.getRange(`L6:L${last}`).format.numberFormat='#,##0';
for(const col of ['I','M','N','O']) cats.getRange(`${col}6:${col}${last}`).format.numberFormat='0.00%';
cats.getRange(`P6:P${last}`).conditionalFormats.add('containsText',{text:'Да',format:{fill:'#E5F1FF',font:{bold:true,color:blue}}});
cats.tables.add(`A5:P${last}`,true,'CategoryAnalysis'); cats.freezePanes.freezeRows(5); cats.freezePanes.freezeColumns(4);
cats.getRange(`A6:D${last}`).format.wrapText=true;cats.getRange(`A6:P${last}`).format.rowHeight=60;
cats.getRange('A:A').format.columnWidth=25; cats.getRange('B:C').format.columnWidth=30; cats.getRange('D:D').format.columnWidth=44; cats.getRange('E:P').format.columnWidth=18;
for(let i=0;i<d.selectedCount;i++){
 const r=i+8,c=i+6;
 summary.getRange(`A${r}:H${r}`).formulas=[[`='Категории'!D${c}`,`='Категории'!E${c}`,`='Категории'!F${c}`,`='Категории'!O${c}`,`='Категории'!N${c}`,`='Категории'!M${c}`,`='Категории'!P${c}`,`=${c}`]];
}
summary.getRange(`D8:F${d.selectedCount+7}`).format.numberFormat='0.00%';
summary.getRange(`A8:A${d.selectedCount+7}`).format.wrapText=true;summary.getRange(`A8:H${d.selectedCount+7}`).format.rowHeight=42;
summary.getRange('A20').values=[['Пороги в жёлтых ячейках можно менять. Полный список фильтруйте на листе Категории.']];
summary.getRange('A21').values=[['Состав короткого списка фиксирован для базового отбора; для нового списка перезапустите анализ.']];
summary.getRange('A22').values=[['Доля лидера — часть заказов категории у крупнейшего продавца в этой выборке.']];
summary.getRange('A23').values=[['Источник содержит ТОП товаров. Результаты описывают выгрузку, а не весь рынок Ozon.']];
summary.getRange('A:A').format.columnWidth=54;summary.getRange('B:H').format.columnWidth=17;
base(products,`A1:J${d.products.length+5}`);title(products,'Примеры товаров для проверки поставки','J');
products.getRange('A4').values=[['Исторические примеры: заказы >1 млн ₽, ДРР ≤3%. Строка указывает на исходный Excel.']];
products.getRange('A5:J5').values=[['Категория','Название товара','Ссылка','Продавец','Заказы, ₽','ДРР','Цена, ₽','Доступность','Строка источника','Схема']];header(products,'A5:J5');
products.getRange(`A6:J${d.products.length+5}`).values=d.products;
products.getRange(`A6:B${d.products.length+5}`).format.wrapText=true;products.getRange(`A6:J${d.products.length+5}`).format.rowHeight=72;
products.getRange(`E6:E${d.products.length+5}`).format.numberFormat='#,##0';products.getRange(`G6:G${d.products.length+5}`).format.numberFormat='#,##0';
for(const c of ['F','H'])products.getRange(`${c}6:${c}${d.products.length+5}`).format.numberFormat='0.00%';
products.getRange('A:A').format.columnWidth=40;products.getRange('B:B').format.columnWidth=72;products.getRange('C:C').format.columnWidth=48;products.getRange('D:D').format.columnWidth=22;products.getRange('E:J').format.columnWidth=19;
products.tables.add(`A5:J${d.products.length+5}`,true,'ProductExamples');products.freezePanes.freezeRows(5);
base(source,'A1:C25');title(source,'Источник и контрольные показатели','C');
source.getRange(`A5:B${d.source.length+4}`).values=d.source;
header(source,'A5:B5'); source.getRange('A:A').format.columnWidth=49;source.getRange('B:B').format.columnWidth=105;
source.getRange('A19:B22').values=[['Контроль по категориям','Значение'],['Товары',null],['Товары >1 млн',null],['Сумма заказов, ₽',null]];header(source,'A19:B19');
source.getRange('B20:B22').formulas=[[`=SUM('Категории'!E6:E${last})`],[`=SUM('Категории'!F6:F${last})`],[`=SUM('Категории'!G6:G${last})`]];source.getRange('B20:B22').format.numberFormat='#,##0.00';
// Verify a real change of thresholds and a source aggregate; restore all test edits.
summary.getRange('B5').values=[[.01]];
if(cats.getRange('P6').values[0][0]!=='Нет' || cats.getRange('P7').values[0][0]!=='Да')throw Error('Threshold change did not propagate');
summary.getRange('B5').values=[[.03]];
const originalCost=cats.getRange('H6').values[0][0];
const originalRevenue=cats.getRange('G6').values[0][0];
cats.getRange('H6').values=[[originalRevenue*.04]];
if(Math.abs(cats.getRange('N6').values[0][0]-.04)>1e-10 || cats.getRange('P6').values[0][0]!=='Нет')throw Error('Source change did not propagate');
cats.getRange('H6').values=[[originalCost]];
wb.recalculate();
console.log('Threshold and aggregate perturbation tests passed; original inputs restored');
console.log((await wb.inspect({kind:'table',range:'Итоги!A7:H10',include:'values,formulas',tableMaxRows:4,tableMaxCols:8,maxChars:3000})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'Formula error scan',maxChars:2000})).ndjson);
await fs.mkdir(path.join(root,'qa'),{recursive:true});
for(const [s,r] of [['Итоги','A1:H18'],['Категории','D4:P12'],['Товары','A4:J9'],['Источник','A1:B22']]){
 const preview=await wb.render({sheetName:s,range:r,scale:1.5,format:'png'});
 await fs.writeFile(path.join(root,'qa',s+'.png'),new Uint8Array(await preview.arrayBuffer()));
}
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path.join(root,'outputs','Ozon_анализ.xlsx'));
console.log('Workbook exported');
