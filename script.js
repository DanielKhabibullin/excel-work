// script.js (полная версия со всеми удалениями)

// Константы
const EXCLUDED_ITEMS = ['Скоба Y4DT (7281.05.00.208)'];

const ITEMS_TO_REMOVE = [
    'Заготовка Каркас подлокотника 35/7',
    'Каркас подголовника пластинчатого',
    'Перемычка верхняя внутренняя',
    'Перемычка верхняя наружняя',
    'Пластина под лейбл',
    'Переходная деталь кронштейна подголовника Ф10-2D',
    'Полоса верхней перемычки',
    'Рукоятка малая',
    'Стенка левая 1D подголовника для пластинчатой перемычки',
    'Стенка правая 1D подголовника для пластинчатой перемычки',
    'Уголок кронштейна 1D подголовника для пластинчатой перемычки',
    'Уголок MSS 1,0 передний',
    'Уголок S1,0 MS1,0 передний',
    'Уголок S1,0 MSS1,0 MS1,0 задний с гайкой',
    'Фиксатор кронштейна 1D подголовника для пластинчатой перемычки',
    'Ухо',
    'Ухо стальное',
    'Гантель',
    'Короб S1,0 с гайками',
    'Короб MSS 1,0 с гайками',
    'Короб MS2,0 MSS 2,0 с гайками',
    'Подлокотник (каркас) 35/7 Левый б/п',
    'Подлокотник (каркас) 35/7 Левый б/п',
    'Подлокотник (каркас) 35/7 Правый б/п',
    'Подлокотник (каркас) 35/7 Правый б/п',
    'Пластина хребта (9392.00.00.005)',
    'Пластина хребта (9392.00.00.006)',
    'Скоба Y4DT (7281.05.00.208) б/п',
    'Боковина Держателя спинки Y2DM',
    'Крепление газовой пружины Держателя спинки Y2DM-L',
    'Крепление газовой пружины Держателя спинки Y2DM',
    'Проушина с гайкой Держателя спинки Y2DM',
    'Косынка Держателя спинки Y2DM',

];

const COMBINE_GROUPS = [
    { name: 'Хребет + Скоба+пружина РПС + Пластина прижимная Y4DF', items: ['Хребет Y4DT', 'Скоба РПС', 'Пружина РПС', 'Пластина прижимная Y4DF'], char: 'ч/м', sumQuantity: false },
    { name: 'Кронштейн + Пружина + Переходная 1д подголовника', items: ['Кронштейн 1D подголовника для пластинчатой перемычки', 'Пружина кронштейна подголовника 1D', 'Переходная деталь кронштейна подголовника ф10-2D'], char: 'ч/м', sumQuantity: false },
    { name: 'Барашек малый + Эмблема МЕТТА', items: ['Барашек малый', 'Эмблема "МЕТТА"'], char: 'хром', sumQuantity: false },
    { name: 'Крепление к мультиблоку + Ручка РУА', items: ['Крепление к мультиблоку', 'Ручка РУА'], char: 'ч/м', sumQuantity: false },
    { name: 'Втулка EQ + Пластина подпорная задняя EQ-B2', items: ['Втулка EQ', 'Пластина подпорная задняя EQ-B2'], char: 'ч/м', sumQuantity: false },
    { name: 'Крепл. газ. пруж. + Пластина ограничителя + .005 + .006 Y4DT', items: ['Крепление газовой пружины Y4DT', 'Пластина ограничителя', 'Пластина хребта 9392.00.00.005', 'Пластина хребта 9392.00.00.006'], char: 'ч/м', sumQuantity: false },
    { name: 'Уголки передние S1,0 MS2,0 MSS1,0 MSS2,0', items: ['Уголок MSS1,0 MSS2,0 передний', 'Уголок S1,0 MS2,0 передний'], char: 'ч/м', sumQuantity: true }
];

// DOM элементы
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const fileInfo = document.getElementById('fileInfo');
const loading = document.getElementById('loading');
const tableContainer = document.getElementById('tableContainer');
const tableHeader = document.getElementById('tableHeader');
const tableBody = document.getElementById('tableBody');
const stats = document.getElementById('stats');
const error = document.getElementById('error');

// Инициализация
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
});
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
});
['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
});

dropZone.addEventListener('drop', handleDrop, false);
fileInput.addEventListener('change', handleFileSelect, false);
// window.addEventListener('load', () => setTimeout(loadTestFile, 100));

// Вспомогательные функции
function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

// function shouldProcessRow(row) {
//     if (!row || !row.cells || row.cells.length === 0) return false;
//     const fullName = row.cells[0]?.textContent?.trim() || '';
//     return !EXCLUDED_ITEMS.some(excluded => fullName.includes(excluded));
// }

function createSeparatorRow(text) {
    const row = document.createElement('tr');
    row.className = 'separator-row';
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = text;
    row.appendChild(cell);
    return row;
}

// Обработчики файлов
function handleDrop(e) {
    const files = e.dataTransfer.files;
    if (files.length) handleFile(files[0]);
}

function handleFileSelect(e) {
    if (e.target.files.length) handleFile(e.target.files[0]);
}

function handleFile(file) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        showError('Пожалуйста, выберите файл .xlsx, .xls или .csv');
        return;
    }

    fileInfo.textContent = `Файл: ${file.name}`;
    loading.style.display = 'block';
    tableContainer.style.display = 'none';
    stats.style.display = 'none';
    error.style.display = 'none';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (!jsonData.length) throw new Error('Файл пуст');

            displayTable(jsonData);
        } catch (err) {
            showError('Ошибка: ' + err.message);
        } finally {
            loading.style.display = 'none';
        }
    };
    reader.onerror = () => showError('Ошибка чтения файла');
    reader.readAsArrayBuffer(file);
}

// function loadTestFile() {
//     const testFileName = 'перерасчет 0403.xlsx';
//     fetch(testFileName)
//         .then(response => {
//             if (!response.ok) throw new Error(`Файл ${testFileName} не найден`);
//             return response.arrayBuffer();
//         })
//         .then(data => {
//             fileInfo.textContent = `Автозагрузка: ${testFileName}`;
//             loading.style.display = 'block';
//             const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
//             const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//             const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
//             if (!jsonData.length) throw new Error('Файл пуст');
//             displayTable(jsonData);
//         })
//         .catch(error => {
//             showError('Ошибка автозагрузки: ' + error.message);
//             document.querySelector('h1').style.display = 'block';
//             document.getElementById('dropZone').style.display = 'block';
//         })
//         .finally(() => (loading.style.display = 'none'));
// }

// Основная функция отображения
function displayTable(data) {
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    const headers = data[0] || [];
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header || 'Без названия';
        headerRow.appendChild(th);
    });
    tableHeader.appendChild(headerRow);

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tr = document.createElement('tr');
        for (let j = 0; j < headers.length; j++) {
            const td = document.createElement('td');
            td.textContent = row[j] !== undefined ? row[j] : '';
            tr.appendChild(td);
        }
        tableBody.appendChild(tr);
    }

    tableContainer.style.display = 'block';
    updateStats(data);
    shortenHeaders();    // СОКРАЩЕНИЕ ЗАГОЛОВКОВ
    // Последовательная обработка
    sortTableByFirstColumn();
    normalizeCharacteristics();
    normalizeLeftRightNames();
    removeDrawingNumberItems();      // ← новая функция: удаление строк с чертежным номером
    removeBPDuplicates();            // ← новая функция: удаление б/п дублей
    removeBPInNameDuplicates();
    removeSpecificItems();
    combineItems();
    combineLeftRightItems();
    sortTableByFirstColumn();
    reorganizeSpinKarkas();
    removeLastColumn();
    hideEverythingExceptTable();
    document.getElementById('saveBtn').style.display = 'block';


}

// СОКРАЩЕНИЕ ЗАГОЛОВКОВ
function shortenHeaders() {
  const headerRow = document.querySelector('#tableHeader tr');
  if (!headerRow) return;

  Array.from(headerRow.cells).forEach(th => {
      const text = th.textContent.trim();
      if (text === 'Характеристика') {
          th.textContent = 'Хар-ка';
      }
      // Если хочешь сокращать и другие заголовки:
      // if (text === 'Номенклатура') th.textContent = 'Наим.';
      // if (text === 'Количество') th.textContent = 'Кол-во';
  });
}

// СОРТИРОВКА
function sortTableByFirstColumn() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    rows.sort((a, b) => {
        const aText = a.cells[0]?.textContent?.trim() || '';
        const bText = b.cells[0]?.textContent?.trim() || '';
        return aText.localeCompare(bText, 'ru');
    });
    tableBody.innerHTML = '';
    rows.forEach(row => tableBody.appendChild(row));
}

// УДАЛЕНИЕ ПОСЛЕДНЕГО СТОЛБЦА
function removeLastColumn() {
    const headerRow = document.querySelector('#tableHeader tr');
    if (headerRow?.cells?.length) {
        headerRow.deleteCell(-1);
    }

    document.querySelectorAll('#tableBody tr').forEach(row => {
        if (row.cells?.length) {
            row.deleteCell(-1);
        }
    });
}

// СКРЫТИЕ ВСЕГО КРОМЕ ТАБЛИЦЫ
function hideEverythingExceptTable() {
    document.querySelector('h1').style.display = 'none';
    document.getElementById('dropZone').style.display = 'none';
    document.getElementById('stats').style.display = 'none';
    document.getElementById('error').style.display = 'none';

    tableContainer.style.display = 'block';
    tableContainer.style.maxHeight = '100vh';
    tableContainer.style.height = 'calc(100vh - 40px)';

    document.body.style.padding = '0';
    document.querySelector('.container').style.padding = '0';
    document.querySelector('.container').style.maxWidth = '100%';
    document.querySelector('.container').style.borderRadius = '0';
}

// НОРМАЛИЗАЦИЯ ХАРАКТЕРИСТИК
function normalizeCharacteristics() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    rows.forEach(row => {
        if (row.cells.length < 2) return;

        const charCell = row.cells[1];
        const char = charCell.textContent?.trim() || '';
        let newChar = char;

        if (char === 'б/п под хром' || char === 'под хром') newChar = 'хром';
        if (char === 'б/п окраска' || char === 'окраска') newChar = 'ч/м';

        if (newChar !== char) {
            charCell.textContent = newChar;
        }
    });
}

// УДАЛЕНИЕ СТРОК С ЧЕРТЕЖНЫМ НОМЕРОМ (В СКОБКАХ)
function removeDrawingNumberItems() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const nameMap = new Map();

    // Собираем все названия без скобок
    rows.forEach(row => {
        if (row.cells.length < 1) return;
        const nameCell = row.cells[0];
        const fullName = nameCell.textContent?.trim() || '';
        if (EXCLUDED_ITEMS.some(ex => fullName.includes(ex))) return;

        const baseName = fullName.replace(/\s*\([^)]*\)\s*$/, '').trim();

        if (!nameMap.has(baseName)) {
            nameMap.set(baseName, []);
        }
        nameMap.get(baseName).push({
            row: row,
            fullName: fullName,
            hasDrawing: fullName.includes('(') && fullName.includes(')') && fullName !== baseName
        });
    });

    const toRemove = new Set();
    nameMap.forEach((items, baseName) => {
        const hasBase = items.some(item => !item.hasDrawing);
        const drawingItems = items.filter(item => item.hasDrawing);

        if (hasBase && drawingItems.length > 0) {
            drawingItems.forEach(item => toRemove.add(item.row));
        }
    });

    toRemove.forEach(row => row.remove());
}

// УДАЛЕНИЕ ДУБЛЕЙ С "б/п" (если есть такая же позиция с ч/м или хром и тем же количеством)
function removeBPDuplicates() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const itemMap = new Map(); // ключ: название

    // Собираем все строки
    rows.forEach(row => {
        if (row.cells.length < 3) return;

        const name = row.cells[0].textContent?.trim() || '';
        const char = row.cells[1].textContent?.trim() || '';
        const qty = row.cells[2].textContent?.trim() || '';

        const key = name;
        if (!itemMap.has(key)) {
            itemMap.set(key, []);
        }
        itemMap.get(key).push({ row, char, qty, hasBP: char.includes('б/п') });
    });

    const toRemove = new Set();

    itemMap.forEach((items, name) => {
        if (items.length < 2) return;

        // Группируем по количеству
        const qtyGroups = new Map();
        items.forEach(item => {
            if (!qtyGroups.has(item.qty)) {
                qtyGroups.set(item.qty, []);
            }
            qtyGroups.get(item.qty).push(item);
        });

        qtyGroups.forEach((group, qty) => {
            if (group.length < 2) return;

            const bpItems = group.filter(item => item.hasBP);
            const nonBPItems = group.filter(item => !item.hasBP);

            if (bpItems.length > 0 && nonBPItems.length > 0) {
                bpItems.forEach(item => toRemove.add(item.row));
            }
        });
    });

    toRemove.forEach(row => row.remove());
}

// УДАЛЕНИЕ КОНКРЕТНЫХ ПОЗИЦИЙ ИЗ СПИСКА
function removeSpecificItems() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const spinMap = new Map();
    const skbMap = new Map();

    // Сбор информации
    rows.forEach(row => {
        if (row.cells.length < 3) return;
        const name = row.cells[0].textContent?.trim() || '';
        const char = row.cells[1].textContent?.trim() || '';
        const qty = row.cells[2].textContent?.trim() || '';

        if (name.includes('Каркас спинки') && name.includes('с прутком')) {
            spinMap.set(`${name.replace('с прутком', '').trim()}|${qty}`, true);
        }
        if (name.includes('Каркас спинки SKB')) {
            const key = `${name}|${char}|${qty}`;
            if (!skbMap.has(key)) skbMap.set(key, []);
            skbMap.get(key).push(row);
        }
    });

    // Удаление
    const toRemove = new Set();
    rows.forEach(row => {
        if (row.cells.length < 3) return;
        const name = row.cells[0].textContent?.trim() || '';
        const char = row.cells[1].textContent?.trim() || '';
        const qty = row.cells[2].textContent?.trim() || '';

        // Из списка
        if (ITEMS_TO_REMOVE.some(item => name.includes(item))) {
            toRemove.add(row);
            return;
        }

        // Каркасы без прутка
        if (name.includes('Каркас спинки') && !name.includes('с прутком') && spinMap.has(`${name}|${qty}`)) {
            toRemove.add(row);
            return;
        }

        // SKB дубли
        if (name.includes('Каркас спинки SKB')) {
            const key = `${name}|${char}|${qty}`;
            const items = skbMap.get(key) || [];
            if (items.length > 1 && items[0] !== row) {
                toRemove.add(row);
                return;
            }
        }

        // Y4DF с б/п
        if (name.includes('Каркас спинки Y4DF') && char.includes('б/п')) {
            toRemove.add(row);
        }
    });

    toRemove.forEach(row => row.remove());
}

// ОБЪЕДИНЕНИЕ ГРУПП
function combineItems() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const toRemove = new Set();
    const newRows = [];

    COMBINE_GROUPS.forEach(group => {
        const found = [];
        let totalQty = 0;
        let allFound = true;

        group.items.forEach(itemName => {
            const match = rows.find(row => 
                !toRemove.has(row) &&
                row.cells[0]?.textContent?.includes(itemName) &&
                row.cells[1]?.textContent === group.char
            );
            if (match) {
                const qty = parseInt(match.cells[2]?.textContent) || 0;
                found.push({ row: match, qty, name: match.cells[0]?.textContent });
                toRemove.add(match);
                if (group.sumQuantity) {
                    totalQty += qty;
                } else {
                    totalQty = qty;
                }
            } else {
                allFound = false;
            }
        });

        if (allFound && found.length === group.items.length) {
            const tr = document.createElement('tr');
            tr.className = 'combined-row';

            const td1 = document.createElement('td');
            td1.textContent = group.name;
            tr.appendChild(td1);

            const td2 = document.createElement('td');
            td2.textContent = group.char;
            tr.appendChild(td2);

            const td3 = document.createElement('td');
            td3.textContent = totalQty;
            tr.appendChild(td3);

            const td4 = document.createElement('td');
            td4.textContent = 'СП';
            tr.appendChild(td4);

            tr.title = found.map(f => `${f.name} (${f.qty})`).join('\n');
            newRows.push(tr);
        } else {
            // Если не все найдены, возвращаем строки обратно
            found.forEach(f => toRemove.delete(f.row));
        }
    });

    toRemove.forEach(row => row.remove());
    newRows.forEach(row => tableBody.appendChild(row));
}

// НОРМАЛИЗАЦИЯ НАЗВАНИЙ (Левый/Правый)
function normalizeLeftRightNames() {
  const rows = Array.from(tableBody.querySelectorAll('tr'));
  let changes = 0;

  rows.forEach(row => {
      if (row.cells.length < 1) return;
      
      const nameCell = row.cells[0];
      let name = nameCell.textContent?.trim() || '';
      let newName = name;

      // Приводим все варианты к единому виду "Левый"/"Правый"
      newName = newName.replace(/\bлевая\b/gi, 'Левый');
      newName = newName.replace(/\bлевое\b/gi, 'Левый');
      newName = newName.replace(/\bправая\b/gi, 'Правый');
      newName = newName.replace(/\bправое\b/gi, 'Правый');

      if (newName !== name) {
          nameCell.textContent = newName;
          changes++;
      }
  });

  console.log(`✅ Нормализовано названий: ${changes}`);
  return changes;
}

// ОБЪЕДИНЕНИЕ ЛЕВЫХ И ПРАВЫХ 
function combineLeftRightItems() {
  const rows = Array.from(tableBody.querySelectorAll('tr'));
  const groups = new Map();
  const toRemove = new Set();
  const newRows = [];

  console.log('🔍 ПОИСК ПАР ЛЕВЫЙ/ПРАВЫЙ:');

  rows.forEach(row => {
      if (row.cells.length < 3) return;

      const name = row.cells[0].textContent?.trim() || '';
      const char = row.cells[1].textContent?.trim() || '';
      const qty = parseInt(row.cells[2]?.textContent) || 0;

      // Правильная проверка на все варианты написания
      const nameLower = name.toLowerCase();
      const isLeft = nameLower.includes('левый') || nameLower.includes('левая') || nameLower.includes('лев');
      const isRight = nameLower.includes('правый') || nameLower.includes('правая') || nameLower.includes('прав');
      
      if (!isLeft && !isRight) return;

      // Получаем базовое название, убирая "Левый"/"Правый" и их варианты
      let baseName = name
          .replace(/\s*Левый\s*/gi, '')
          .replace(/\s*левый\s*/gi, '')
          .replace(/\s*Левая\s*/gi, '')
          .replace(/\s*левая\s*/gi, '')
          .replace(/\s*Лев\s*/gi, '')
          .replace(/\s*лев\s*/gi, '')
          .replace(/\s*Правый\s*/gi, '')
          .replace(/\s*правый\s*/gi, '')
          .replace(/\s*Правая\s*/gi, '')
          .replace(/\s*правая\s*/gi, '')
          .replace(/\s*Прав\s*/gi, '')
          .replace(/\s*прав\s*/gi, '')
          .replace(/\s*б\/п\s*$/i, '')
          .replace(/\s+/g, ' ')
          .trim();

      const key = `${baseName}|${char}`;

      if (!groups.has(key)) {
          groups.set(key, { 
              baseName, 
              char, 
              left: [], 
              right: [],
              originalName: name
          });
      }
      
      const group = groups.get(key);
      
      if (isLeft) {
          group.left.push({ row, qty, name });
          console.log(`  Левый: "${name}" (${qty}) → группа: "${baseName}"`);
      }
      if (isRight) {
          group.right.push({ row, qty, name });
          console.log(`  Правый: "${name}" (${qty}) → группа: "${baseName}"`);
      }
  });

  console.log('\n📋 НАЙДЕННЫЕ ГРУППЫ:');
  let pairCount = 0;
  
  groups.forEach((group, key) => {
      if (group.left.length > 0 && group.right.length > 0) {
          const pairs = Math.min(group.left.length, group.right.length);
          console.log(`\n  ✅ "${group.baseName}" (${group.char}) - ${pairs} пар:`);
          
          for (let i = 0; i < pairs; i++) {
              const left = group.left[i];
              const right = group.right[i];
              const total = left.qty + right.qty;
              
              console.log(`     Пара ${i+1}: ${left.qty} + ${right.qty} = ${total}`);

              const tr = document.createElement('tr');
              tr.className = 'combined-leftright';

              const td1 = document.createElement('td');
              td1.textContent = `${group.baseName} (Левый + Правый)`;
              tr.appendChild(td1);

              const td2 = document.createElement('td');
              td2.textContent = group.char;
              tr.appendChild(td2);

              const td3 = document.createElement('td');
              td3.textContent = total;
              tr.appendChild(td3);

              const td4 = document.createElement('td');
              td4.textContent = 'СП';
              tr.appendChild(td4);

              tr.title = `Левый: ${left.name} (${left.qty})\nПравый: ${right.name} (${right.qty})`;

              toRemove.add(left.row);
              toRemove.add(right.row);
              newRows.push(tr);
              pairCount++;
          }
          
          if (group.left.length > group.right.length) {
              console.log(`     ⚠️ Осталось левых без пары: ${group.left.length - group.right.length}`);
          } else if (group.right.length > group.left.length) {
              console.log(`     ⚠️ Осталось правых без пары: ${group.right.length - group.left.length}`);
          }
      }
  });

  // Удаляем старые строки
  toRemove.forEach(row => row.remove());
  
  // Добавляем новые строки
  newRows.forEach(row => tableBody.appendChild(row));

  console.log(`\n📊 ИТОГО:`);
  console.log(`   Удалено старых строк: ${toRemove.size}`);
  console.log(`   Добавлено новых строк: ${newRows.length}`);
  console.log(`   Объединено пар: ${pairCount}`);
  
  return { removed: toRemove.size, added: newRows.length, pairs: pairCount };
}

// УДАЛЕНИЕ ДУБЛЕЙ С "б/п" В КОНЦЕ НАЗВАНИЯ
function removeBPInNameDuplicates() {
  const rows = Array.from(tableBody.querySelectorAll('tr'));
  const nameMap = new Map();

  // Собираем все названия
  rows.forEach(row => {
      if (row.cells.length < 3) return;
      const name = row.cells[0].textContent?.trim() || '';
      const char = row.cells[1].textContent?.trim() || '';
      const qty = row.cells[2].textContent?.trim() || '';

      const hasBPAtEnd = /\s*б\/п\s*$/i.test(name);
      const baseName = name.replace(/\s*б\/п\s*$/i, '').trim();

      const key = `${baseName}|${char}|${qty}`;
      if (!nameMap.has(key)) {
          nameMap.set(key, { baseName, char, qty, rows: [] });
      }
      nameMap.get(key).rows.push({ row, name, hasBPAtEnd });
  });

  const toRemove = new Set();

  nameMap.forEach((group, key) => {
      if (group.rows.length < 2) return;

      const bpRows = group.rows.filter(r => r.hasBPAtEnd);
      const nonBPRows = group.rows.filter(r => !r.hasBPAtEnd);

      if (bpRows.length > 0 && nonBPRows.length > 0) {
          bpRows.forEach(r => toRemove.add(r.row));
      }
  });

  toRemove.forEach(row => row.remove());
}

// ПЕРЕОРГАНИЗАЦИЯ КАРКАСОВ СПИНОК
function reorganizeSpinKarkas() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const spinCHM = [];
    const spinHrom = [];
    const spinSub1 = [];
    const other = [];

    rows.forEach(row => {
        if (row.cells.length < 2) {
            other.push(row);
            return;
        }

        const name = row.cells[0].textContent?.trim() || '';
        const char = row.cells[1].textContent?.trim() || '';

        if (name.includes('Спинка SUB1')) {
            spinSub1.push(row);
        } else if (name.includes('Каркас спинки')) {
            if (char.includes('хром')) {
                spinHrom.push(row);
            } else {
                spinCHM.push(row);
            }
        } else {
            other.push(row);
        }
    });

    const sortAZ = (a, b) => {
        const nameA = a.cells[0]?.textContent?.trim() || '';
        const nameB = b.cells[0]?.textContent?.trim() || '';
        return nameA.localeCompare(nameB, 'ru');
    };
    spinCHM.sort(sortAZ);
    spinHrom.sort(sortAZ);

    tableBody.innerHTML = '';
    other.forEach(row => tableBody.appendChild(row));

    if (spinSub1.length || spinCHM.length || spinHrom.length) {
        tableBody.appendChild(createSeparatorRow('━━━ СПИНКИ ━━━'));
    }
    spinSub1.forEach(row => tableBody.appendChild(row));

    if (spinCHM.length) {
        tableBody.appendChild(createSeparatorRow('━━━ КАРКАСЫ СПИНОК (ч/м) ━━━'));
        spinCHM.forEach(row => tableBody.appendChild(row));
    }
    if (spinHrom.length) {
        tableBody.appendChild(createSeparatorRow('━━━ КАРКАСЫ СПИНОК (хром) ━━━'));
        spinHrom.forEach(row => tableBody.appendChild(row));
    }
}

// СОХРАНЕНИЕ В EXCEL
function saveToExcel() {
  try {
      // Получаем данные из таблицы
      const headers = [];
      const headerRow = document.querySelector('#tableHeader tr');
      if (headerRow) {
          Array.from(headerRow.cells).forEach(th => {
              headers.push(th.textContent);
          });
      }

      const data = [headers];
      const rows = document.querySelectorAll('#tableBody tr');
      
      rows.forEach(row => {
          const rowData = [];
          Array.from(row.cells).forEach(cell => {
              rowData.push(cell.textContent);
          });
          data.push(rowData);
      });

      // Создаем книгу Excel
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, 'Лист1');

      // Сохраняем файл
      const fileName = `обработка_${new Date().toISOString().slice(0,10)}.xlsx`;
      XLSX.writeFile(wb, fileName);

      // Показываем сообщение (опционально)
      const saveMsg = document.createElement('div');
      saveMsg.textContent = '✅ Файл сохранен!';
      saveMsg.style.cssText = `
          position: fixed;
          bottom: 20px;
          right: 20px;
          background: #28a745;
          color: white;
          padding: 10px 20px;
          border-radius: 5px;
          z-index: 1000;
      `;
      document.body.appendChild(saveMsg);
      setTimeout(() => saveMsg.remove(), 3000);

  } catch (error) {
      alert('Ошибка при сохранении: ' + error.message);
  }
}

function updateStats(data) {
    stats.innerHTML = `<strong>Статистика:</strong> Строк: ${data.length - 1} | Колонок: ${data[0]?.length || 0} | Лист: Лист 1`;
    stats.style.display = 'block';
}

function showError(message) {
    error.textContent = message;
    error.style.display = 'block';
    tableContainer.style.display = 'none';
    stats.style.display = 'none';
    fileInfo.textContent = 'или перетащите файл сюда';
}