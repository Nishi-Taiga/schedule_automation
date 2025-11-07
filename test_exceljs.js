const ExcelJS = require('exceljs');
const fs = require('fs');

// セル値を安全に取得するヘルパー関数
function getCellValue(cell){
  if (cell.value === undefined || cell.value === null) {
    return null;
  }
  // リッチテキストの場合
  if (typeof cell.value === 'object' && cell.value.richText) {
    return cell.value.richText.map(part => part.text).join('');
  }
  // その他（文字列、数値、日付など）
  return cell.value;
}

async function testBoothExcel() {
  console.log('🔵 ブース表テスト開始');
  try {
    const workbook = new ExcelJS.Workbook();
    const buffer = fs.readFileSync('./ブース表テンプレート.xlsx');
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.worksheets[0];
    console.log('✓ ワークシート名:', worksheet.name);
    console.log('✓ 行数:', worksheet.rowCount);

    // 丸付き数字の検出テスト
    const circledRegex = /[\u2460-\u2473]/; // ①..⑳
    let circledCount = 0;
    let rowCount = 0;

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber <= 5 || (rowNumber >= 6 && rowNumber <= 10)) {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          let cellValue;
          try {
            // 日付・時刻セルの場合は cell.value (Date object) を優先
            if (cell.value instanceof Date) {
              cellValue = cell.value;
            } else if (cell.text !== undefined && cell.text !== null) {
              cellValue = cell.text;
            } else {
              cellValue = getCellValue(cell);
            }
          } catch (e) {
            cellValue = getCellValue(cell);
          }
          rowData.push(cellValue);
        });
        console.log(`  行${rowNumber}:`, rowData.slice(0, 8));
      }

      // 全行で丸付き数字を検索
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let cellValue;
        try {
          if (cell.value instanceof Date) {
            cellValue = cell.value;
          } else if (cell.text !== undefined && cell.text !== null) {
            cellValue = cell.text;
          } else {
            cellValue = getCellValue(cell);
          }
        } catch (e) {
          cellValue = getCellValue(cell);
        }

        if (cellValue && String(cellValue).match(circledRegex)) {
          circledCount++;
          if (circledCount <= 10) {
            console.log(`  🔵 丸付き数字検出: 行${rowNumber}, 列${colNumber}, 値="${cellValue}"`);
          }
        }
      });

      rowCount++;
    });
    console.log('✓ 総行数:', rowCount);
    console.log('✓ 丸付き数字の数:', circledCount);
    console.log('🔵 ブース表テスト完了\n');
    return true;
  } catch (e) {
    console.error('❌ エラー:', e.message);
    console.error(e.stack);
    return false;
  }
}

async function testTeacherExcel() {
  console.log('🟢 元シートテスト開始');
  try {
    const workbook = new ExcelJS.Workbook();
    const buffer = fs.readFileSync('./元シートテンプレート.xlsx');
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.worksheets[0];
    console.log('✓ ワークシート名:', worksheet.name);
    console.log('✓ 行数:', worksheet.rowCount);

    // 最初の10行を表示
    let rowCount = 0;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber <= 5) {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          let cellValue;
          try {
            // 日付・時刻セルの場合は cell.value (Date object) を優先
            if (cell.value instanceof Date) {
              cellValue = cell.value;
            } else if (cell.text !== undefined && cell.text !== null) {
              cellValue = cell.text;
            } else {
              cellValue = getCellValue(cell);
            }
          } catch (e) {
            cellValue = getCellValue(cell);
          }
          rowData.push(cellValue);
        });
        console.log(`  行${rowNumber}:`, rowData.slice(0, 8));
      }
      rowCount++;
    });
    console.log('✓ 総行数:', rowCount);
    console.log('🟢 元シートテスト完了\n');
    return true;
  } catch (e) {
    console.error('❌ エラー:', e.message);
    console.error(e.stack);
    return false;
  }
}

async function main() {
  console.log('=== ExcelJS 読み込みテスト ===\n');
  const result1 = await testBoothExcel();
  const result2 = await testTeacherExcel();

  if (result1 && result2) {
    console.log('✅ すべてのテスト成功！');
    process.exit(0);
  } else {
    console.log('❌ テスト失敗');
    process.exit(1);
  }
}

main();
