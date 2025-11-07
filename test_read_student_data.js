const ExcelJS = require('exceljs');

async function testReadStudentData() {
  console.log('=== 生徒コマ数表読み込みテスト ===\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./生徒コマ数表テンプレート.xlsx');

    const worksheet = workbook.worksheets[0];
    console.log(`シート名: ${worksheet.name}\n`);

    // ヘッダー行を読み込み
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber] = cell.value;
    });

    console.log('列ヘッダー:');
    headers.forEach((header, index) => {
      if (header) {
        console.log(`  列${index}: ${header}`);
      }
    });

    console.log('\n全データ行:');
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // ヘッダーをスキップ

      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowData[colNumber] = cell.value;
      });

      // 空行をスキップ
      if (!rowData[1] && !rowData[2] && !rowData[3]) return;

      console.log(`\n行${rowNumber}:`);
      console.log(`  学年: ${rowData[1]}`);
      console.log(`  学校名: ${rowData[2]}`);
      console.log(`  生徒名: ${rowData[3]}`);

      // 科目コマ数
      const subjects = {
        '英': rowData[4],
        '英検': rowData[5],
        '数': rowData[6],
        '算': rowData[7],
        '国': rowData[8],
        '理': rowData[9],
        '社': rowData[10],
        '古': rowData[11],
        '物': rowData[12],
        '化': rowData[13],
        '生': rowData[14],
        '地': rowData[16],
        '政': rowData[17],
        '世': rowData[18],
        '日': rowData[19]
      };

      console.log('  科目コマ数:');
      Object.entries(subjects).forEach(([subject, count]) => {
        if (count) {
          console.log(`    ${subject}: ${count}コマ`);
        }
      });

      console.log(`  希望講師: ${rowData[21] || '指定なし'}`);
      console.log(`  NG講師: ${rowData[22] || '指定なし'}`);
      console.log(`  NG生徒: ${rowData[23] || '指定なし'}`);
      console.log(`  希望時間: ${rowData[24] || '指定なし'}`);
      console.log(`  NG日程: ${rowData[25] || '指定なし'}`);
      console.log(`  備考: ${rowData[26] || ''}`);
    });

    console.log('\n=== 読み込み完了 ===');

  } catch (error) {
    console.error('❌ エラー:', error);
    process.exit(1);
  }
}

testReadStudentData();
