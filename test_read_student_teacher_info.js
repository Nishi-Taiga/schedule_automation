const ExcelJS = require('exceljs');

async function testReadStudentTeacherInfo() {
  console.log('=== 生徒・講師情報.xlsx 読み込みテスト ===\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./生徒・講師情報.xlsx');

    console.log(`ワークブック内のシート数: ${workbook.worksheets.length}\n`);

    // 全シートの概要を表示
    workbook.worksheets.forEach((worksheet, index) => {
      console.log(`\n=== シート${index + 1}: ${worksheet.name} ===`);
      console.log(`行数: ${worksheet.rowCount}`);
      console.log(`列数: ${worksheet.columnCount}\n`);

      // ヘッダー行を表示
      const headerRow = worksheet.getRow(1);
      const headers = [];
      headerRow.eachCell((cell, colNumber) => {
        headers.push(`列${colNumber}: ${cell.value}`);
      });
      console.log('ヘッダー行:');
      headers.forEach(h => console.log(`  ${h}`));

      // 最初の5行のデータを表示
      console.log('\nデータサンプル（最初の5行）:');
      let rowCount = 0;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // ヘッダーをスキップ
        if (rowCount >= 5) return;

        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData.push(cell.value);
        });

        console.log(`  行${rowNumber}:`, rowData.slice(0, 10)); // 最初の10列のみ表示
        rowCount++;
      });
    });

    console.log('\n=== 読み込み完了 ===');

  } catch (error) {
    console.error('❌ エラー:', error);
    process.exit(1);
  }
}

testReadStudentTeacherInfo();
