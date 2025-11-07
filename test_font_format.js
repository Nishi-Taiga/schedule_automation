const ExcelJS = require('exceljs');
const fs = require('fs');

async function testFontFormatting() {
  console.log('=== フォント設定テスト開始 ===\n');

  try {
    // ステップ1: ブース表テンプレートを読み込み
    console.log('📖 ステップ1: ブース表テンプレートを読み込み');
    const workbook = new ExcelJS.Workbook();
    const buffer = fs.readFileSync('./ブース表テンプレート.xlsx');
    await workbook.xlsx.load(buffer);
    console.log('✓ 読み込み完了\n');

    const worksheet = workbook.worksheets[0];
    console.log('ワークシート名:', worksheet.name);

    // ステップ2: テストデータをセルに書き込み（2段階方式）
    console.log('\n📝 ステップ2: セル値を書き込み（1段階目）');

    const testCells = [
      { row: 7, col: 4, value: '西T' },      // 行7, 列D
      { row: 7, col: 9, value: '田中T' },    // 行7, 列I
      { row: 8, col: 4, value: '佐藤T' },    // 行8, 列D
    ];

    const formattedCells = [];

    for (const {row, col, value} of testCells) {
      const cell = worksheet.getCell(row, col);

      // 既存の値を確認
      const existingValue = cell.value || '';
      console.log(`  セル(${row}, ${col}): 既存値="${existingValue}"`);

      // 値を書き込み
      cell.value = value;
      console.log(`  → 新しい値="${value}" を書き込み`);

      // フォーマット設定用に記録
      formattedCells.push({ cell, row, col, value });
    }

    console.log('✓ 値の書き込み完了\n');

    // ステップ3: フォーマット設定を適用（2段階目）
    console.log('🎨 ステップ3: フォーマット設定を適用（2段階目）');

    for (const {cell, row, col, value} of formattedCells) {
      console.log(`\n  セル(${row}, ${col}): "${value}"`);

      // 設定前の状態を確認
      console.log('    設定前:');
      console.log('      font:', cell.font);
      console.log('      alignment:', cell.alignment);

      // フォント設定（縦書き用フォントは '@' を先頭に付ける）
      cell.font = {
        name: '@MS PGothic',
        size: 8,
        family: 1,
        charset: 128
      };

      // 縦書き設定（ExcelJSでは'vertical'文字列を使用）
      cell.alignment = {
        textRotation: 'vertical',  // ExcelJSでは文字列'vertical'を使用
        vertical: 'top',
        horizontal: 'center',
        wrapText: true
      };

      // 設定後の状態を確認
      console.log('    設定後:');
      console.log('      font:', cell.font);
      console.log('      alignment:', cell.alignment);
    }

    console.log('\n✓ フォーマット設定完了\n');

    // ステップ4: ファイルを保存
    console.log('💾 ステップ4: ファイルを保存');
    const outputPath = './test_output.xlsx';
    const outputBuffer = await workbook.xlsx.writeBuffer();
    fs.writeFileSync(outputPath, outputBuffer);
    console.log(`✓ ファイル保存完了: ${outputPath}\n`);

    // ステップ5: 保存したファイルを再読み込みして検証
    console.log('🔍 ステップ5: 保存したファイルを再読み込みして検証');
    const verifyWorkbook = new ExcelJS.Workbook();
    await verifyWorkbook.xlsx.readFile(outputPath);
    const verifyWorksheet = verifyWorkbook.worksheets[0];

    console.log('\n【検証結果】');
    for (const {row, col, value} of testCells) {
      const cell = verifyWorksheet.getCell(row, col);
      console.log(`\nセル(${row}, ${col}): "${cell.value}"`);
      console.log('  フォント:');
      console.log('    name:', cell.font?.name);
      console.log('    size:', cell.font?.size);
      console.log('    family:', cell.font?.family);
      console.log('    charset:', cell.font?.charset);
      console.log('  配置:');
      console.log('    textRotation:', cell.alignment?.textRotation);
      console.log('    vertical:', cell.alignment?.vertical);
      console.log('    horizontal:', cell.alignment?.horizontal);
      console.log('    wrapText:', cell.alignment?.wrapText);

      // 検証（@MS PGothicとtextRotation: 255を確認）
      const fontOK = cell.font?.name === '@MS PGothic' && cell.font?.size === 8;
      // textRotationは255または'vertical'の場合OK
      const textRotation = cell.alignment?.textRotation;
      const alignmentOK = textRotation === 255 || textRotation === 'vertical';

      if (fontOK && alignmentOK) {
        console.log('  ✅ フォント設定OK');
      } else {
        console.log('  ❌ フォント設定NG');
        if (!fontOK) {
          console.log('     - フォントが正しくありません');
          console.log(`       期待: @MS PGothic 8pt, 実際: ${cell.font?.name} ${cell.font?.size}pt`);
        }
        if (!alignmentOK) {
          console.log('     - 縦書きが正しくありません');
          console.log(`       期待: 255 or 'vertical', 実際: ${textRotation}`);
        }
      }
    }

    console.log('\n=== テスト完了 ===');
    console.log(`\n生成されたファイル: ${outputPath}`);
    console.log('Excelで開いてフォント設定を確認してください。');

  } catch (error) {
    console.error('❌ エラー:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

testFontFormatting();
