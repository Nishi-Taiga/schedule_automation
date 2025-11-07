const ExcelJS = require('exceljs');

async function createStudentDemandTemplate() {
  console.log('=== 生徒コマ数表テンプレート作成開始 ===\n');

  const workbook = new ExcelJS.Workbook();

  // ========================================
  // シート1: 生徒コマ数表（1生徒1行形式）
  // ========================================
  const mainSheet = workbook.addWorksheet('生徒コマ数表');

  // ヘッダー行
  const headers = [
    '学年',      // 1
    '学校名',    // 2
    '生徒名',    // 3
    '英',        // 4: 英語
    '英検',      // 5
    '数',        // 6: 数学
    '算',        // 7: 算数
    '国',        // 8: 国語
    '理',        // 9: 理科
    '社',        // 10: 社会
    '古',        // 11: 古文
    '物',        // 12: 物理
    '化',        // 13: 化学
    '生',        // 14: 生物
    '',          // 15: 空欄
    '地',        // 16: 地理
    '政',        // 17: 政治経済
    '世',        // 18: 世界史
    '日',        // 19: 日本史
    '',          // 20: 空欄
    '希望講師',  // 21
    'NG講師',    // 22
    'NG生徒',    // 23
    '希望時間',  // 24
    'NG日程',    // 25
    '備考'       // 26
  ];

  mainSheet.addRow(headers);

  // ヘッダー行のスタイル設定
  const headerRow = mainSheet.getRow(1);
  headerRow.font = { bold: true, size: 10, name: 'Meiryo UI' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E1F2' }  // 薄い青
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 18;

  // 列幅設定
  mainSheet.getColumn(1).width = 6;   // 学年
  mainSheet.getColumn(2).width = 12;  // 学校名
  mainSheet.getColumn(3).width = 12;  // 生徒名
  // 科目列（4-20）
  for (let col = 4; col <= 20; col++) {
    mainSheet.getColumn(col).width = 4;
  }
  mainSheet.getColumn(21).width = 12; // 希望講師
  mainSheet.getColumn(22).width = 12; // NG講師
  mainSheet.getColumn(23).width = 15; // NG生徒
  mainSheet.getColumn(24).width = 25; // 希望時間
  mainSheet.getColumn(25).width = 12; // NG日程
  mainSheet.getColumn(26).width = 15; // 備考

  // サンプルデータ（実際のフォーマットに準拠）
  const sampleData = [
    ['S4', '南蒲小', '松橋', '', '', '', 4, 4, '', '', '', '', '', '', '', '', '', '', '', '', '西T', '', '', '水17,金17,月17,木17', '火,土', ''],
    ['S4', '糀谷小', '深澤', '', '', '', 8, 4, '', '', '', '', '', '', '', '', '', '', '', '', '', '田中T', '松橋', '月17,火17,木17,月16,火16,木16', '', ''],
    ['M1', '東中', '佐藤', 2, '', 3, '', 2, '', '', '', '', '', '', '', '', '', '', '', '', '鈴木T', '山田T', '深澤', '月18,水18,金18', '日', '受験生'],
    ['H2', '西高', '田中', 3, '', 2, '', '', 2, '', '', 2, 2, '', '', '', '', '', '', '', '山田T,西T', '', '', '火19,木19,土17', '月,水', '理系'],
  ];

  sampleData.forEach(data => {
    const row = mainSheet.addRow(data);
    row.alignment = { vertical: 'middle' };
    row.font = { name: 'Meiryo UI', size: 9 };
  });

  // 罫線を追加
  const dataRowCount = sampleData.length + 1;
  for (let row = 1; row <= dataRowCount; row++) {
    for (let col = 1; col <= headers.length; col++) {
      const cell = mainSheet.getCell(row, col);
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }

  // 説明欄を追加（下部）
  const instructionRow = dataRowCount + 2;
  mainSheet.mergeCells(`A${instructionRow}:Z${instructionRow}`);
  const instructionCell = mainSheet.getCell(`A${instructionRow}`);
  instructionCell.value = '【記入例】学年: S4=小4, M1=中1, H2=高2。科目コマ数: 週に必要なコマ数を数字で入力（0または空欄でスキップ）。希望講師/NG講師/NG生徒: 苗字+"T"または生徒名。複数はカンマ区切り。希望時間: 曜日+時刻形式（例: 月17,水19 = 月曜17時,水曜19時）。NG日程: 曜日のカンマ区切り（例: 火,土）';
  instructionCell.font = { size: 9, color: { argb: 'FF666666' }, name: 'Meiryo UI' };
  instructionCell.alignment = { vertical: 'middle', wrapText: true };
  mainSheet.getRow(instructionRow).height = 40;

  // ========================================
  // ファイル保存
  // ========================================
  const outputPath = './生徒コマ数表テンプレート.xlsx';
  await workbook.xlsx.writeFile(outputPath);

  console.log(`✅ テンプレートファイルを作成しました: ${outputPath}\n`);
  console.log('【フォーマット】');
  console.log('  - 1生徒1行形式');
  console.log('  - 実際のコマ数表フォーマットに準拠\n');
  console.log('【列項目】');
  console.log('  学年, 学校名, 生徒名');
  console.log('  科目: 英, 英検, 数, 算, 国, 理, 社, 古, 物, 化, 生, 地, 政, 世, 日');
  console.log('  条件: 希望講師, NG講師, NG生徒, 希望時間, NG日程, 備考\n');
  console.log('【記入方法】');
  console.log('  - 学年: S4=小4, M1=中1, H2=高2');
  console.log('  - 科目コマ数: 週に必要なコマ数を数字で入力');
  console.log('  - 希望講師/NG講師: 苗字+"T"（例: 西T）、カンマ区切り可');
  console.log('  - NG生徒: 生徒名をカンマ区切り（例: 松橋,深澤）');
  console.log('  - 希望時間: 曜日+時刻形式（例: 月17,水19 = 月曜17時,水曜19時）');
  console.log('  - NG日程: 曜日のカンマ区切り（例: 火,土）\n');
  console.log('=== 作成完了 ===');
}

createStudentDemandTemplate().catch(err => {
  console.error('❌ エラー:', err);
  process.exit(1);
});
