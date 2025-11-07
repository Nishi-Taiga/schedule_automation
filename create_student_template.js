const ExcelJS = require('exceljs');

async function createStudentDemandTemplate() {
  console.log('=== 生徒コマ数表テンプレート作成開始 ===\n');

  const workbook = new ExcelJS.Workbook();

  // ========================================
  // シート1: 生徒コマ数表
  // ========================================
  const mainSheet = workbook.addWorksheet('生徒コマ数表');

  // ヘッダー行
  const headers = [
    '生徒名',
    '科目',
    'コマ数',
    '希望講師',
    'NG講師',
    '希望曜日',
    '希望時間帯',
    'NG曜日',
    'NG時間帯',
    '備考'
  ];

  mainSheet.addRow(headers);

  // ヘッダー行のスタイル設定
  const headerRow = mainSheet.getRow(1);
  headerRow.font = { bold: true, size: 11, name: 'Meiryo UI' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E1F2' }  // 薄い青
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 20;

  // 列幅設定
  mainSheet.getColumn(1).width = 15;  // 生徒名
  mainSheet.getColumn(2).width = 10;  // 科目
  mainSheet.getColumn(3).width = 8;   // コマ数
  mainSheet.getColumn(4).width = 15;  // 希望講師
  mainSheet.getColumn(5).width = 15;  // NG講師
  mainSheet.getColumn(6).width = 15;  // 希望曜日
  mainSheet.getColumn(7).width = 15;  // 希望時間帯
  mainSheet.getColumn(8).width = 15;  // NG曜日
  mainSheet.getColumn(9).width = 15;  // NG時間帯
  mainSheet.getColumn(10).width = 20; // 備考

  // サンプルデータ
  const sampleData = [
    ['田中太郎', '数学', 3, '西T', '佐藤T,山田T', '月,水,金', '15:00-18:00', '', '', '受験生'],
    ['田中太郎', '英語', 2, '鈴木T', '', '月,水,金', '15:00-18:00', '', '', ''],
    ['佐藤花子', '数学', 2, '', '西T', '火,木', '16:00-19:00', '', '19:00-21:00', '部活あり'],
    ['鈴木一郎', '理科', 1, '山田T', '', '土', '10:00-15:00', '', '', '土曜のみ通塾'],
  ];

  sampleData.forEach(data => {
    const row = mainSheet.addRow(data);
    row.alignment = { vertical: 'middle' };
    row.font = { name: 'Meiryo UI', size: 10 };
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
  mainSheet.mergeCells(`A${instructionRow}:J${instructionRow}`);
  const instructionCell = mainSheet.getCell(`A${instructionRow}`);
  instructionCell.value = '【記入例】希望講師・NG講師: 苗字+"T"で指定（例: 西T, 田中T）。複数指定時はカンマ区切り。希望曜日: 月,火,水,木,金,土,日。希望時間帯: HH:MM-HH:MM形式（例: 15:00-18:00）';
  instructionCell.font = { size: 9, color: { argb: 'FF666666' }, name: 'Meiryo UI' };
  instructionCell.alignment = { vertical: 'middle', wrapText: true };
  mainSheet.getRow(instructionRow).height = 30;

  // ========================================
  // シート2: 隣接NG設定
  // ========================================
  const ngSheet = workbook.addWorksheet('隣接NG設定');

  // ヘッダー行
  const ngHeaders = ['生徒名', '隣接NG生徒名', '理由'];
  ngSheet.addRow(ngHeaders);

  // ヘッダー行のスタイル設定
  const ngHeaderRow = ngSheet.getRow(1);
  ngHeaderRow.font = { bold: true, size: 11, name: 'Meiryo UI' };
  ngHeaderRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFCE4D6' }  // 薄いオレンジ
  };
  ngHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
  ngHeaderRow.height = 20;

  // 列幅設定
  ngSheet.getColumn(1).width = 15;  // 生徒名
  ngSheet.getColumn(2).width = 20;  // 隣接NG生徒名
  ngSheet.getColumn(3).width = 30;  // 理由

  // サンプルデータ
  const ngSampleData = [
    ['田中太郎', '鈴木一郎,高橋次郎', '私語防止'],
    ['佐藤花子', '山田三郎', '集中力維持'],
  ];

  ngSampleData.forEach(data => {
    const row = ngSheet.addRow(data);
    row.alignment = { vertical: 'middle' };
    row.font = { name: 'Meiryo UI', size: 10 };
  });

  // 罫線を追加
  const ngDataRowCount = ngSampleData.length + 1;
  for (let row = 1; row <= ngDataRowCount; row++) {
    for (let col = 1; col <= ngHeaders.length; col++) {
      const cell = ngSheet.getCell(row, col);
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }

  // 説明欄を追加
  const ngInstructionRow = ngDataRowCount + 2;
  ngSheet.mergeCells(`A${ngInstructionRow}:C${ngInstructionRow}`);
  const ngInstructionCell = ngSheet.getCell(`A${ngInstructionRow}`);
  ngInstructionCell.value = '【記入例】隣接NG生徒名: 複数指定時はカンマ区切り（例: 田中太郎,佐藤花子）。同じブースまたは隣のブースに座らせたくない生徒を指定します。';
  ngInstructionCell.font = { size: 9, color: { argb: 'FF666666' }, name: 'Meiryo UI' };
  ngInstructionCell.alignment = { vertical: 'middle', wrapText: true };
  ngSheet.getRow(ngInstructionRow).height = 30;

  // ========================================
  // ファイル保存
  // ========================================
  const outputPath = './生徒コマ数表テンプレート.xlsx';
  await workbook.xlsx.writeFile(outputPath);

  console.log(`✅ テンプレートファイルを作成しました: ${outputPath}\n`);
  console.log('【含まれるシート】');
  console.log('  1. 生徒コマ数表 - 生徒ごとの科目別コマ数と条件');
  console.log('  2. 隣接NG設定 - 隣接させたくない生徒の組み合わせ\n');
  console.log('【記入方法】');
  console.log('  - 希望講師/NG講師: 苗字+"T"で指定（例: 西T）');
  console.log('  - 複数指定: カンマ区切り（例: 西T,田中T）');
  console.log('  - 希望曜日: 月,火,水,木,金,土,日 から選択');
  console.log('  - 希望時間帯: HH:MM-HH:MM形式（例: 15:00-18:00）');
  console.log('  - コマ数: その科目で週に必要なコマ数\n');
  console.log('=== 作成完了 ===');
}

createStudentDemandTemplate().catch(err => {
  console.error('❌ エラー:', err);
  process.exit(1);
});
