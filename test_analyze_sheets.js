const ExcelJS = require('exceljs');

async function analyzeSheets() {
  console.log('=== 生徒・講師情報.xlsx 詳細解析 ===\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./生徒・講師情報.xlsx');

    // ===== シート2: 生徒コマ数表 =====
    console.log('========================================');
    console.log('シート2: 生徒コマ数表');
    console.log('========================================\n');

    const studentSheet = workbook.getWorksheet('生徒コマ数表');

    // ヘッダー行を取得
    const headerRow = studentSheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber - 1] = cell.value;
    });

    console.log('列ヘッダー:');
    headers.forEach((header, index) => {
      if (header) {
        console.log(`  列${index + 1} (${String.fromCharCode(65 + index)}): ${header}`);
      }
    });

    // 全生徒データを読み込み
    const students = [];
    studentSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // ヘッダーをスキップ

      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });

      // 空行または説明行をスキップ
      if (!rowData[0] || !rowData[1] || !rowData[2]) return;
      if (String(rowData[2]).includes('【記入例】')) return;

      students.push({
        rowNumber: rowNumber,
        grade: rowData[0],
        schoolName: rowData[1],
        studentName: rowData[2],
        英: rowData[3],
        英検: rowData[4],
        数: rowData[5],
        算: rowData[6],
        国: rowData[7],
        理: rowData[8],
        社: rowData[9],
        古: rowData[10],
        物: rowData[11],
        化: rowData[12],
        生: rowData[13],
        地: rowData[15],
        政: rowData[16],
        世: rowData[17],
        日: rowData[18],
        希望講師: rowData[20],
        NG講師: rowData[21],
        NG生徒: rowData[22],
        希望時間: rowData[23],
        NG日程: rowData[24],
        備考: rowData[25]
      });
    });

    console.log(`\n総生徒数: ${students.length}名\n`);

    // サンプル生徒を詳細表示
    console.log('サンプルデータ（最初の5名）:');
    students.slice(0, 5).forEach(student => {
      console.log(`\n[${student.studentName}] (${student.grade}, ${student.schoolName})`);

      // 科目コマ数
      const subjects = [];
      if (student.英) subjects.push(`英語:${student.英}`);
      if (student.英検) subjects.push(`英検:${student.英検}`);
      if (student.数) subjects.push(`数学:${student.数}`);
      if (student.算) subjects.push(`算数:${student.算}`);
      if (student.国) subjects.push(`国語:${student.国}`);
      if (student.理) subjects.push(`理科:${student.理}`);
      if (student.社) subjects.push(`社会:${student.社}`);
      if (student.古) subjects.push(`古文:${student.古}`);
      if (student.物) subjects.push(`物理:${student.物}`);
      if (student.化) subjects.push(`化学:${student.化}`);
      if (student.生) subjects.push(`生物:${student.生}`);
      if (student.地) subjects.push(`地理:${student.地}`);
      if (student.政) subjects.push(`政経:${student.政}`);
      if (student.世) subjects.push(`世界史:${student.世}`);
      if (student.日) subjects.push(`日本史:${student.日}`);

      console.log(`  科目: ${subjects.join(', ')}`);
      if (student.希望講師) console.log(`  希望講師: ${student.希望講師}`);
      if (student.NG講師) console.log(`  NG講師: ${student.NG講師}`);
      if (student.NG生徒) console.log(`  NG生徒: ${student.NG生徒}`);
      if (student.希望時間) console.log(`  希望時間: ${student.希望時間}`);
      if (student.NG日程) console.log(`  NG日程: ${student.NG日程}`);
      if (student.備考) console.log(`  備考: ${student.備考}`);
    });

    // ===== シート1: 指導可能教科一覧 =====
    console.log('\n\n========================================');
    console.log('シート1: 指導可能教科一覧');
    console.log('========================================\n');

    const teacherSheet = workbook.getWorksheet('指導可能教科一覧');

    // 構造を解析
    console.log('シート構造の解析:');

    const rows = [];
    teacherSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });
      rows[rowNumber - 1] = rowData;
    });

    // カテゴリ行（行2）とヘッダー行（行3）を確認
    console.log('\n行2（カテゴリ行）:');
    console.log(rows[1].slice(0, 20));

    console.log('\n行3（科目ヘッダー行）:');
    console.log(rows[2].slice(0, 20));

    // 講師データを解析
    console.log('\n講師リストと指導可能科目:');
    const teachers = [];
    for (let rowIdx = 3; rowIdx < rows.length; rowIdx++) {
      const row = rows[rowIdx];
      if (!row || !row[1]) continue; // 講師名が空の場合スキップ

      const teacherName = String(row[1]).trim();
      if (!teacherName || teacherName === '講師名') continue;

      // 各列の科目を確認
      const subjects = [];
      for (let colIdx = 2; colIdx < row.length; colIdx++) {
        if (row[colIdx] === '◯' || row[colIdx] === '○') {
          const subjectHeader = rows[2][colIdx];
          const categoryHeader = rows[1][colIdx];
          if (subjectHeader) {
            subjects.push({
              subject: subjectHeader,
              category: categoryHeader || '（カテゴリなし）'
            });
          }
        }
      }

      teachers.push({
        name: teacherName,
        subjects: subjects
      });
    }

    console.log(`\n総講師数: ${teachers.length}名\n`);

    // サンプル講師を詳細表示
    console.log('サンプルデータ（最初の5名）:');
    teachers.slice(0, 5).forEach(teacher => {
      console.log(`\n[${teacher.name}]`);
      const subjectsByCategory = {};
      teacher.subjects.forEach(s => {
        if (!subjectsByCategory[s.category]) {
          subjectsByCategory[s.category] = [];
        }
        subjectsByCategory[s.category].push(s.subject);
      });
      Object.entries(subjectsByCategory).forEach(([category, subjects]) => {
        console.log(`  ${category}: ${subjects.join(', ')}`);
      });
    });

    console.log('\n=== 解析完了 ===');

  } catch (error) {
    console.error('❌ エラー:', error);
    console.error(error.stack);
    process.exit(1);
  }
}

analyzeSheets();
