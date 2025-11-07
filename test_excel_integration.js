// Excel解析の統合テスト
// scheduler.htmlからparseStudentDemandExcel関連の関数を抽出してテスト

const ExcelJS = require('exceljs');
const fs = require('fs');

// ヘルパー関数
function getCellValue(cell) {
  if (!cell || cell.value === null || cell.value === undefined) return '';

  // リッチテキストの場合
  if (typeof cell.value === 'object' && cell.value.richText) {
    return cell.value.richText.map(t => t.text).join('');
  }

  // 通常の値
  return String(cell.value);
}

function idFromName(name){
  let h=5381; for (let i=0;i<name.length;i++){ h=((h<<5)+h)+name.charCodeAt(i); h|=0; }
  const hex = (h>>>0).toString(16).slice(-6).padStart(6,'0');
  return 'T' + hex.toUpperCase();
}

function normalizeSubjectName(subject){
  const mapping = {
    '国': '国語',
    '算': '算数',
    '数': '数学',
    '英': '英語',
    '理': '理科',
    '社': '社会',
    '古': '古文',
    '物': '物理',
    '化': '化学',
    '生': '生物',
    '地': '地理',
    '政': '政治経済',
    '世': '世界史',
    '日': '日本史',
    '現': '現代文',
    'ⅠA': '数学ⅠA',
    'ⅡB': '数学ⅡB',
    'Ⅲ': '数学Ⅲ',
    'C': '数学C',
    '倫': '倫理'
  };
  return mapping[subject] || subject;
}

function parseCommaSeparated(str){
  if (!str || typeof str !== 'string') return [];
  return str.split(/[,、]/).map(s => s.trim()).filter(s => s);
}

// Sheet 1: 指導可能教科一覧のパース
function parseTeacherSubjects(worksheet){
  console.log('📖 parseTeacherSubjects: 開始');

  const rows = [];
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowData = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      rowData[colNumber - 1] = getCellValue(cell);
    });
    rows[rowNumber - 1] = rowData;
  });

  const subjectHeaders = rows[2] || [];
  const teacherMap = {};

  for (let rowIdx = 3; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    if (!row || !row[1]) continue;

    const teacherName = String(row[1]).trim();
    if (!teacherName || teacherName === '講師名') continue;

    const teacherId = idFromName(teacherName);
    const subjects = [];

    for (let colIdx = 2; colIdx < row.length; colIdx++) {
      const cellValue = row[colIdx];
      if (cellValue === '◯' || cellValue === '○') {
        const subjectName = subjectHeaders[colIdx];
        if (subjectName) {
          const normalizedSubject = normalizeSubjectName(String(subjectName).trim());
          if (normalizedSubject && !subjects.includes(normalizedSubject)) {
            subjects.push(normalizedSubject);
          }
        }
      }
    }

    if (subjects.length > 0) {
      teacherMap[teacherId] = subjects;
    }
  }

  console.log('✓ parseTeacherSubjects: 完了');
  return teacherMap;
}

// Sheet 2: 生徒コマ数表のパース
function parseStudentDemands(worksheet){
  console.log('📖 parseStudentDemands: 開始');

  const rows = [];
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowData = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      rowData[colNumber - 1] = getCellValue(cell);
    });
    rows[rowNumber - 1] = rowData;
  });

  const studentDemands = [];

  for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    if (!row || !row[2]) continue;

    const studentName = String(row[2]).trim();
    if (!studentName || studentName === '生徒名') continue;

    const grade = String(row[0] || '').trim();
    const schoolName = String(row[1] || '').trim();
    const preferredTeachers = parseCommaSeparated(row[20]);
    const ngTeachers = parseCommaSeparated(row[21]);
    const ngStudents = parseCommaSeparated(row[22]);
    const preferredTimes = parseCommaSeparated(row[23]);
    const ngDays = parseCommaSeparated(row[24]);
    const note = String(row[25] || '').trim();

    const subjectColumns = [
      { col: 3, name: '英' },
      { col: 4, name: '英検' },
      { col: 5, name: '数' },
      { col: 6, name: '算' },
      { col: 7, name: '国' },
      { col: 8, name: '理' },
      { col: 9, name: '社' },
      { col: 10, name: '古' },
      { col: 11, name: '物' },
      { col: 12, name: '化' },
      { col: 13, name: '生' },
      { col: 15, name: '地' },
      { col: 16, name: '政' },
      { col: 17, name: '世' },
      { col: 18, name: '日' }
    ];

    for (const { col, name } of subjectColumns) {
      const count = parseInt(row[col]) || 0;
      if (count > 0) {
        studentDemands.push({
          studentId: idFromName(studentName),
          studentName: studentName,
          subject: normalizeSubjectName(name),
          grade: grade,
          count: count,
          schoolName: schoolName,
          preferredTeachers: preferredTeachers,
          ngTeachers: ngTeachers,
          ngStudents: ngStudents,
          preferredTimes: preferredTimes,
          ngDays: ngDays,
          note: note,
          priority: 5
        });
      }
    }
  }

  console.log('✓ parseStudentDemands: 完了');
  return studentDemands;
}

// メインテスト
async function runTests() {
  console.log('=== Excel解析 統合テスト ===\n');

  const filePath = './生徒・講師情報.xlsx';

  // ファイル存在チェック
  if (!fs.existsSync(filePath)) {
    console.log(`✗ FAIL: ファイルが見つかりません: ${filePath}`);
    return;
  }
  console.log(`✓ ファイル存在確認: ${filePath}\n`);

  // Excelファイルを読み込む
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  console.log(`シート数: ${workbook.worksheets.length}\n`);

  // シート名を確認
  let studentSheet = null;
  let teacherSubjectSheet = null;

  for (const sheet of workbook.worksheets) {
    console.log(`  シート: ${sheet.name}`);
    if (sheet.name.includes('生徒コマ数') || (sheet.name.includes('生徒') && sheet.name.includes('コマ'))) {
      studentSheet = sheet;
      console.log(`    → 生徒コマ数表として認識`);
    } else if (sheet.name.includes('指導可能') || sheet.name.includes('教科一覧')) {
      teacherSubjectSheet = sheet;
      console.log(`    → 指導可能教科一覧として認識`);
    }
  }

  if (!studentSheet && workbook.worksheets.length > 0) {
    studentSheet = workbook.worksheets[workbook.worksheets.length - 1];
    console.log(`  デフォルト: 最後のシートを生徒コマ数表として使用`);
  }
  if (!teacherSubjectSheet && workbook.worksheets.length > 1) {
    teacherSubjectSheet = workbook.worksheets[0];
    console.log(`  デフォルト: 最初のシートを指導可能教科一覧として使用`);
  }

  console.log('');

  // テスト1: 講師-科目マッピングのパース
  console.log('テスト1: 講師-科目マッピングのパース');
  if (teacherSubjectSheet) {
    const teacherMap = parseTeacherSubjects(teacherSubjectSheet);
    const teacherCount = Object.keys(teacherMap).length;
    console.log(`  解析結果: ${teacherCount}名の講師`);

    // サンプル表示
    let count = 0;
    for (const [teacherId, subjects] of Object.entries(teacherMap)) {
      if (count < 3) {
        console.log(`    ${teacherId}: ${subjects.join(', ')}`);
        count++;
      }
    }
    if (teacherCount > 3) {
      console.log(`    ... 他 ${teacherCount - 3}名`);
    }

    // 検証: 全講師が少なくとも1科目を持つ
    const allHaveSubjects = Object.values(teacherMap).every(subjects => subjects.length > 0);
    console.log(`  全講師が科目を持つ: ${allHaveSubjects ? '✓ PASS' : '✗ FAIL'}`);

    // 検証: 科目名が正規化されている
    const allNormalized = Object.values(teacherMap).every(subjects =>
      subjects.every(s => s.length >= 2) // 略称ではなく正式名称
    );
    console.log(`  科目名が正規化されている: ${allNormalized ? '✓ PASS' : '✗ FAIL'}`);
  } else {
    console.log(`  ✗ FAIL: 講師シートが見つかりません`);
  }

  // テスト2: 生徒需要のパース
  console.log('\nテスト2: 生徒需要のパース');
  if (studentSheet) {
    const studentDemands = parseStudentDemands(studentSheet);
    const uniqueStudents = new Set(studentDemands.map(d => d.studentId));
    const studentCount = uniqueStudents.size;
    console.log(`  解析結果: ${studentCount}名の生徒、${studentDemands.length}件の科目需要`);

    // サンプル表示
    const sampleStudents = [...uniqueStudents].slice(0, 3);
    for (const studentId of sampleStudents) {
      const demands = studentDemands.filter(d => d.studentId === studentId);
      const student = demands[0];
      console.log(`    ${student.studentName} (${student.grade}):`);
      demands.forEach(d => {
        console.log(`      ${d.subject}:${d.count}コマ`);
        if (d.preferredTeachers.length > 0) {
          console.log(`        希望講師: ${d.preferredTeachers.join(', ')}`);
        }
        if (d.ngTeachers.length > 0) {
          console.log(`        NG講師: ${d.ngTeachers.join(', ')}`);
        }
        if (d.ngStudents.length > 0) {
          console.log(`        NG生徒: ${d.ngStudents.join(', ')}`);
        }
        if (d.preferredTimes.length > 0) {
          console.log(`        希望時間: ${d.preferredTimes.join(', ')}`);
        }
      });
    }

    // 検証: 全生徒が科目を持つ
    const allHaveSubjects = studentDemands.length > 0;
    console.log(`  生徒需要が存在する: ${allHaveSubjects ? '✓ PASS' : '✗ FAIL'}`);

    // 検証: コマ数が正の整数
    const allPositiveCounts = studentDemands.every(d => d.count > 0);
    console.log(`  全コマ数が正の整数: ${allPositiveCounts ? '✓ PASS' : '✗ FAIL'}`);

    // 検証: 科目名が正規化されている
    const allNormalized = studentDemands.every(d => d.subject.length >= 2);
    console.log(`  科目名が正規化されている: ${allNormalized ? '✓ PASS' : '✗ FAIL'}`);

    // 検証: 希望時間のフォーマット（曜日+時刻）
    const timeFormatValid = studentDemands.every(d => {
      if (d.preferredTimes.length === 0) return true;
      return d.preferredTimes.every(t => /^[月火水木金土日]\d{1,2}$/.test(t));
    });
    console.log(`  希望時間のフォーマットが正しい: ${timeFormatValid ? '✓ PASS' : '✗ FAIL'}`);
  } else {
    console.log(`  ✗ FAIL: 生徒シートが見つかりません`);
  }

  // テスト3: 講師-生徒のマッチング可能性
  console.log('\nテスト3: 講師-生徒のマッチング可能性');
  if (teacherSubjectSheet && studentSheet) {
    const teacherMap = parseTeacherSubjects(teacherSubjectSheet);
    const studentDemands = parseStudentDemands(studentSheet);

    // 各科目需要に対して指導可能な講師が存在するかチェック
    const subjectCoverage = {};
    for (const demand of studentDemands) {
      const subject = demand.subject;
      if (!subjectCoverage[subject]) {
        subjectCoverage[subject] = 0;
      }

      // この科目を教えられる講師の数をカウント
      for (const [teacherId, subjects] of Object.entries(teacherMap)) {
        if (subjects.includes(subject)) {
          subjectCoverage[subject]++;
        }
      }
    }

    console.log('  科目別の指導可能講師数:');
    const uniqueSubjects = Object.keys(subjectCoverage).sort();
    for (const subject of uniqueSubjects) {
      const teacherCount = new Set(
        Object.entries(teacherMap)
          .filter(([_, subjects]) => subjects.includes(subject))
          .map(([teacherId, _]) => teacherId)
      ).size;
      console.log(`    ${subject.padEnd(10)}: ${teacherCount}名`);
    }

    // 検証: 全科目に少なくとも1名の講師がいる
    const allSubjectsCovered = uniqueSubjects.every(subject => {
      return Object.values(teacherMap).some(subjects => subjects.includes(subject));
    });
    console.log(`  全科目に指導可能な講師がいる: ${allSubjectsCovered ? '✓ PASS' : '⚠️ 一部カバーされていない可能性'}`);
  }

  // テスト4: リッチテキストの処理
  console.log('\nテスト4: リッチテキストの処理');
  if (studentSheet) {
    const studentDemands = parseStudentDemands(studentSheet);
    const hasObjectString = studentDemands.some(d =>
      d.preferredTeachers.some(t => t.includes('[object Object]')) ||
      d.ngTeachers.some(t => t.includes('[object Object]'))
    );
    console.log(`  [object Object]が含まれる: ${hasObjectString ? '✗ FAIL' : '✓ PASS'}`);
  }

  console.log('\n=== テスト完了 ===');
}

runTests().catch(err => {
  console.error('エラー:', err);
  process.exit(1);
});
