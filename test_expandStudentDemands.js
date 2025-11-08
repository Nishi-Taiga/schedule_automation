// expandStudentDemands関数の統合テスト
// scheduler.htmlから抽出した関数

function pad2(n){ return String(n).padStart(2,'0'); }

function idFromName(name){
  // simple stable hash -> Txxxxxx
  let h=5381; for (let i=0;i<name.length;i++){ h=((h<<5)+h)+name.charCodeAt(i); h|=0; }
  const hex = (h>>>0).toString(16).slice(-6).padStart(6,'0');
  return 'T' + hex.toUpperCase();
}

function expandStudentDemands(studentDemands, slots){
  if (!studentDemands || studentDemands.length === 0) return [];

  const expanded = [];
  const dayMap = { '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6, '日': 0 };

  // スロットから利用可能な日付と時刻の組み合わせを取得
  const availableSlots = new Set();
  for (const slot of slots) {
    availableSlots.add(`${slot.date}|${slot.time}`);
  }

  for (const demand of studentDemands) {
    // 希望時間が指定されている場合
    if (demand.preferredTimes && demand.preferredTimes.length > 0) {
      // 希望時間をパース（例: "月17" -> 曜日: 月, 時刻: 17:00:00）
      for (const timeSpec of demand.preferredTimes) {
        const dayChar = timeSpec.charAt(0);
        const hourStr = timeSpec.substring(1);
        const hour = parseInt(hourStr);

        if (!dayMap.hasOwnProperty(dayChar) || isNaN(hour)) {
          console.warn(`無効な希望時間フォーマット: ${timeSpec}`);
          continue;
        }

        const targetDayOfWeek = dayMap[dayChar];
        const timeStr = `${pad2(hour)}:00:00`;

        // スロットから該当する日付を探す
        for (const slot of slots) {
          const dateObj = new Date(slot.date);
          const dayOfWeek = dateObj.getDay();

          if (dayOfWeek === targetDayOfWeek && slot.time === timeStr) {
            // NG日程のチェック
            if (demand.ngDays && demand.ngDays.includes(dayChar)) {
              continue; // NG日程なのでスキップ
            }

            // NGTeachersをフォーマット（講師名からID形式に変換）
            const ngTeachersSet = new Set();
            if (demand.ngTeachers) {
              for (const teacherName of demand.ngTeachers) {
                // "西T" -> "nishiT" のような形式に変換
                ngTeachersSet.add(idFromName(teacherName));
              }
            }

            expanded.push({
              date: slot.date,
              time: slot.time,
              studentId: demand.studentId,
              studentName: demand.studentName,
              subject: demand.subject,
              grade: demand.grade,
              preferredTeacherId: demand.preferredTeachers && demand.preferredTeachers.length > 0
                ? idFromName(demand.preferredTeachers[0])
                : null,
              ngTeachers: Array.from(ngTeachersSet),
              ngStudents: demand.ngStudents || [],
              priority: demand.priority || 5
            });
          }
        }
      }
    } else {
      // 希望時間が指定されていない場合は全スロットに展開
      for (const slot of slots) {
        const dateObj = new Date(slot.date);
        const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
        const dayChar = dayNames[dateObj.getDay()];

        // NG日程のチェック
        if (demand.ngDays && demand.ngDays.includes(dayChar)) {
          continue;
        }

        const ngTeachersSet = new Set();
        if (demand.ngTeachers) {
          for (const teacherName of demand.ngTeachers) {
            ngTeachersSet.add(idFromName(teacherName));
          }
        }

        expanded.push({
          date: slot.date,
          time: slot.time,
          studentId: demand.studentId,
          studentName: demand.studentName,
          subject: demand.subject,
          grade: demand.grade,
          preferredTeacherId: demand.preferredTeachers && demand.preferredTeachers.length > 0
            ? idFromName(demand.preferredTeachers[0])
            : null,
          ngTeachers: Array.from(ngTeachersSet),
          ngStudents: demand.ngStudents || [],
          priority: demand.priority || 5
        });
      }
    }
  }

  return expanded;
}

console.log('=== expandStudentDemands 統合テスト ===\n');

// テストデータ: スロット（2025年1月の1週間分）
const testSlots = [
  // 月曜日 2025-01-06
  { date: '2025-01-06', time: '16:00:00', boothId: 'A' },
  { date: '2025-01-06', time: '17:00:00', boothId: 'A' },
  { date: '2025-01-06', time: '18:00:00', boothId: 'A' },
  // 火曜日 2025-01-07
  { date: '2025-01-07', time: '16:00:00', boothId: 'A' },
  { date: '2025-01-07', time: '17:00:00', boothId: 'A' },
  { date: '2025-01-07', time: '18:00:00', boothId: 'A' },
  // 水曜日 2025-01-08
  { date: '2025-01-08', time: '16:00:00', boothId: 'A' },
  { date: '2025-01-08', time: '17:00:00', boothId: 'A' },
  { date: '2025-01-08', time: '18:00:00', boothId: 'A' },
  // 木曜日 2025-01-09
  { date: '2025-01-09', time: '16:00:00', boothId: 'A' },
  { date: '2025-01-09', time: '17:00:00', boothId: 'A' },
  { date: '2025-01-09', time: '18:00:00', boothId: 'A' },
  // 金曜日 2025-01-10
  { date: '2025-01-10', time: '16:00:00', boothId: 'A' },
  { date: '2025-01-10', time: '17:00:00', boothId: 'A' },
  { date: '2025-01-10', time: '18:00:00', boothId: 'A' },
];

// テスト1: 希望時間が指定されている場合
console.log('テスト1: 希望時間が指定されている場合');
const demand1 = {
  studentId: 'S001',
  studentName: '松橋',
  subject: '算数',
  grade: 'S4',
  preferredTeachers: ['西T'],
  ngTeachers: [],
  ngStudents: [],
  preferredTimes: ['月17', '水17', '金17'], // 月水金の17時
  ngDays: [],
  priority: 5
};

const expanded1 = expandStudentDemands([demand1], testSlots);
console.log(`  入力: 生徒${demand1.studentName}, 希望時間=${demand1.preferredTimes.join(',')}`);
console.log(`  出力: ${expanded1.length}件のスロットに展開`);
console.log('  展開結果:');
expanded1.forEach(slot => {
  const date = new Date(slot.date);
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dayName = dayNames[date.getDay()];
  console.log(`    ${slot.date} (${dayName}) ${slot.time} - ${slot.studentName} (${slot.subject})`);
});
console.log(`  期待値: 3件（月17, 水17, 金17）`);
console.log(`  結果: ${expanded1.length === 3 ? '✓ PASS' : '✗ FAIL'}`);

// テスト2: NG講師が指定されている場合
console.log('\nテスト2: NG講師が指定されている場合');
const demand2 = {
  studentId: 'S002',
  studentName: '深澤',
  subject: '算数',
  grade: 'S4',
  preferredTeachers: [],
  ngTeachers: ['田中T', '佐藤T'],
  ngStudents: ['松橋'],
  preferredTimes: ['月17', '火17'],
  ngDays: [],
  priority: 5
};

const expanded2 = expandStudentDemands([demand2], testSlots);
console.log(`  入力: 生徒${demand2.studentName}, NG講師=${demand2.ngTeachers.join(',')}, NG生徒=${demand2.ngStudents.join(',')}`);
console.log(`  出力: ${expanded2.length}件のスロットに展開`);
expanded2.forEach(slot => {
  console.log(`    NG講師ID: ${slot.ngTeachers.join(', ')}`);
  console.log(`    NG生徒: ${slot.ngStudents.join(', ')}`);
});
const expectedNgTeacherIds = demand2.ngTeachers.map(t => idFromName(t));
console.log(`  期待されるNG講師ID: ${expectedNgTeacherIds.join(', ')}`);
const allNgTeachersMatch = expanded2.every(slot =>
  slot.ngTeachers.length === expectedNgTeacherIds.length &&
  slot.ngTeachers.every(id => expectedNgTeacherIds.includes(id))
);
console.log(`  結果: ${allNgTeachersMatch ? '✓ PASS' : '✗ FAIL'}`);

// テスト3: NG日程が指定されている場合
console.log('\nテスト3: NG日程が指定されている場合');
const demand3 = {
  studentId: 'S003',
  studentName: '山口',
  subject: '算数',
  grade: 'S4',
  preferredTeachers: [],
  ngTeachers: [],
  ngStudents: [],
  preferredTimes: ['月17', '火17', '水17', '木17', '金17'], // 全曜日指定
  ngDays: ['火', '木'], // 火木はNG
  priority: 5
};

const expanded3 = expandStudentDemands([demand3], testSlots);
console.log(`  入力: 生徒${demand3.studentName}, 希望時間=${demand3.preferredTimes.join(',')}, NG日程=${demand3.ngDays.join(',')}`);
console.log(`  出力: ${expanded3.length}件のスロットに展開`);
console.log('  展開結果:');
expanded3.forEach(slot => {
  const date = new Date(slot.date);
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dayName = dayNames[date.getDay()];
  console.log(`    ${slot.date} (${dayName}) ${slot.time}`);
});
const hasNgDay = expanded3.some(slot => {
  const date = new Date(slot.date);
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dayName = dayNames[date.getDay()];
  return demand3.ngDays.includes(dayName);
});
console.log(`  期待値: 3件（月17, 水17, 金17）、NG日程（火木）は除外`);
console.log(`  結果: ${expanded3.length === 3 && !hasNgDay ? '✓ PASS' : '✗ FAIL'}`);

// テスト4: 希望時間が指定されていない場合（全スロット展開）
console.log('\nテスト4: 希望時間が指定されていない場合（全スロット展開）');
const demand4 = {
  studentId: 'S004',
  studentName: '鈴木',
  subject: '国語',
  grade: 'S5',
  preferredTeachers: [],
  ngTeachers: [],
  ngStudents: [],
  preferredTimes: [], // 希望時間なし
  ngDays: [],
  priority: 5
};

const expanded4 = expandStudentDemands([demand4], testSlots);
console.log(`  入力: 生徒${demand4.studentName}, 希望時間なし`);
console.log(`  出力: ${expanded4.length}件のスロットに展開`);
console.log(`  期待値: ${testSlots.length}件（全スロット）`);
console.log(`  結果: ${expanded4.length === testSlots.length ? '✓ PASS' : '✗ FAIL'}`);

// テスト5: 複数生徒の同時展開
console.log('\nテスト5: 複数生徒の同時展開');
const demands = [demand1, demand2, demand3, demand4];
const expandedAll = expandStudentDemands(demands, testSlots);
console.log(`  入力: ${demands.length}名の生徒`);
console.log(`  出力: ${expandedAll.length}件のスロット需要`);
const expectedTotal = expanded1.length + expanded2.length + expanded3.length + expanded4.length;
console.log(`  期待値: ${expectedTotal}件`);
console.log(`  結果: ${expandedAll.length === expectedTotal ? '✓ PASS' : '✗ FAIL'}`);

// テスト6: 無効なフォーマットの処理
console.log('\nテスト6: 無効なフォーマットの処理');
const demand6 = {
  studentId: 'S006',
  studentName: 'テスト生徒',
  subject: '数学',
  grade: 'S5',
  preferredTeachers: [],
  ngTeachers: [],
  ngStudents: [],
  preferredTimes: ['月17', '無効99', 'XYZ', '火18'], // 一部無効なフォーマット
  ngDays: [],
  priority: 5
};

const expanded6 = expandStudentDemands([demand6], testSlots);
console.log(`  入力: 希望時間=${demand6.preferredTimes.join(',')}`);
console.log(`  出力: ${expanded6.length}件のスロットに展開`);
console.log(`  期待値: 2件（月17, 火18）、無効なフォーマットは無視`);
console.log(`  結果: ${expanded6.length === 2 ? '✓ PASS' : '✗ FAIL'}`);

// テスト7: 希望講師のID変換
console.log('\nテスト7: 希望講師のID変換');
const demand7 = {
  studentId: 'S007',
  studentName: 'テスト生徒2',
  subject: '英語',
  grade: 'S6',
  preferredTeachers: ['西T', '橋本T'], // 複数指定（最初のみ使用）
  ngTeachers: [],
  ngStudents: [],
  preferredTimes: ['月17'],
  ngDays: [],
  priority: 5
};

const expanded7 = expandStudentDemands([demand7], testSlots);
console.log(`  入力: 希望講師=${demand7.preferredTeachers.join(',')}`);
if (expanded7.length > 0) {
  console.log(`  出力: preferredTeacherId=${expanded7[0].preferredTeacherId}`);
  const expectedId = idFromName(demand7.preferredTeachers[0]);
  console.log(`  期待値: ${expectedId} (${demand7.preferredTeachers[0]})`);
  console.log(`  結果: ${expanded7[0].preferredTeacherId === expectedId ? '✓ PASS' : '✗ FAIL'}`);
} else {
  console.log(`  ✗ FAIL: 展開結果が空`);
}

console.log('\n=== テスト完了 ===');
