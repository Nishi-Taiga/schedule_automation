// 科目マッチング機能のテスト
// scheduler.htmlから関連関数を抽出してテスト

// 科目名を正規化（略称を統一）
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

// ヘルパー関数: 講師が指定科目を指導可能かチェック
function canTeachSubject(teacherId, subject, teacherSubjectMap) {
  // 科目が指定されていない場合は常にOK
  if (!subject) return true;
  // teacherSubjectMapが空の場合は常にOK（後方互換性）
  if (!teacherSubjectMap || Object.keys(teacherSubjectMap).length === 0) return true;
  // 講師の指導可能科目リストを取得
  const teachableSubjects = teacherSubjectMap[teacherId];
  if (!teachableSubjects) return false; // 講師情報がない場合はNG
  // 科目名を正規化して比較
  const normalizedSubject = normalizeSubjectName(subject);
  return teachableSubjects.includes(normalizedSubject);
}

console.log('=== 科目マッチング機能テスト ===\n');

// テストデータ
const teacherSubjectMap = {
  '講師A': ['数学', '英語', '物理'],
  '講師B': ['国語', '古文', '現代文'],
  '講師C': ['算数', '数学', '理科']
};

console.log('講師の指導可能科目:');
Object.entries(teacherSubjectMap).forEach(([id, subjects]) => {
  console.log(`  ${id}: ${subjects.join(', ')}`);
});

console.log('\nテストケース:');

// テスト1: 正確な科目名でマッチング
console.log('\n1. 正確な科目名でマッチング');
console.log(`  講師A が 数学 を教えられるか: ${canTeachSubject('講師A', '数学', teacherSubjectMap)}`); // true
console.log(`  講師A が 国語 を教えられるか: ${canTeachSubject('講師A', '国語', teacherSubjectMap)}`); // false
console.log(`  講師B が 国語 を教えられるか: ${canTeachSubject('講師B', '国語', teacherSubjectMap)}`); // true

// テスト2: 略称でのマッチング
console.log('\n2. 略称でのマッチング（正規化機能）');
console.log(`  講師A が 数 を教えられるか: ${canTeachSubject('講師A', '数', teacherSubjectMap)}`); // true (数 -> 数学)
console.log(`  講師C が 算 を教えられるか: ${canTeachSubject('講師C', '算', teacherSubjectMap)}`); // true (算 -> 算数)
console.log(`  講師B が 古 を教えられるか: ${canTeachSubject('講師B', '古', teacherSubjectMap)}`); // true (古 -> 古文)

// テスト3: 科目が指定されていない場合
console.log('\n3. 科目が指定されていない場合（常にtrue）');
console.log(`  講師A が null を教えられるか: ${canTeachSubject('講師A', null, teacherSubjectMap)}`); // true
console.log(`  講師A が undefined を教えられるか: ${canTeachSubject('講師A', undefined, teacherSubjectMap)}`); // true
console.log(`  講師A が '' を教えられるか: ${canTeachSubject('講師A', '', teacherSubjectMap)}`); // true

// テスト4: teacherSubjectMapが空の場合（後方互換性）
console.log('\n4. teacherSubjectMapが空の場合（後方互換性）');
console.log(`  講師A が 数学 を教えられるか（マップ空）: ${canTeachSubject('講師A', '数学', {})}`); // true
console.log(`  講師A が 数学 を教えられるか（マップなし）: ${canTeachSubject('講師A', '数学', null)}`); // true

// テスト5: 講師情報がない場合
console.log('\n5. 講師情報がない場合（false）');
console.log(`  講師D が 数学 を教えられるか: ${canTeachSubject('講師D', '数学', teacherSubjectMap)}`); // false

// テスト6: 正規化機能の確認
console.log('\n6. 正規化機能の確認');
const testCases = [
  { input: '国', expected: '国語' },
  { input: '算', expected: '算数' },
  { input: '数', expected: '数学' },
  { input: '英', expected: '英語' },
  { input: '物', expected: '物理' },
  { input: '化', expected: '化学' },
  { input: '生', expected: '生物' },
  { input: '地', expected: '地理' },
  { input: '政', expected: '政治経済' },
  { input: '世', expected: '世界史' },
  { input: '日', expected: '日本史' },
  { input: '数学ⅠA', expected: '数学ⅠA' } // マッピングなし
];

testCases.forEach(({ input, expected }) => {
  const result = normalizeSubjectName(input);
  const status = result === expected ? '✓' : '✗';
  console.log(`  ${status} ${input} -> ${result} (期待値: ${expected})`);
});

console.log('\n=== テスト完了 ===');
