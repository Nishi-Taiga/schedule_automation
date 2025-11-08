// idFromName関数のテスト
// scheduler.htmlから抽出した関数

function idFromName(name){
  // simple stable hash -> Txxxxxx
  let h=5381; for (let i=0;i<name.length;i++){ h=((h<<5)+h)+name.charCodeAt(i); h|=0; }
  const hex = (h>>>0).toString(16).slice(-6).padStart(6,'0');
  return 'T' + hex.toUpperCase();
}

console.log('=== idFromName関数テスト ===\n');

// テスト1: 基本的な変換
console.log('1. 基本的な変換テスト');
const basicTests = [
  '西T',
  '田中T',
  '佐藤T',
  '鈴木T',
  '高橋T',
  '伊藤T',
  '飯村T',
  '宗岡T'
];

console.log('講師名 -> teacherId:');
basicTests.forEach(name => {
  const id = idFromName(name);
  console.log(`  ${name.padEnd(10)} -> ${id}`);
});

// テスト2: 一貫性チェック（同じ名前は同じIDを生成）
console.log('\n2. 一貫性チェック（同じ名前は同じIDを生成）');
const name = '西T';
const id1 = idFromName(name);
const id2 = idFromName(name);
const id3 = idFromName(name);
console.log(`  ${name} -> ${id1}`);
console.log(`  ${name} -> ${id2}`);
console.log(`  ${name} -> ${id3}`);
console.log(`  一貫性: ${id1 === id2 && id2 === id3 ? '✓ PASS' : '✗ FAIL'}`);

// テスト3: 衝突チェック（異なる名前は異なるIDを生成）
console.log('\n3. 衝突チェック（異なる名前は異なるIDを生成）');
const names = ['西T', '田中T', '佐藤T', '鈴木T', '高橋T'];
const ids = names.map(n => idFromName(n));
const uniqueIds = new Set(ids);
console.log(`  テスト対象: ${names.length}名`);
console.log(`  ユニークなID数: ${uniqueIds.size}`);
console.log(`  衝突なし: ${names.length === uniqueIds.size ? '✓ PASS' : '✗ FAIL'}`);

// テスト4: 実際のデータでのマッピング確認
console.log('\n4. 実際のデータでのマッピング確認');

// 講師リスト（実際のデータから）
const teacherNames = [
  '飯村T', '稲田T', '宇佐T', '大久保T', '大澤T', '加藤T', '熊野T', '小林T',
  '島田T', '田中T', '西T', '橋本T', '福島T', '細木T', '宗岡T', '横松T'
];

// 生徒の希望講師やNG講師の例
const studentPreferences = [
  { student: '松橋', preferred: '西T', ng: [] },
  { student: '深澤', preferred: '', ng: ['田中T'] },
  { student: '宗岡', preferred: '西T, 橋本T', ng: [] }
];

console.log('講師IDマッピング:');
const teacherIdMap = {};
teacherNames.forEach(name => {
  const id = idFromName(name);
  teacherIdMap[name] = id;
  console.log(`  ${name.padEnd(12)} -> ${id}`);
});

console.log('\n生徒の希望講師/NG講師のID変換:');
studentPreferences.forEach(({ student, preferred, ng }) => {
  console.log(`\n  生徒: ${student}`);
  if (preferred) {
    const preferredList = preferred.split(',').map(s => s.trim()).filter(s => s);
    preferredList.forEach(name => {
      const id = idFromName(name);
      console.log(`    希望講師: ${name.padEnd(12)} -> ${id}`);
      // マッピング確認
      if (teacherIdMap[name]) {
        console.log(`      マッチング: ${teacherIdMap[name] === id ? '✓ OK' : '✗ NG'}`);
      }
    });
  }
  if (ng.length > 0) {
    ng.forEach(name => {
      const id = idFromName(name);
      console.log(`    NG講師:   ${name.padEnd(12)} -> ${id}`);
      // マッピング確認
      if (teacherIdMap[name]) {
        console.log(`      マッチング: ${teacherIdMap[name] === id ? '✓ OK' : '✗ NG'}`);
      }
    });
  }
});

// テスト5: エッジケース
console.log('\n5. エッジケース');
const edgeCases = [
  { name: '', description: '空文字列' },
  { name: 'T', description: '1文字' },
  { name: '西　泰我T', description: '全角スペース含む' },
  { name: '西 泰我T', description: '半角スペース含む' },
  { name: 'にしT', description: 'ひらがな' },
  { name: 'ニシT', description: 'カタカナ' }
];

console.log('エッジケーステスト:');
edgeCases.forEach(({ name, description }) => {
  const id = idFromName(name);
  console.log(`  ${description.padEnd(20)} "${name}" -> ${id}`);
});

// テスト6: パフォーマンステスト
console.log('\n6. パフォーマンステスト');
const iterations = 100000;
const testName = '西T';
const startTime = Date.now();
for (let i = 0; i < iterations; i++) {
  idFromName(testName);
}
const endTime = Date.now();
const duration = endTime - startTime;
const avgTime = duration / iterations;
console.log(`  ${iterations.toLocaleString()}回の変換を ${duration}ms で完了`);
console.log(`  平均処理時間: ${avgTime.toFixed(6)}ms`);
console.log(`  パフォーマンス: ${avgTime < 0.01 ? '✓ 高速' : '⚠️ 要最適化'}`);

console.log('\n=== テスト完了 ===');
