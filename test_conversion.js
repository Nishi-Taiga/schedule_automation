// ブース番号変換ロジックのテスト

function circledToBoothId(ch){
  // 丸付き数字（①-⑳）を B1, B2... に変換
  const code = ch.charCodeAt(0);
  if (code >= 0x2460 && code <= 0x2473) {
    const num = code - 0x2460 + 1;
    return `B${num}`;
  }
  return ch; // 丸付き数字でない場合はそのまま返す
}

function boothIdToCircled(boothId){
  // B1, B2... を丸付き数字（①②...）に変換
  const match = boothId.match(/^B(\d+)$/);
  if (match) {
    const num = parseInt(match[1], 10);
    if (num >= 1 && num <= 20) {
      return String.fromCharCode(0x2460 + num - 1);
    }
  }
  return boothId; // B1形式でない場合はそのまま返す
}

console.log('=== ブース番号変換テスト ===\n');

// circledToBoothId のテスト
console.log('【丸付き数字 → B1形式】');
const circledNumbers = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
for (const ch of circledNumbers) {
  const result = circledToBoothId(ch);
  console.log(`  ${ch} → ${result}`);
}

console.log('\n【B1形式 → 丸付き数字】');
for (let i = 1; i <= 10; i++) {
  const boothId = `B${i}`;
  const result = boothIdToCircled(boothId);
  console.log(`  ${boothId} → ${result}`);
}

console.log('\n【往復変換テスト】');
for (const ch of circledNumbers) {
  const b1 = circledToBoothId(ch);
  const back = boothIdToCircled(b1);
  const ok = (ch === back) ? '✓' : '✗';
  console.log(`  ${ch} → ${b1} → ${back} ${ok}`);
}

console.log('\n✅ すべてのテスト完了');
