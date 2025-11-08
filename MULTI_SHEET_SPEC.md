# 複数シート対応仕様書

## 背景
現在のシステムは単一シート（1週間分）を前提としているが、実際の運用では1週間ごとにシートが分かれているExcelファイルを使用している。
例: 「ブース表 2025.11.02-08」「ブース表 2025.11.09-15」など

## 要件

### 1. ブース表の複数シート対応
**入力**: 複数のシートを持つExcelファイル
```
シート1: ブース表 2025.11.02-08 (第1週)
シート2: ブース表 2025.11.09-15 (第2週)
シート3: ブース表 2025.11.16-22 (第3週)
...
```

**処理**:
1. 全シートを読み込み
2. 各シートから年月日を抽出（シート名または内容から）
3. 全シートのスロットデータを統合
4. 日付順にソート

**出力**: 統合されたスロットデータ
```javascript
[
  { date: '2025/11/02', time: '16:00:00', boothId: 'B1' },
  { date: '2025/11/02', time: '16:00:00', boothId: 'B2' },
  ...
  { date: '2025/11/09', time: '16:00:00', boothId: 'B1' },  // 第2週
  ...
]
```

### 2. 講師表（元シート）の複数シート対応
**入力**: 複数のシートを持つExcelファイル
```
シート1: 11/2-8 (第1週)
シート2: 11/9-15 (第2週)
シート3: 11/16-22 (第3週)
...
```

**処理**:
1. 全シートを読み込み
2. 各シートから講師の可用枠を抽出
3. 全シートのデータを統合

**出力**: 統合された講師可用データ
```javascript
[
  { date: '2025/11/02', time: '16:00:00', teacherId: 'T6B23F8', teacherName: '西T' },
  ...
  { date: '2025/11/09', time: '16:00:00', teacherId: 'T6B23F8', teacherName: '西T' },  // 第2週
  ...
]
```

### 3. シート名の検出パターン

#### パターン1: 日付範囲形式
- `ブース表　2025.11.02-08`
- `11/2-8`
- `2025.11.02-2025.11.08`

#### パターン2: 週番号形式
- `第1週 (11/2-8)`
- `Week 1`

#### パターン3: 月形式
- `2025年11月`
- `11月分`

### 4. 年月の推定優先順位
1. シート名から明示的な年月を検出
2. シート内のセルから年月を検出
3. 前のシートの年月を継承（週を加算）
4. 現在の年月をデフォルト値として使用

### 5. エッジケース処理

#### ケース1: シート名が不明
- デフォルトで全シートを読み込み
- 日付が取得できない場合は警告を表示
- ユーザーに手動で年月を指定させる（UI拡張）

#### ケース2: 重複する日付
- 後のシートのデータで上書き
- 警告メッセージを表示

#### ケース3: 空のシート
- スキップして次のシートへ

## 実装計画

### Phase 1: parseBoothExcel の拡張
```javascript
async function parseBoothExcel(input) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await input.arrayBuffer());

  let allSlots = [];

  // 全シートを処理
  for (const sheet of workbook.worksheets) {
    console.log(`📋 シート処理中: ${sheet.name}`);

    // シート名から年月を検出
    const yearMonth = detectYearMonthFromSheetName(sheet.name);

    // 単一シートとして処理
    const slots = await parseSingleBoothSheet(sheet, yearMonth);
    allSlots = allSlots.concat(slots);
  }

  // 日付順にソート
  allSlots.sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    if (a.time !== b.time) return a.time.localeCompare(b.time);
    return a.boothId.localeCompare(b.boothId);
  });

  return allSlots;
}
```

### Phase 2: parseTeacherExcel の拡張
```javascript
async function parseTeacherExcel(input) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await input.arrayBuffer());

  let allTeachers = [];

  // 全シートを処理
  for (const sheet of workbook.worksheets) {
    console.log(`👨‍🏫 シート処理中: ${sheet.name}`);

    const yearMonth = detectYearMonthFromSheetName(sheet.name);
    const teachers = await parseSingleTeacherSheet(sheet, yearMonth);
    allTeachers = allTeachers.concat(teachers);
  }

  // 重複除去（同じdate+time+teacherId）
  const uniqueTeachers = deduplicateTeachers(allTeachers);

  return uniqueTeachers;
}
```

### Phase 3: UI拡張（オプション）
- シート選択機能（特定のシートのみを処理）
- 年月の手動指定機能
- プレビュー機能（読み込んだシート一覧を表示）

## テストケース

### TC001: 単一シートの後方互換性
- 1シートのみのファイルで正常動作
- 既存の動作と同じ結果

### TC002: 複数シート（連続した週）
- 3週間分（3シート）のファイルを読み込み
- 全てのスロットが統合される
- 日付順にソートされる

### TC003: シート名の検出
- 様々なシート名フォーマットで年月を正しく検出
- パターンマッチングが正常動作

### TC004: 重複データの処理
- 同じ日付のデータが2つのシートにある場合
- 適切に重複除去または警告

### TC005: 空シートのスキップ
- 空のシートが含まれていても正常動作
- エラーが発生しない

## リスク・制約

### リスク
1. **パフォーマンス**: 大量のシート（10+）で処理が遅延する可能性
   - 対策: プログレスバー表示、バックグラウンド処理

2. **メモリ使用量**: 全シートを同時にメモリに読み込むため、メモリ不足の可能性
   - 対策: シートごとに処理してデータを破棄、またはストリーム処理

3. **シート名の多様性**: 想定外のシート名フォーマット
   - 対策: 柔軟なパターンマッチング、フォールバック処理

### 制約
1. Excel形式のみサポート（.xlsx, .xls）
2. 最大シート数: 52週（1年分）を想定
3. 各シートのフォーマットは統一されている必要がある

## マイルストーン

- [x] 仕様書作成
- [ ] parseBoothExcel の拡張実装
- [ ] parseTeacherExcel の拡張実装
- [ ] シート名検出の強化
- [ ] テストケース作成
- [ ] 統合テスト
- [ ] ドキュメント更新
- [ ] リリース
