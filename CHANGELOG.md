# CHANGELOG

## Version 1.1.2.2 (2026年9月4日)

### カラー置換機能の改善
- **ColorReplacer.cs**: 表（Table）シェイプのセルの塗り色を色収集・一括置換の対象に追加
  - `ForEachTableCell` で走査する各セルの `Shape.Fill`（単色塗りのみ）を収集・置換対象に追加（これまで文字色のみ対応）
- **ColorReplaceLibrary.cs**: 置換色リストのエクスポート/インポート機能を追加
  - `ExportList(name, filePath)`: 指定リストを単体のJSONファイルとして書き出し
  - `ImportList(filePath, out overwritten)`: JSONファイルからリストを読み込み、同名リストは上書き
- **ColorReplaceDialog.cs**: 「保存済みリスト」欄に「エクスポート」「インポート」ボタンを追加し、`SaveFileDialog`/`OpenFileDialog` でファイルを介したリスト共有に対応

## Version 1.1.2.1 (2026年9月4日)

### カラー置換機能の改善
- **ColorReplacer.cs**: 表（Table）シェイプのセル内文字色を色収集・一括置換の対象に追加
  - `HasTextFrame` が `false` となる表シェイプについて `HasTable`/`Table.Cell(r,c).Shape.TextFrame` を走査するよう対応
  - 通常シェイプ・表セルで共通の色抽出/置換ロジックを `ExtractTextFrameColors`/`ApplyTextFrameReplacements` に共通化
- **ColorReplaceDialog.cs**: 置換色をHEXコードで直接入力できる欄を追加（選択行の置換色にリアルタイム反映）
- **ColorReplaceDialog.cs**: ダイアログ内ボタン・ラベルの中心線ズレを修正し、全ボタンの高さを28pxに統一

## Version 1.1.2.0 (2026年9月3日)

### 🎨 カラー置換機能の追加
- **ColorReplacer.cs**: 色収集・一括置換のコアロジック（新規作成）
  - `CollectUsedColors(ColorReplaceScope)`: 現在のスライド/プレゼンテーション全体を対象に、塗り（単色のみ）・線・フォント色（テキストRun単位）を収集
  - `ApplyReplacements(ColorReplaceScope, Dictionary<int,int>)`: 収集ロジックを再利用し、対象色を一括置換
  - グループ化図形は `GroupItems` を再帰的に走査
  - 全COM操作を `ComExceptionHandler.ExecuteComOperation(suppressErrors:true)` でラップし、1図形の失敗で処理全体を止めない
- **ColorReplaceLibrary.cs**: 置換色リストのJSON永続化（新規作成、`ShapeStyleLibrary.cs`と同一パターン）
  - 保存先: `%APPDATA%\MagosaAddIn\ColorReplaceLibrary.json`（最大50件）
- **ColorReplaceDialog.cs**: カラー置換ダイアログ（新規作成）
  - 対象範囲（現在のスライドのみ/プレゼンテーション全体）を選択して色を取得
  - ListViewで色（スウォッチ＋カラーコード）・使用数・置換後（スウォッチ＋カラーコード）を一覧表示
  - 標準の `ColorDialog`（`FullOpen = true`）による置換色選択
  - 置換色リストの名前付き保存・読込・削除
- **CustomRibbon**: 「カラー置換」ボタンを図形操作グループに追加

### 🔧 技術的変更
- `MagosaAddIn.csproj` に新規ファイル3件を追加
- `DataModels.cs` へ `ColorReplaceScope` 列挙型・`ColorUsageInfo`/`ColorReplacementEntry`/`ColorReplaceListEntry` クラスを追加

---

## Version 1.1.1.0 (2026年6月16日)

### 🎨 画像色編集機能の追加
- **ImageColorEditor.cs**: Magick.NET による画像処理パイプライン（新規作成）
  - 明るさ・コントラスト（BrightnessContrast）調整
  - 色相・彩度（Modulate）調整
  - グレースケール・セピア・白黒（Threshold）変換（排他制御）
  - カラーライズ：HSL カラースペースで H・S チャンネルを直接置換（L 保持）
  - ToneMode 指定時はカラーライズをスキップする排他ガード追加
  - アルファチャンネルを白背景にフラット化して処理
- **ImageColorSettings.cs**: 色編集パラメータ DTO（新規作成）
  - Brightness / Contrast / Hue / Saturation / ColorizeEnabled / ColorizeIntensity / ColorizeRgb / ToneMode / BlackWhiteThreshold
- **ImageColorEditDialog.cs**: 編集ダイアログ（新規作成）
  - GradientBar スライダー × 5（明るさ・コントラスト・色相・彩度・カラーライズ強度）
  - 色調変換ラジオボタン（なし/グレースケール/セピア/白黒）と閾値スライダー
  - カラーライズ チェックボックスと色選択ボタン（ColorPickerDialog連携）
  - 300ms デバウンスによる非同期リアルタイムプレビュー（変更前/後を横並び表示）
  - ToneMode 選択時にカラーライズを自動解除する UI 連動
- **GradientBar.cs**: 共通グラデーションスライダーコントロール（新規作成）
  - DrawGradient デリゲートで背景グラデーションをカスタマイズ可能
  - Value / Minimum / Maximum / ValueChanged イベント
- **ColorPickerDialog.cs**: HSL カラーピッカーダイアログ（新規作成）
  - 色相・彩度・輝度スライダー（GradientBar）+ HEX 入力

### 🎯 選択順序変更機能の追加
- **ShapeOrderManager.cs**: レイヤー適用・スタック保存ロジック（新規作成）
  - `ApplyLayerOrder(List<Shape>)`: 並び替え後の順序でレイヤー（Z 順）を適用
  - `SaveToStack(List<Shape>, ShapeStack)`: 並び替え後の図形リストをスタックに保存
- **SelectionOrderDialog.cs**: 図形並び替えダイアログ（新規作成）
  - ListView で選択図形を一覧表示（# / 図形名 / 種類 / Z 順）
  - ↑↓ ボタンで並び替え（ListViewItem.Tag に PowerPoint.Shape を保持）
  - 後続アクション選択: KeepSelection / ApplyLayer / SaveToStack
- **CustomRibbon**: 「選択順序変更」ボタンを選択補助グループに追加

### 🔧 技術的変更
- `MagosaAddIn.csproj` に新規ファイル5件を追加
- リボン「画像操作」グループに「色調編集」ボタンを追加

---

## Version 1.0.10 (2026年6月12日)

### 🖼️ 画像倍率同期機能の追加
- **ImageScaler.cs**: スケール係数計算・リサイズ適用のコアロジック（新規作成）
  - 実寸法÷ピクセル長による画像スケール（mm/px）の算出
  - クリップボードビットマップの解像度補正（スクリーン解像度・原寸解像度両対応）
  - スケール係数に基づく画像②のリサイズ実行・バリデーション
- **ImageScaleSyncDialog.cs**: 画像倍率同期ダイアログUI（新規作成）
  - 横並びレイアウト（1100×735px）で画像①・画像②を同時表示
  - プレビュー上のクリックによる測定区間の指定（起点→終点）
  - 座標の直接数値入力にも対応
  - 測定方向指定（自由・水平のみ・垂直のみ）
  - ズームボタンまたはマウスホイールで25%〜800%ズーム
  - ピクセル長・スケール係数・変換後サイズをリアルタイム計算・表示
  - 位置保持オプション（中心位置保持 / 左上位置保持）
- **DataModels.cs** へ `MeasurementMode` 列挙型・`ImageInfo` クラスを追加

### 🔧 技術的変更
- リボン「画像操作」グループを新設し「画像倍率同期」ボタンを追加
- `MagosaAddIn.csproj` に新規ファイル2件を追加

---

## Version 1.0.9 (2026年2月27日)

### ✏️ テキスト一括編集機能の追加
- **ShapeTextEditor.cs**: テキスト操作のコアロジック（新規作成）
  - テキスト情報の一括取得（`ShapeTextInfo`）
  - 個別テキスト適用・全図形一括設定・テキスト配布・一括削除
  - 検索・置換（大文字小文字区別オプション）
  - フォント書式一括変更（FontSettings クラス）
  - テキストレイアウト一括変更（TextLayoutSettings クラス：余白・行間）
- **TextBulkEditDialog.cs**: 4タブ構成のUI（新規作成）
  - **テキスト編集タブ**: DataGridViewで一覧表示・直接編集、テキスト配布（専用マルチラインTextBox、Enterで改行）
  - **検索・置換タブ**: 検索のみ・すべて置換
  - **フォント書式タブ**: チェックした項目のみ変更（全フォント対応・オートコンプリート付き）
  - **レイアウトタブ**: 行間・余白をチェックした項目のみ変更

### 📸 図形スタイルライブラリ機能の追加
- **ShapeStyleLibrary.cs**: スタイル管理のコアロジック（新規作成）
  - スタイル情報の抽出（塗り・グラデーション・枠線・影・フォント）
  - スタイルの保存・適用・削除・お気に入り管理
  - DataContractJsonSerializer によるJSON永続化
  - `%APPDATA%\MagosaAddIn\StyleLibrary.json` への自動保存・読み込み
  - JSONエクスポート・インポート（上書き/スキップ選択）
  - 最大100件保存
- **StyleLibraryDialog.cs**: ライブラリ管理UIダイアログ（新規作成）
  - 左ペイン: スタイル一覧（ListView）+ 検索フィルター + お気に入りフィルター
  - 右ペイン: GDI+によるリアルタイムプレビュー（塗り・グラデーション・枠線・影・テキスト色を描画）
  - スタイル保存・適用・削除・お気に入り切替・インポート/エクスポートボタン
  - ダブルクリックで即適用
- **StyleNameInputDialog**: スタイル名入力ダイアログ（StyleLibraryDialog内クラス）
  - 同名スタイル存在時の警告表示

### 🔧 技術的変更
- `System.Runtime.Serialization` アセンブリ参照を追加（DataContract/DataMember/DataContractJsonSerializer使用のため）
- `MagosaAddIn.csproj` に新規ファイル4件を追加
- リボン「図形操作」グループに「テキスト一括編集」「スタイルライブラリ」ボタンを追加

---

## Version 1.0.8 (2026年2月26日)

### 🎯 選択オブジェクトスタック機能の追加
- **ShapeStack.cs**: 複数の選択状態を保存・復元する機能（新規作成）
  - 最大20個のスタックを管理可能
  - 各スタックに図形参照を保存
  - 動的メニューで各スタックの状態を表示
  - 削除された図形は「無効」として自動表示
  - ShapeRangeを使用した複数図形の一括選択
- **リボンUI拡張**:
  - 選択補助グループに「スタックに追加」「スタック復元」「スタック数表示」を追加
  - 動的メニュー（Dynamic=true）でスタック一覧を自動更新

### 🐛 バグ修正
- 図形置き換えダイアログの継承設定
  - スタイル継承・テキスト継承のチェックボックス割り当てが逆になっていた問題を修正

---

## Version 1.0.7 (2026年2月20日)

### 新機能
- **📏 図形サイズ調整機能の追加**: ShapeResizer.cs（基準・比率保持・最大/最小・パーセント・固定サイズ）
- **🔁 配列複製機能の追加**: ShapeArrayer.cs（線形・円形・グリッド・パス・回転コピー）
- **🎨 テーマカラー生成機能の追加**: ThemeColorGenerator.cs（17種類の配色パターン）・ColorConverter.cs・ColorPaletteArranger.cs

---

## Version 1.0.6 (2026年2月18日)

### 改善・修正
- **🎨 UI/UX大幅改善**: BaseDialog.csの拡張によるダイアログ統一レイアウト実現
- **🐛 バグ修正**: ナンバリング機能の1個選択対応
