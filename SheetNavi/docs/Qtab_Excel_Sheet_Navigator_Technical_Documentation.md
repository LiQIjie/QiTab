# Qtab - Excel シートナビゲーター  
**プロジェクト技術文書 (詳細版)**

本書は **Qtab (Excel Sheet Navigator)** の技術仕様、設計思想、実装詳細、運用ガイドを網羅した開発者向けドキュメントです。

---

## 目次
1. [プロジェクト概要](#1-プロジェクト概要)
2. [技術スタック](#2-技術スタック)
3. [アーキテクチャ設計](#3-アーキテクチャ設計)
4. [主要コンポーネント](#4-主要コンポーネント)
5. [データフロー](#5-データフロー)
6. [主要機能フローチャート](#6-主要機能フローチャート)
7. [永続化仕様](#7-永続化仕様)
8. [イベント駆動設計](#8-イベント駆動設計)
9. [開発ガイドライン](#9-開発ガイドライン)
10. [トラブルシューティング](#10-トラブルシューティング)
11. [今後の拡張案](#11-今後の拡張案)
12. [ライセンス・著作権](#12-ライセンス著作権)
13. [変更履歴](#13-変更履歴)

---

## 1. プロジェクト概要

### 1.1 プロジェクト名
**Qtab - Excel Sheet Navigator**

### 1.2 目的
Excel のワークシート管理を効率化するための VSTO アドイン。多数のシートを扱うユーザーに対し、検索・グループ化・並べ替え・色設定・リネームを統合的に提供します。さらに、トルク・回転数・電力の相互計算機能と Excel 分析ツール（Solver / Data Analysis Tool）の統合管理を提供します。

### 1.3 種別
- **アドインタイプ**: Excel VSTO (Visual Studio Tools for Office) アドイン
- **UI フレームワーク**: Windows Forms (WinForms) ベースの CustomTaskPane
- **対象 Excel バージョン**: Excel 2016 以降（Office 2016, 2019, 2021, Microsoft 365）

### 1.4 開発環境
- **.NET Framework**: 4.7.2
- **C# バージョン**: 7.3
- **IDE**: Visual Studio 2019/2022
- **プロジェクトテンプレート**: Excel VSTO Add-in

### 1.5 主要機能一覧

| 機能 | 説明 | 実装状態 |
|------|------|----------|
| シート一覧表示 | タスクペインに全シートをリスト表示 | ✅ 完了 |
| 高速検索 | シート名の部分一致フィルタリング | ✅ 完了 |
| グループ化/解除 | 複数シートを論理グループに整理 | ✅ 完了 |
| インライン編集 | グループ名・シート名の直接編集 | ✅ 完了 |
| タブカラー設定 | 選択シートの色を一括変更 | ✅ 完了 |
| ドラッグ&ドロップ並べ替え | シート/グループの物理順序変更 | ✅ 完了 |
| ペインドック切替 | 左/右ドック位置の切替 | ✅ 完了 |
| CustomXML 永続化 | ブック単位でグループ状態を保存 | ✅ 完了 |
| 複数ウィンドウ対応 | ウィンドウ別にペイン管理 | ✅ 完了 |
| トルク・回転数・電力計算 | T/N/P 相互計算（自動計算） | ✅ 完了 |
| Excel ツール統合管理 | Solver/Data Analysis 統合切替 | ✅ 完了 |
| シート削除 | 選択シートの一括削除 | ✅ 完了 |

### 1.6 ユーザー対象
- 大量シート（10+ 枚）を扱う Excel ユーザー
- プロジェクト管理・データ分析担当者
- 財務・会計レポート作成者
- 機械設計・電力計算を行うエンジニア

---

## 2. 技術スタック

### 2.1 コア技術

| カテゴリ | 技術 | 用途 |
|---------|------|------|
| フレームワーク | .NET Framework 4.7.2 | ランタイム基盤 |
| 言語 | C# 7.3 | アプリケーションロジック |
| Excel 連携 | Microsoft.Office.Interop.Excel | Excel オブジェクトモデル操作 |
| VSTO | Microsoft.Office.Tools | CustomTaskPane, リボン |
| UI | Windows Forms (System.Windows.Forms) | カスタムコントロール |
| 描画 | System.Drawing | グラフィックス・色変換 |
| XML | System.Xml.Linq (LINQ to XML) | 永続化データ操作 |
| 右クリックメニュー | Office.Core.CommandBars | Ply メニュー拡張 |

### 2.2 プロジェクト構成

```
SheetNavi/
├── Qtab.csproj              // VSTO プロジェクトファイル
├── ThisAddIn.cs             // アドインエントリポイント
├── ThisAddIn.Designer.cs    // 自動生成コード
├── ThisAddIn.Designer.xml   // VSTO マニフェスト設定
├── QTabNavi.cs              // メインUIコントロール（計算パネル含む）
├── QtabList.cs              // シート一覧コントロール
├── Properties/              // アセンブリ情報
├── docs/                    // ドキュメント
└── README.md                // プロジェクト概要
```

### 2.3 依存パッケージ
- **Microsoft.Office.Interop.Excel** (15.0+): Excel オブジェクトモデル
- **Microsoft.Office.Tools** (10.0+): VSTO ツールキット
- **Office.Core** (15.0+): CommandBars API

---

## 3. アーキテクチャ設計

### 3.1 レイヤー構成

```
┌─────────────────────────────────────────────────────────┐
│                   UI Layer (Presentation)               │
│  ┌──────────────────┐      ┌──────────────────┐         │
│  │   QTabNavi       │◄─────┤   QtabList       │         │
│  │  (Main Panel)    │      │  (Sheet List)    │         │
│  │  + 計算パネル     │      │                  │         │
│  │  + ツール管理     │      │                  │         │
│  └────────┬─────────┘      └──────────────────┘         │
└───────────┼─────────────────────────────────────────────┘
            │ Events (14 types)
┌───────────▼──────────────────────────────────────────────┐
│            Application Layer (Business Logic)            │
│  ┌──────────────────────────────────────────────────┐    │
│  │                 ThisAddIn                        │    │
│  │  • Excel イベント購読 (8 types)                   │    │
│  │  • タスクペイン生成・管理                          │    │
│  │  • Excel 実体操作 (Move/Rename/Color)             │    │
│  │  • 状態管理 (_groupsByWorkbook)                   │    │
│  │  • Excel AddIn 管理 (Solver/Analysis Tool)   │    │
│  └──────────────────────────────────────────────────┘    │
└───────────┬──────────────────────────────────────────────┘
            │ Interop API / CustomXML
┌───────────▼──────────────────────────────────────────────┐
│              Persistence / Excel Layer                   │
│  ┌──────────────────┐      ┌──────────────────┐          │
│  │ CustomXMLParts   │      │ Excel.Worksheet  │          │
│  │ (SheetNaviGroups)│      │ Excel.Workbook   │          │
│  │                  │      │ Excel.AddIns     │          │
│  └──────────────────┘      └──────────────────┘          │
└──────────────────────────────────────────────────────────┘
```

### 3.2 責務分離原則

| レイヤー | クラス | 責務 | Excel 操作 |
|---------|--------|------|------------|
| UI Layer | `QTabNavi` | 検索・ドック・計算UI・ツールUI・イベント転送 | ❌ 不可 |
| UI Layer | `QtabList` | 描画・選択・D&D・インライン編集 | ❌ 不可 |
| Application | `ThisAddIn` | Excel操作・永続化・状態管理・AddIn管理 | ✅ 許可 |

**設計原則**:
- UI は **イベント発火のみ**（Excel への直接アクセス禁止）
- `ThisAddIn` が **Excel Interop の唯一のエントリポイント**
- 例外は UI の継続性を優先（ペイン可視化を維持）

### 3.3 ウィンドウ別ペイン管理

```csharp
// ThisAddIn.cs (抜粋)
private Dictionary<string, CustomTaskPane> _panesByWindow;
private Dictionary<string, QTabNavi> _controlsByWindow;

private string GetActiveWindowKey() 
{
    return this.Application?.ActiveWindow?.Caption ?? string.Empty;
}

private CustomTaskPane GetPaneForActiveWindow()
{
    var key = GetActiveWindowKey();
    if (_panesByWindow.TryGetValue(key, out var pane)) return pane;
    
    // フォールバック: Window.Caption でマッチング
    foreach (CustomTaskPane p in this.CustomTaskPanes)
    {
        var wcap = (p.Window as Excel.Window)?.Caption;
        if (wcap == key)
        {
            _panesByWindow[key] = p;
            return p;
        }
    }
    return null;
}
```

**特徴**:
- ウィンドウごとに独立したペインと QTabNavi インスタンスを生成
- キャッシュでパフォーマンス向上
- Excel の複数ウィンドウ機能に完全対応

---

## 4. 主要コンポーネント

### 4.1 ThisAddIn.cs (アプリケーション層)

#### 4.1.1 責務
- VSTO エントリポイント (`ThisAddIn_Startup`, `ThisAddIn_Shutdown`)
- Excel イベント購読（8種類）
- タスクペイン生成・キャッシュ管理（ウィンドウ別）
- Excel 実体操作（Activate, Move, Rename, Tab.Color, Delete）
- グループ状態の永続化（CustomXMLParts）
- ドラッグ&ドロップ処理
- Excel AddIn 管理（Solver, Analysis Tool の統合切替）

#### 4.1.2 主要メソッド

| メソッド名 | 戻り値 | 説明 |
|-----------|--------|------|
| `ThisAddIn_Startup` | void | イベント購読・UI初期化 |
| `EnsurePaneForActiveWindow` | CustomTaskPane | ペイン生成/取得（ウィンドウ別） |
| `RefreshSheetList` | void | シート一覧更新 |
| `GroupSheetsForActiveWorkbook` | void | グループ化実行 |
| `HandleSheetDragDrop` | void | シートD&D処理 |
| `HandleGroupDragDrop` | void | グループD&D処理 |
| `SaveGroupsForWorkbook` | void | XML保存 |
| `LoadGroupsFromWorkbook` | Dictionary | XML読込 |
| `DeleteSheetsForActiveWorkbook` | void | シート削除 |
| `ToggleBothTools` | void | Solver/Analysis 統合切替 |
| `UpdateToolButtonsStatus` | void | ツールボタン状態更新 |
| `GetToolsStatus` | (bool, bool) | 現在のツール状態取得 |

#### 4.1.3 状態管理

```csharp
// グループ情報（ワークブック別）
private Dictionary<string, Dictionary<string, (List<string> Members, Color Color)>> 
    _groupsByWorkbook;

// ウィンドウ別ペインキャッシュ
private Dictionary<string, CustomTaskPane> _panesByWindow;

// ウィンドウ別コントロールキャッシュ
private Dictionary<string, QTabNavi> _controlsByWindow;

// シート数・順序キャッシュ（変更検出用）
private Dictionary<string, int> _sheetCountsByWorkbook;
private Dictionary<string, List<string>> _sheetOrderByWorkbook;

// 防抖タイマー（200ms）
private Timer _refreshTimer;

// 軽量ポーリングタイマー（500ms）- D&D等の未検出イベント対応
private Timer _pollTimer;
```

#### 4.1.4 Excel AddIn 管理

```csharp
// 統合切替（両方同時にON/OFF)
private void ToggleBothTools()
{
    var (solverEnabled, analysisEnabled) = GetToolsStatus();
    bool enableBoth = !(solverEnabled && analysisEnabled);
    
    var solver = FindAddInByKeywords(new[] { "solver", "ソルバー", "求解" });
    var atp = FindAddInByKeywords(new[] { "analysis Tool", "データ分析", "分析ツール" });
    
    if (solver != null) solver.Installed = enableBoth;
    if (atp != null) atp.Installed = enableBoth;
    
    UpdateToolButtonsStatus();
    MessageBox.Show(enableBoth ? "Solver と 数据分析 已同时启用" : "Solver と 数据分析 已同时关闭");
}

// 多言語対応キーワード検索
private Excel.AddIn FindAddInByKeywords(string[] keywords)
{
    foreach (Excel.AddIn ai in Application.AddIns)
    {
        if (MatchesAddIn(ai, keywords)) return ai;
    }
    // AddIns2 も確認
    try
    {
        foreach (Excel.AddIn ai in (Application.AddIns2 as Excel.AddIns2))
        {
            if (MatchesAddIn(ai, keywords)) return ai;
        }
    }
    catch { }
    return null;
}
```

### 4.2 QTabNavi.cs (UI層 - メインパネル)

#### 4.2.1 責務
- 検索入力ボックス（TextBox）
- ドック切替ボタン（左右）
- 計算パネル（トルク・回転数・電力の相互計算）
- Excel ツール統合管理ボタン
- 統計情報表示（Total / Hidden シート数）
- `QtabList` イベントの転送

#### 4.2.2 UI レイアウト

```
┌─────────────────────────────────────────────────┐
│  [🔍 Search sheets...]                          │ ← TextBox (DockStyle.Top)
├─────────────────────────────────────────────────┤
│  ┌─ Group 1 ───────────────┐                    │
│  │  ├─ Sheet A             │                    │
│  │  └─ Sheet B             │                    │ ← QtabList (DockStyle.Fill)
│  └─────────────────────────┘                    │
│  Sheet C                                        │
├─────────────────────────────────────────────────┤
│  [○] 5 N·m [●] 523.599 W [SolverData:ON]       │ ← 計算パネル
│  [○] 1000 rpm                                   │
│  Total: 3 | Hidden: 0                     [→]  │ ← ステータス + ドック切替
└─────────────────────────────────────────────────┘
```

#### 4.2.3 計算パネル仕様

**物理式**: `P = T × 2π × N / 60` (SI単位系)

| 入力 | 単位 | 対応表記 |
|-----|------|---------|
| トルク (T) | N·m | `5`, `5 N·m`, `5000 mNm` |
| 回転数 (N) | rpm | `1000`, `1 krpm`, `1000 r/min` |
| 電力 (P) | W | `200`, `0.2 kW` |

**動作仕様**:
- 初期値: `N=1000 rpm`, `T=5 N·m`, 計算対象: `P`（自動算出 ≈ 523.6W）
- ラジオボタンで計算対象を選択（T/N/P）
- 入力値変更時に自動計算（被計算フィールド変更時は再計算なし）
- 入力が空白または不正な場合は `Err.` を表示

```csharp
// 自動計算ロジック（抜粋）
private void AutoCalculateIfReady(string changedField)
{
    if (_suppressAutoCalc) return;
    
    // 被計算対象が変更された場合は計算しない
    if (changedField == "T" && _rbT.Checked) return;
    if (changedField == "N" && _rbN.Checked) return;
    if (changedField == "P" && _rbP.Checked) return;
    
    double? tNm = ParseNumber(_txtT.Text);
    double? nRpm = ParseNumber(_txtN.Text);
    double? pW = ParsePowerW(_txtP.Text);
    
    if (_rbP.Checked)
    {
        if (tNm == null || nRpm == null)
            _txtP.Text = "Err.";
        else
            _txtP.Text = Math.Round((tNm.Value * 2.0 * Math.PI * nRpm.Value) / 60.0, 6).ToString();
    }
    // ... N, T の計算も同様
}
```

#### 4.2.4 ツール管理ボタン

**状態表示**:
- **緑色**: `"SolverData:ON"` - 両方有効
- **黄色**: `"Solver:OFF"` または `"Data:OFF"` - 一方のみ無効
- **赤色**: `"SolverData:OFF"` - 両方無効

```csharp
public void SetToolStatus(bool solverEnabled, bool analysisEnabled)
{
    if (solverEnabled && analysisEnabled)
    {
        _btnTools.Text = "SolverData:ON";
        _btnTools.BackColor = Color.LightGreen;
    }
    else if (!solverEnabled && !analysisEnabled)
    {
        _btnTools.Text = "SolverData:OFF";
        _btnTools.BackColor = Color.LightCoral;
    }
    else
    {
        _btnTools.Text = solverEnabled ? "Data:OFF" : "Solver:OFF";
        _btnTools.BackColor = Color.Khaki;
    }
}
```

#### 4.2.5 イベント転送

```csharp
public class QTabNavi : UserControl
{
    // 14種類のイベント
    public event Action<string> OnSheetSelected;
    public event Action<List<string>> GroupRequested;
    public event Action<List<string>> DeleteSheetRequested;
    public event Action ToolsToggleRequested; // 新規: ツール統合切替
    // ... 他10種
    
    public QTabNavi()
    {
        // QtabList → ThisAddIn への橋渡し
        _list.OnSheetSelected += s => OnSheetSelected?.Invoke(s);
        _list.GroupRequested += names => GroupRequested?.Invoke(names);
        _list.DeleteSheetRequested += names => DeleteSheetRequested?.Invoke(names);
        
        // ツールボタン → ThisAddIn
        _btnTools.Click += (s, e) => ToolsToggleRequested?.Invoke();
    }
}
```

### 4.3 QtabList.cs (UI層 - リスト表示)

#### 4.3.1 責務
- カスタム描画（`OnPaint` オーバーライド）
- マウス操作処理（Click, DoubleClick, MouseMove）
- 右クリックメニュー（ContextMenuStrip - 6項目）
- インライン編集（TextBox オーバーレイ）
- ドラッグ&ドロップ（DoDragDrop, DragOver, DragDrop）
- Excel 右クリックメニュー表示リクエスト

#### 4.3.2 右クリックメニュー項目

| 項目 | ショートカット | 機能 |
|-----|--------------|------|
| Group | - | 選択シートをグループ化 |
| Ungroup | - | グループ解除 |
| Rename | F2 | シート/グループ名変更 |
| Delete | Del | シート削除 |
| Set Tab Color... | - | 色選択ダイアログ |
| Excel Context Menu | - | Excel 標準メニュー（統合） |

#### 4.3.3 描画ロジック

```csharp
protected override void OnPaint(PaintEventArgs e)
{
    var g = e.Graphics;
    g.Clear(BackColor);
    int y = 4;
    
    // Excel の物理順序で描画
    foreach (var sheetName in _allSheets)
    {
        if (sheetToGroup.TryGetValue(sheetName, out var gm))
        {
            if (!drawnGroups.Contains(gm.Name))
            {
                DrawGroupHeader(g, gm, rect, arrowSize); // グループヘッダー
                if (!gm.Collapsed)
                {
                    DrawSheetItem(g, member, rect, color); // メンバーシート
                }
            }
        }
        else
        {
            DrawSheetItem(g, sheetName, rect, Color.Empty); // 未グループ
        }
    }
    
    // ドラッグインジケータ（青線）
    if (_dragOverHit != null) DrawDropIndicator(g, _dragOverHit);
}
```

#### 4.3.4 D&D フロー

```
[MouseDown] → _dragStartHit 記録
     ↓
[MouseMove] → 閾値5px超過 → DoDragDrop("SHEET:A")
     ↓
[DragOver] → HitTest(cursor) → _dragOverHit 更新 → Invalidate (青線表示)
     ↓
[DragDrop] → SheetDragDropRequested(dragged, target, insertBefore)
     ↓
[Finally] → _dragOverHit = null → Invalidate
```

---

## 5. データフロー

### 5.1 起動時フロー

```
[Excel 起動]
    ↓
ThisAddIn_Startup
    ├─ CreateControlAndPane
    │   ├─ new QTabNavi() → イベント購読（14種）
    │   └─ EnsurePaneForActiveWindow
    │       └─ CustomTaskPanes.Add(_control, "📑 Qtab", ActiveWindow)
    ├─ Excel イベント購読 (WorkbookActivate, SheetActivate 他8種)
    ├─ 防抖タイマー初期化 (_refreshTimer: 200ms)
    ├─ ポーリングタイマー初期化 (_pollTimer: 500ms)
    └─ TryAddPlyContextItems (右クリックメニュー拡張)
    ↓
RefreshSheetList
    ├─ EnsurePaneExists
    ├─ LoadSheets(names) → QTabNavi → QtabList
    ├─ ApplyGroups(groups) → QtabList
    └─ UpdateToolButtonsStatus (Solver/Analysis 状態取得)
    ↓
ShowPane (Width=210, Visible=true)
```

### 5.2 グループ化フロー

```
[右クリック → Group]
    ↓
QtabList.GroupRequested(names)
    ↓
QTabNavi.GroupRequested(names)
    ↓
ThisAddIn.GroupSheetsForActiveWorkbook(names)
    ├─ 既存グループから除外
    ├─ 新規グループ作成 ("Group N")
    ├─ Excel で連続配置 (Move)
    ├─ Tab.Color = LightBlue
    ├─ SaveGroupsForWorkbook (CustomXML)
    └─ ApplyGroups(groups) → QtabList
    ↓
Invalidate → OnPaint → 描画更新
```

### 5.3 ツール切替フロー

```
[ツールボタンクリック]
    ↓
QTabNavi.ToolsToggleRequested
    ↓
ThisAddIn.ToggleBothTools
    ├─ GetToolsStatus() → (solverOn, analysisOn)
    ├─ enableBoth = !(両方ON)
    ├─ FindAddInByKeywords(["solver", "ソルバー", "求解"])
    ├─ FindAddInByKeywords(["analysis Tool", "データ分析", "分析ツール"])
    ├─ solver.Installed = enableBoth
    ├─ atp.Installed = enableBoth
    ├─ atpVba.Installed = enableBoth (VBA版も同期)
    ├─ UpdateToolButtonsStatus()
    └─ MessageBox 表示
```

### 5.4 計算パネルフロー

```
[初期化時]
    ↓
InitializeDefaultCalculationState
    ├─ _txtN.Text = "1000" (rpm)
    ├─ _txtT.Text = "5" (N·m)
    ├─ _rbP.Checked = true (電力を計算対象)
    └─ AutoCalculateIfReady(null) → _txtP.Text ≈ "523.599"
    ↓
[ユーザー入力: _txtT → "10"]
    ↓
_txtT.TextChanged
    ↓
AutoCalculateIfReady("T")
    ├─ 被計算対象チェック (_rbT.Checked? → false)
    ├─ ParseNumber("10") → 10.0
    ├─ ParseNumber("1000") → 1000.0
    ├─ P = 10 × 2π × 1000 / 60 ≈ 1047.2
    └─ _txtP.Text = "1047.197551" (自動更新)
```

### 5.5 ドラッグ&ドロップフロー

```
[マウス押下: Sheet A]
    ↓
OnMouseDown → _dragStartHit = HitInfo(Sheet A)
    ↓
[マウス移動: 5px+]
    ↓
OnMouseMove → DoDragDrop("SHEET:A", DragDropEffects.Move)
    ↓
[ドラッグ中: Sheet B 上]
    ↓
OnDragOver
    ├─ HitTest(cursor) → targetHit = HitInfo(Sheet B)
    ├─ _dragOverHit = targetHit
    └─ Invalidate → DrawDropIndicator (青線)
    ↓
[マウス離す: Sheet B 上]
    ↓
OnDragDrop
    ├─ insertBefore = (Y < 中央)
    ├─ SheetDragDropRequested(A, B, before)
    └─ ThisAddIn.HandleSheetDragDrop
        ├─ wsA.Move(wsB, Type.Missing) または Move(Type.Missing, wsB)
        ├─ グループ辞書更新（同グループ内/跨グループ処理）
        ├─ SaveGroupsForWorkbook
        └─ RefreshSheetList
```

### 5.6 インライン編集フロー

```
[ダブルクリック: Sheet A or Group 1]
    ↓
OnMouseDoubleClick
    ├─ HitType == GroupHeader → BeginInlineRenameGroup(name)
    └─ HitType == Sheet → BeginInlineRenameSheet(name)
    ↓
BeginEdit(rect, original, isGroup)
    ├─ _editBox.SetBounds(rect) → TextBox 表示
    ├─ _editBox.Text = original
    └─ _editBox.Focus()
    ↓
[Enter 押下]
    ↓
CommitEdit(true)
    ├─ newName = _editBox.Text
    ├─ isGroup ? GroupRenameCommitted(old, new) : RenameSheetCommitted(old, new)
    └─ ThisAddIn イベントハンドラ
        ├─ ws.Name = newName (Excelリネーム)
        ├─ グループ辞書更新
        ├─ SaveGroupsForWorkbook
        └─ RefreshSheetList
```

---

## 6. 主要機能フローチャート

### 6.1 起動・初期化

```
[Excel 起動]
    ↓
ThisAddIn_Startup
    ↓
CreateControlAndPane
    ├─ new QTabNavi()
    ├─ イベント購読 (14種)
    └─ EnsurePaneForActiveWindow
        └─ CustomTaskPanes.Add
    ↓
Excel イベント購読
    ├─ WorkbookActivate
    ├─ WindowActivate
    ├─ WorkbookNewSheet
    ├─ SheetActivate
    └─ SheetBeforeDelete
    ↓
タイマー初期化
    ├─ _refreshTimer (200ms 防抖)
    └─ _pollTimer (500ms 順序変更検出)
    ↓
RefreshSheetList
    ├─ LoadSheets → QtabList
    ├─ ApplyGroups → QtabList
    └─ UpdateToolButtonsStatus
    ↓
ShowPane (表示)
```

### 6.2 シート選択・アクティブ化

```
[ユーザーがシートをクリック]
    ↓
QtabList.OnMouseUp
    ├─ Ctrl? → 複数選択
    ├─ Shift? → 範囲選択
    └─ なし → 単一選択
    ↓
OnSheetSelected(name)
    ↓
QTabNavi → ThisAddIn.ActivateSheet(name)
    ↓
wb.Worksheets[name].Activate()
    ↓
[SheetActivate イベント発火]
    ↓
Highlight(name) → QtabList
```

### 6.3 グループ化処理

```
[選択シートを右クリック → Group]
    ↓
GroupRequested(names)
    ↓
GroupSheetsForActiveWorkbook
    ├─ 既存グループから除外
    ├─ 新規グループ ("Group N") 作成
    ├─ 辞書に追加
    └─ Excel 操作
        ├─ 連続配置 (Move)
        └─ Tab.Color = LightBlue
    ↓
SaveGroupsForWorkbook (CustomXML)
    ↓
ApplyGroups → QtabList
    ↓
Invalidate → 描画更新
```

### 6.4 ドラッグ&ドロップ並べ替え

```
[マウス押下: シートA]
    ↓
_dragStartHit = A
    ↓
[マウス移動: 5px超]
    ↓
DoDragDrop("SHEET:A")
    ↓
[ドラッグ中]
OnDragOver
    ├─ HitTest → _dragOverHit = B
    └─ Invalidate (青線表示)
    ↓
[マウス離す: シートB上]
    ↓
OnDragDrop
    ├─ insertBefore = (Y < 中央)
    ├─ SheetDragDropRequested(A, B, before)
    └─ HandleSheetDragDrop
        ├─ Excel: wsA.Move(wsB, ...)
        ├─ 辞書更新（同グループ内/跨グループ）
        ├─ Save
        └─ RefreshSheetList
```

### 6.5 インライン編集

```
[ダブルクリック: シート or グループ]
    ↓
BeginInlineRename*
    ├─ TextBox 表示 (オーバーレイ)
    └─ _editBox.Focus()
    ↓
[Enter]
    ↓
CommitEdit
    ├─ *RenameCommitted イベント
    └─ ThisAddIn ハンドラ
        ├─ Excel リネーム
        ├─ 辞書更新
        ├─ Save
        └─ Refresh
```

### 6.6 ツール統合切替

```
[ツールボタンクリック]
    ↓
ToolsToggleRequested
    ↓
ToggleBothTools
    ├─ 現在状態確認 GetToolsStatus()
    ├─ 切替判定 (両方ON → OFF, それ以外 → ON)
    ├─ FindAddInByKeywords (多言語対応)
    ├─ Solver.Installed = enableBoth
    ├─ AnalysisToolPak.Installed = enableBoth
    ├─ AnalysisToolPak-VBA.Installed = enableBoth
    ├─ UpdateToolButtonsStatus()
    └─ MessageBox 表示
```

---

## 7. 永続化仕様

### 7.1 保存形式: CustomXMLParts

#### 7.1.1 XML スキーマ定義

```xml
<?xml version="1.0" encoding="utf-8"?>
<SheetNaviGroups>
  <Group name="グループ名" color="#RRGGBB">
    <Sheet>シート名1</Sheet>
    <Sheet>シート名2</Sheet>
    ...
  </Group>
  <Group name="別グループ" color="#ADD8E6">
    <Sheet>シート名3</Sheet>
  </Group>
</SheetNaviGroups>
```

#### 7.1.2 実装例

```csharp
// 保存 (SaveGroupsForWorkbook)
var root = new XElement("SheetNaviGroups",
    groups.Select(kv => new XElement("Group",
        new XAttribute("name", kv.Key),
        kv.Value.Item2 != Color.Empty ? 
            new XAttribute("color", ColorTranslator.ToHtml(kv.Value.Item2)) : null,
        kv.Value.Item1.Select(s => new XElement("Sheet", s))
    ))
);

// 既存XMLパーツ削除
for (int i = wb.CustomXMLParts.Count; i >= 1; i--)
{
    var p = wb.CustomXMLParts[i];
    if (p.XML.IndexOf("<SheetNaviGroups", StringComparison.OrdinalIgnoreCase) >= 0)
        p.Delete();
}

wb.CustomXMLParts.Add(root.ToString());

// 読込 (LoadGroupsFromWorkbook)
var doc = XElement.Parse(part.XML);
var dict = new Dictionary<string, (List<string>, Color)>();
foreach (var g in doc.Elements("Group"))
{
    var gname = (string)g.Attribute("name");
    var colorAttr = (string)g.Attribute("color");
    Color c = !string.IsNullOrEmpty(colorAttr) ? 
        ColorTranslator.FromHtml(colorAttr) : Color.Empty;
    var sheets = g.Elements("Sheet").Select(x => (string)x).ToList();
    dict[gname] = (sheets, c);
    
    // 色復元（Excel タブに適用）
    if (c != Color.Empty)
    {
        foreach (var name in sheets)
        {
            var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == name);
            if (ws != null) ws.Tab.Color = ColorTranslator.ToOle(c);
        }
    }
}
```

### 7.2 保存タイミング

| 操作 | トリガーメソッド | タイミング |
|------|----------------|-----------|
| グループ化 | `GroupSheetsForActiveWorkbook` | 直後 |
| グループ解除 | `UngroupSheetsForActiveWorkbook` | 直後 |
| グループリネーム | `RenameGroupKeyForActiveWorkbook` | 確定時 |
| シートリネーム | `RenameSheetCommitted` ハンドラ | 確定時 |
| タブカラー設定 | `ColorTabsForActiveWorkbook` | 適用後 |
| D&D完了 | `HandleSheetDragDrop`, `HandleGroupDragDrop` | 移動後 |
| シート削除 | `DeleteSheetsForActiveWorkbook` | 削除後 |

### 7.3 色復元ロジック

```csharp
// LoadGroupsFromWorkbook 内
foreach (var kv in dict)
{
    var members = kv.Value.Item1;
    var color = kv.Value.Item2;
    if (color == Color.Empty) continue;
    
    var wbSheets = wb.Worksheets.Cast<Excel.Worksheet>().ToList();
    foreach (var name in members)
    {
        var ws = wbSheets.FirstOrDefault(x => x.Name == name);
        if (ws != null)
        {
            ws.Tab.Color = ColorTranslator.ToOle(color); // 色復元
        }
    }
}
```

---

## 8. イベント駆動設計

### 8.1 イベントフロー階層

```
QtabList (UI Events - 13種)
    ↓ (イベント発火)
QTabNavi (Relay + 1種追加)
    ↓ (イベント転送)
ThisAddIn (Business Logic)
    ↓ (Excel 操作)
Excel Workbook/Worksheet/AddIns
```

### 8.2 UI イベント一覧

| イベント名 | パラメータ | 発火タイミング | 処理内容 |
|-----------|-----------|--------------|----------|
| `OnSheetSelected` | (string name) | シートクリック | Activate |
| `GroupRequested` | (List<string> names) | 右クリック→Group | グループ化 |
| `UngroupRequested` | (List<string> names) | 右クリック→Ungroup | グループ解除 |
| `RenameGroupRequested` | (List<string> names) | 右クリック→Rename | リネーム開始 |
| `ColorRequested` | (List<string>, Color) | 右クリック→Color | 色設定 |
| `RenameSheetRequested` | (string name) | ダブルクリック | リネーム開始 |
| `GroupRenameCommitted` | (string old, string new) | Enter押下 | グループ改名確定 |
| `RenameSheetCommitted` | (string old, string new) | Enter押下 | シート改名確定 |
| `DeleteSheetRequested` | (List<string> names) | 右クリック→Delete | シート削除 |
| `ShowExcelContextMenu` | (string sheet, Point pos) | 右クリック→Excel Menu | Excel 標準メニュー |
| `SheetDragDropRequested` | (string dragged, string target, bool before) | D&Dドロップ | シート並替 |
| `GroupDragDropRequested` | (string group, string target, bool before) | D&Dドロップ | グループ並替 |
| `DockLeftRequested` | () | ボタンクリック | 左ドック |
| `DockRightRequested` | () | ボタンクリック | 右ドック |
| `ToolsToggleRequested` | () | ツールボタンクリック | Solver/Analysis 切替 |

### 8.3 Excel イベント購読

```csharp
// ThisAddIn_Startup 内
var app = this.Application;

app.WorkbookActivate += wb => { 
    PostShowPane(); 
    UpdateToolButtonsStatus(); 
};

app.SheetActivate += sh => {
    if (sh is Excel.Worksheet ws)
    {
        _control?.Highlight(ws.Name);
        RefreshSheetList(); // 色同期
    }
};

app.WorkbookNewSheet += (wb, sh) => {
    RefreshSheetList(); // 即座に反映
    _refreshTimer.Start(); // 防抖
    EnsurePaneExists();
    PostShowPane();
    UpdateToolButtonsStatus();
};

app.SheetBeforeDelete += (object sh) => {
    // グループ情報からシートを除外
    var ws = sh as Excel.Worksheet;
    var wb = ws?.Parent as Excel.Workbook;
    var name = ws?.Name;
    if (wb != null && !string.IsNullOrEmpty(name))
    {
        var groups = GetGroupsForWorkbook(wb, false);
        if (groups != null)
        {
            foreach (var kv in groups.ToList())
            {
                kv.Value.Item1.Remove(name);
                if (kv.Value.Item1.Count == 0) groups.Remove(kv.Key);
            }
            SaveGroupsForWorkbook(wb);
            GetControlForActiveWindow()?.ApplyGroups(groups);
        }
    }
    _refreshTimer.Start();
};

((Excel.AppEvents_Event)app).WindowActivate += (wb, wn) => {
    EnsurePaneExists();
    _pollTimer?.Start(); // ポーリング開始
    PostShowPane();
    UpdateToolButtonsStatus();
};

app.SheetBeforeRightClick += (sh, target, ref bool cancel) => {
    _refreshTimer.Start();
    TryAddPlyContextItems();
};
```

### 8.4 防抖処理 (Debounce)

```csharp
// 200ms 後に一度だけ RefreshSheetList を実行
_refreshTimer = new Timer();
_refreshTimer.Interval = 200;
_refreshTimer.Tick += (s, e) => {
    _refreshTimer.Stop();
    RefreshSheetList();
};

// 頻繁なイベントで再スタート
_refreshTimer.Stop();
_refreshTimer.Start();
```

### 8.5 軽量ポーリング (Polling)

```csharp
// 500ms ごとにシート数・順序変化を検出
_pollTimer = new Timer();
_pollTimer.Interval = 500;
_pollTimer.Tick += (s, e) => {
    var wb = this.Application.ActiveWorkbook;
    if (wb == null) return;
    
    var key = GetWorkbookKey(wb);
    var visibleSheets = wb.Worksheets.Cast<Excel.Worksheet>()
        .Where(w => w.Visible == Excel.XlSheetVisibility.xlSheetVisible)
        .OrderBy(w => w.Index)
        .Select(w => w.Name)
        .ToList();
    
    bool needRefresh = false;
    
    // 数量変化検出
    if (_sheetCountsByWorkbook.TryGetValue(key, out var lastCount) 
        && lastCount != visibleSheets.Count)
        needRefresh = true;
    
    // 順序変化検出
    if (_sheetOrderByWorkbook.TryGetValue(key, out var lastOrder)
        && !lastOrder.SequenceEqual(visibleSheets))
        needRefresh = true;
    
    if (needRefresh)
    {
        _sheetCountsByWorkbook[key] = visibleSheets.Count;
        _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
        RefreshSheetList();
    }
};
```

---

## 9. 開発ガイドライン

### 9.1 コーディング規約

#### 9.1.1 命名規則

| 要素 | 規則 | 例 |
|-----|------|-----|
| クラス | PascalCase | `QTabNavi` |
| メソッド | PascalCase + 動詞 | `RefreshSheetList()` |
| プライベートフィールド | _camelCase | `_refreshTimer` |
| イベント | PascalCase + 過去分詞/名詞 | `GroupRequested` |
| ローカル変数 | camelCase | `sheetName` |

#### 9.1.2 例外処理パターン

```csharp
// 原則: try-catch で囲み、Log のみ（UI継続）
try
{
    var ws = wb.Worksheets[name];
    ws.Activate();
}
catch (Exception ex)
{
    Log("Activate error: {0}", ex.Message);
    // UI は継続（ペイン可視化維持）
}
```

### 9.2 描画最適化

```csharp
public QtabList()
{
    DoubleBuffered = true; // ちらつき防止
    BackColor = SystemColors.Window;
}
```

### 9.3 テスト観点

#### 9.3.1 単体テスト項目
- [ ] グループ化: 連続配置確認
- [ ] D&D: before/after 判定精度
- [ ] 永続化: CustomXML 読み書き整合性
- [ ] リネーム: 重複名処理（サフィックス追加）
- [ ] 色設定: OLE カラー変換正確性
- [ ] 計算パネル: T/N/P 相互計算精度（誤差 < 0.001%）
- [ ] 単位変換: `mNm`, `krpm`, `kW` 正確性
- [ ] ツール切替: Solver/Analysis 状態同期
- [ ] シート削除: グループ情報更新確認

#### 9.3.2 統合テスト項目
- [ ] 複数ワークブック切替
- [ ] 複数ウィンドウ同時操作（独立ペイン）
- [ ] 大量シート（100+）での性能
- [ ] Excel バージョン互換性（2016, 2019, 2021, 365）
- [ ] 多言語環境での AddIn 検索（日本語/中国語/英語）
- [ ] ポーリングによるD&D検出精度

### 9.4 デバッグTips

#### 9.4.1 ログ有効化

```csharp
// Debug ビルドでログ有効化
#if DEBUG
private void Log(string fmt, params object[] args) 
{ 
    System.Diagnostics.Debug.WriteLine(fmt, args); 
}
#else
private void Log(string fmt, params object[] args) { }
#endif
```

#### 9.4.2 推奨ブレークポイント
- `RefreshSheetList`: UI 更新タイミング確認
- `HandleSheetDragDrop`: D&D ロジック検証
- `SaveGroupsForWorkbook`: XML 生成内容確認
- `OnDragDrop`: ドロップイベント引数確認
- `AutoCalculateIfReady`: 計算トリガー確認
- `ToggleBothTools`: AddIn 切替ロジック確認

---

## 10. トラブルシューティング

### 10.1 よくある問題

| 症状 | 考えられる原因 | 対処法 |
|------|--------------|--------|
| ペインが表示されない | `GetActiveWindowKey()` が空 | `WindowActivate` イベントフロー確認 |
| グループ化後シートが移動しない | `ws.Move()` 失敗 | シート保護/共有ブック制約確認 |
| カラーが保存されない | CustomXML 書込失敗 | ブックの読み取り専用・権限確認 |
| ドラッグが反応しない | 閾値未到達 | `DragThreshold = 5px` 確認 |
| 色が復元されない | XML 欠損/color属性不正 | `LoadGroupsFromWorkbook` ログ確認 |
| リネーム時エラー | Excel名前制約違反 | 不正文字（\ / ? * [ ]）チェック |
| 計算が `Err.` になる | 単位パース失敗/0除算 | `ParseNumber`/`ParsePowerW` 確認 |
| ツールボタンが `?` 表示 | AddIn 検索失敗 | `FindAddInByKeywords` 多言語対応確認 |
| D&D後に順序が戻る | ポーリングタイミング | `_pollTimer.Interval` 調整 (500ms) |

### 10.2 デバッグコマンド

```powershell
# ログ出力先確認
Get-EventLog -LogName Application -Source "Qtab"

# VSTO レジストリ確認
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\Qtab"

# CustomXMLParts 内容ダンプ (Excel VBA 即時ウィンドウ)
For Each p In ActiveWorkbook.CustomXMLParts
    If InStr(p.XML, "SheetNaviGroups") > 0 Then
        Debug.Print p.XML
    End If
Next

# すべての AddIn を列挙（VBA 即時ウィンドウ）
For Each ai In Application.AddIns
    Debug.Print ai.Name & " | " & ai.Title & " | " & ai.FullName & " | Installed=" & ai.Installed
Next
```

### 10.3 計算パネルデバッグ

```csharp
// デバッグ用: 計算値と入力値をログ出力
private void AutoCalculateIfReady(string changedField)
{
    Log("AutoCalc: changed={0}, T={1}, N={2}, P={3}, target={4}", 
        changedField, _txtT.Text, _txtN.Text, _txtP.Text,
        _rbT.Checked ? "T" : (_rbN.Checked ? "N" : "P"));
    
    // ... 計算ロジック ...
    
    Log("Result: T={0}, N={1}, P={2}", _txtT.Text, _txtN.Text, _txtP.Text);
}
```

### 10.4 AddIn 検索デバッグ

```csharp
// すべての AddIn を列挙（VBA 即時ウィンドウ）
For Each ai In Application.AddIns
    Debug.Print ai.Name & " | " & ai.Title & " | " & ai.FullName & " | Installed=" & ai.Installed
Next
```

---

## 11. 今後の拡張案

### 11.1 機能拡張

| 機能 | 優先度 | 実装難易度 | 説明 |
|------|--------|-----------|------|
| グループアイコン | 中 | 低 | グループごとにアイコン設定 |
| シート検索履歴 | 中 | 中 | 最近アクセスしたシート履歴 |
| ショートカットキー | 高 | 中 | Ctrl+Q でペイン表示など |
| エクスポート/インポート | 低 | 高 | グループ設定の JSON 保存 |
| クラウド同期 | 低 | 高 | OneDrive 経由でグループ共有 |
| 計算式カスタマイズ | 中 | 中 | ユーザー定義計算式の追加 |
| ツールバッチ切替 | 中 | 低 | 複数 AddIn を一括管理 |
| テーマ切替 | 低 | 中 | ダーク/ライトモード |

### 11.2 パフォーマンス改善

- **仮想化リスト**: 大量シート（1000+）対応
- **増分描画**: OnPaint の部分更新
- **非同期ロード**: CustomXML 読込の非同期化

### 11.3 互換性拡張

- **Excel Online 対応**: Office.js への移植検討
- **Mac Excel 対応**: VSTO → Office Add-in 移行

---

## 12. ライセンス・著作権

### 12.1 ライセンス
本プロジェクトは **MIT License** の下で公開されています。

```
MIT License

Copyright (c) 2024 SheetNavi Project

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### 12.2 著作権
© 2024 SheetNavi Project. All Rights Reserved.

### 12.3 使用ライブラリ
- **Microsoft.Office.Interop.Excel**: © Microsoft Corporation
- **Microsoft.Office.Tools**: © Microsoft Corporation

---

## 13. 変更履歴

### v1.3.0 (2024-12-XX)
**新機能**
- ✨ トルク・回転数・電力の相互計算パネル追加
- ✨ Excel ツール統合管理ボタン追加（Solver / Data Analysis Tool）
- ✨ 自動計算機能（入力変更時に即座に再計算）
- ✨ 統計情報表示（Total / Hidden シート数）
- ✨ シート削除機能（右クリックメニュー）
- ✨ Excel 標準右クリックメニュー統合表示

**改善**
- 🔧 ウィンドウ別に独立した QTabNavi インスタンス管理
- 🔧 ポーリングタイマー追加（D&D等の未検出イベント対応）
- 🔧 SheetBeforeDelete イベント購読（グループ情報自動更新）
- 🔧 多言語対応 AddIn 検索（日本語/中国語/英語）
- 🔧 防抖処理の最適化（200ms）

**バグ修正**
- 🐛 D&D後の順序が元に戻る問題を修正
- 🐛 削除シートがグループに残る問題を修正
- 🐛 複数ウィンドウ時のペイン重複生成を修正

### v1.2.0 (2024-11-XX)
**新機能**
- ✨ ドラッグ&ドロップ並べ替え実装
- ✨ インライン編集（グループ名・シート名）
- ✨ 右クリックメニュー拡張（6項目）

**改善**
- 🔧 CustomXML 永続化の安定性向上
- 🔧 カラー復元ロジックの最適化

### v1.1.0 (2024-10-XX)
**新機能**
- ✨ グループ化/解除機能
- ✨ タブカラー設定
- ✨ CustomXML 永続化

### v1.0.0 (2024-09-XX)
**初回リリース**
- ✨ シート一覧表示
- ✨ 高速検索
- ✨ ペインドック切替

---

**ドキュメント更新日**: 2024-12-XX  
**バージョン**: v1.3.0  
**著者**: SheetNavi Project  
**連絡先**: [GitHub Issues](https://github.com/sheetnavi/qtab/issues)
