# PPTX スライド自動生成スキル v3.1

台本（スプレッドシート / CSV / Markdown）から、Googleスライド互換のPPTXファイルを **一発生成** するAntigravityスキル。

## 2つのパターン

| パターン | 用途 | 画風 |
|---------|------|------|
| `source` | 投資系テロップ | 赤帯ヘッダー＋白背景 |
| `character` | キャラ系テロップ | キャラ＋吹き出し＋コラージュ |

## セットアップ

### 1. Clone

```bash
git clone https://github.com/YOUR_USERNAME/pptx-skill.git
```

### 2. Antigravity skills に配置

```bash
# Windows
mklink /J "%USERPROFILE%\.gemini\antigravity\skills\pptx-skill" "cloneしたパス\pptx-skill"

# Mac
ln -s /path/to/pptx-skill ~/.gemini/antigravity/skills/pptx-skill
```

### 3. 依存パッケージのインストール

```bash
cd pptx-skill
npm install
```

### 4. 素材フォルダの準備

```
C:\KEI IWASAKI\assets\{チャンネル名}\
├── characters\     # キャラPNG（透過）
├── effects\        # エフェクトPNG
└── manifest.json   # 自動生成される
```

ファイル名規則：`カテゴリ_番号.png`（例：`笑顔_1.png`、`爆発_1.png`）

## 使い方

Antigravityに「スライド作って」と言うだけ。5ステップのワークフローで自動進行します。

詳細は [SKILL.md](./SKILL.md) を参照。
