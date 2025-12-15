# PDF to Editable PPTX Converter

NotebookLMで生成したスライドPDFを含む、様々なPDFを編集可能なPowerPointファイルに変換するツール。

## 特徴

- **PDF要素分解**: テキスト、画像、図形、表を個別に抽出
- **編集可能なPPTX**: 各要素を編集可能な状態で配置
- **3つのモード**:
  - `extract`: PDF内のテキストを直接抽出（通常のPDF向け）
  - `ocr`: OCRでテキスト認識（画像ベースのPDF向け、NotebookLM等）
  - `hybrid`: 両方を組み合わせ（自動判断、推奨）
- **文字化け自動修正**: ftfyによる自動修正 + 手動マッピング
- **ローカル完結**: API不要、すべてローカルで処理

## セットアップ

### 1. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. Tesseract OCRのインストール（OCRモード使用時）

**Windows:**
```bash
winget install --id UB-Mannheim.TesseractOCR -e
```

### 3. 日本語言語データの追加（Windows）

```bash
# ステップ1: tessdataフォルダ作成
mkdir ~/tessdata

# ステップ2: 日本語+英語の言語データをダウンロード
curl -L -o ~/tessdata/jpn.traineddata https://github.com/tesseract-ocr/tessdata/raw/main/jpn.traineddata
curl -L -o ~/tessdata/eng.traineddata https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata
```

**重要**: 日本語OCRには`jpn.traineddata`と`eng.traineddata`の両方が必要です。
スクリプトは自動的に`~/tessdata/`を検出し、`TESSDATA_PREFIX`環境変数を設定します。

## 使い方

```bash
# ハイブリッドモード（推奨）
python pdf_to_pptx_advanced.py input.pdf output.pptx

# OCRモード（NotebookLMのPDF向け）
python pdf_to_pptx_advanced.py input.pdf output.pptx --mode ocr

# 背景画像なし（テキストのみ）
python pdf_to_pptx_advanced.py input.pdf output.pptx --no-background
```

## ライセンス

MIT License
