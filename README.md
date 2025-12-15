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

## 使い方

```bash
# 依存関係のインストール
pip install -r requirements.txt

# ハイブリッドモード（推奨）
python pdf_to_pptx_advanced.py input.pdf output.pptx

# OCRモード（NotebookLMのPDF向け）
python pdf_to_pptx_advanced.py input.pdf output.pptx --mode ocr
```

詳細なドキュメントはファイル内のコメントを参照してください。

## ライセンス

MIT License
