<img width="2172" height="724" alt="Image" src="https://github.com/user-attachments/assets/2df8698b-26e5-42df-9778-fda65ab99337" />

🌎[English](https://github.com/ZeeZeX/VBA-UserForm-Transpiler/blob/main/README.md)<br><br>

## 概要

**VBA UserForm Transpiler** は、Excel VBA のユーザーフォームを、さまざまなターゲットプラットフォームおよびプログラミング言語向けの GUI コードに変換するための、GUI ベースのツールです。

Microsoft Excel VBA で作成されたユーザーフォームを、複数のターゲットプラットフォームやプログラミング言語向けの GUI コードに変換できます。

| ターゲット | 生成されるコード | GUI フレームワーク |
|---|---|---|
| Python | `.py` | Tkinter |
| PowerShell | `.ps1` / `.bat` | WinForms |
| HTML | `.html` | HTML / CSS |

<img width="1677" height="587" alt="Image" src="https://github.com/user-attachments/assets/491f411e-6e84-4dbf-a921-eeb193c4542e" />
<img width="1248" height="448" alt="Image" src="https://github.com/user-attachments/assets/9e38f2f8-5232-4adf-86d3-083dd095abfb" />

## 主な機能

- **統合インターフェース** — 複数の変換エンジンを、単一の GUI から操作できます。
- **複数の出力先に対応** — VBAユーザーフォームを、さまざまなターゲットプラットフォームおよびプログラミング言語向けの GUI コードに変換できます（例：Python/Tkinter、PowerShell/WinForms、HTML/CSS）。
- **ブックの選択** — 変換対象のユーザーフォームが含まれる Excel ブックを選択できます。
- **ユーザーフォームの選択** — 選択したブックから、変換するユーザーフォームを指定できます。
- **再読み込み** — 利用可能なブックとユーザーフォームの一覧を更新できます。
- **変換オプション** — 選択したターゲットに応じて、追加のオプションを設定できます。
- **画像の処理** — ユーザーフォームに埋め込まれた画像のエクスポート方法や参照方法を設定できます。

## 使い方

<img width="423" height="499" alt="Image" src="https://github.com/user-attachments/assets/22eee50d-b973-45b7-a921-efc6fdee03f1" />

本アプリケーションでは、簡単な手順でユーザーフォームを変換できます。

開始する前に、変換対象のユーザーフォームを含む Excel ブックを Microsoft Excel で開いてください。

1. 現在開いているブックの一覧から、変換元のブックを選択します。
2. 変換するユーザーフォームを選択します。
3. 出力先のプラットフォームを選択します。
4. 利用可能な変換オプションを設定します。
5. **Convert** をクリックします。

選択した変換先に関連するオプションが、インターフェースに自動的に表示されます。

## 必要条件

### Excel のセキュリティ設定

本ツールを使用するには、Excel で次の設定を有効にする必要があります。

1. ブックを開きます。
2. **ファイル → オプション → トラスト センター → トラスト センターの設定** を開きます。
3. **マクロの設定** を選択します。
4. **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** にチェックを入れます。
5. **OK** をクリックします。

### 動作環境

- Microsoft Windows
- VBA に対応した Microsoft Excel
- 変換対象のユーザーフォームを含むExcelブック

元となる変換エンジンは、Windows 10/11、および 32 ビット版・64 ビット版の Office を含む、さまざまな Windows と Microsoft Office のバージョンでテストされています。

Microsoft Excel 2016 以降を推奨します。

<details>
<summary>古いバージョンでのOfficeでの起動について</summary>

Office2000, Office XP, Office2003でこのアプリケーションを使用する場合はMicrosoft Office Compatibility PackをPCにインストールする必要があります。
[ここ](http://web.archive.org/web/20120201204132/http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=3)よりファイルをダウンロードしインストール後、PCの再起動をしてください。
直接リンクからダウンロードする場合は、以下のURLをブラウザのアドレスバーに貼り付けてください。
```
http://web.archive.org/web/20111016032501/http://download.microsoft.com/download/9/2/2/9222d67f-7630-4f49-bd26-476b51517fc1/FileFormatConverters.exe
```
Microsoft Office Compatibility Packをインストール後、xlsmファイルを開くとxlsファイルに変換された状態で開きますが、一時フォルダにxlsファイルが配置された状態になり、アプリケーションが正常に動作しなくなるため「名前をつけて保存」でxlsファイルを別のフォルダに保存してから使用してください。

</details>

### 生成されたコードの実行環境

選択したターゲットに応じて、次の実行環境が必要です。

| ターゲット | 実行環境 |
|---|---|
| Python / Tkinter | Python 3.4 以降 |
| PowerShell / WinForms | Windows PowerShell 5.1 以降 |
| HTML / CSS | 対応する Web ブラウザー |

生成されたコードや実行環境によっては、追加の要件が必要になる場合があります。

## 変換オプション

利用できるオプションは、選択した出力先によって異なります。

| Option | Python (Tkinter) | PowerShell (WinForms) | HTML/CSS |
|---|:---:|:---:|:---:|
| Use Class/Prefix | ✓ | ✓ | ✓ |
| Save As `.bat` | — | ✓ | — |
| No Mainloop | ✓ | ✓ | — |
| Unique Style Name | ✓ | — | — |
| Language Attribute | — | — | ✓ |
| Image Mode | ✓ | ✓ | ✓ |

### オプションの説明

- **Use Class/Prefix** — ターゲットプラットフォームに応じて、クラスベースの構造で出力するか、生成される識別子に指定したプレフィックスを付与します。これにより、生成コードを整理しやすくなり、名前の競合を防止できます。

- **Save As `.bat`** — PowerShellスクリプトをダブルクリックで起動可能な`.bat`ファイルとして保存します。

- **No Mainloop** — 生成コードからメインイベントループまたは表示メソッドを省略します。生成した GUI を既存のアプリケーションに組み込む場合や、GUI のライフサイクルを外部で管理する場合に便利です。

- **Unique Style Name** — アプリケーション内のほかのスタイルとの名前の競合を防ぐため、Tkinter のスタイルに一意の名前を付けて生成します。

- **Language Attribute** — 生成される HTML ドキュメントの言語属性（`lang`）を指定します。ブラウザーやその他のツールに、ドキュメントの言語を伝えるために使用できます。

- **Image Mode** — VBAユーザーフォームに埋め込まれた画像を、変換時にどのように扱うかを指定します。利用できるモードは、選択したターゲットプラットフォームによって異なります。
  - `file` — 画像を個別のファイルとしてエクスポートします。
  - `disabled` — 画像処理を無効にします。
  - `reference-only` — 画像をエクスポートせず、既存の画像ファイルへの参照を生成します。
  - `base64` — Base64エンコードしたデータとして、画像を生成コード内に直接埋め込みます。
  - `base64-dict` — ターゲット言語で利用可能な連想配列（辞書）に、Base64エンコードした画像データを格納します。
  - `base64-json` — Base64エンコードした画像データを、外部JSONファイルに格納します。
  - `base64-json-reference` — 既存のJSON画像ファイルをエクスポートせずに参照します。

## 出力先フォルダ

専用の出力用フォルダが`VBA-UserForm-Transpiler.xlsm`と同じディレクトリに生成され、すべての生成されたファイルはこのフォルダ内に格納されます。

出力先フォルダの名前は選択した変換先によって異なります:

| 変換先 | 出力先フォルダ名 |
|---|:---:|
| Python (Tkinter) | `VBAForm2Tkinter_output` |
| PowerShell (WinForms) | `VBAForm2PS_output` |
| HTML/CSS | `VBAForm2HTML_output` |

### ディレクトリ構造

以下の例はPython (Tkinter)の場合の出力ディレクトリ構造です:

```
ApplicationFolder/
├─ VBA-UserForm-Transpiler.xlsm
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ exported images...
```

エクスポートされた画像や補助ファイルを含む、生成されたすべてのファイルは、対応する出力フォルダに保存されます。

## 変換の仕組み

VBA UserForm Transpiler は、複数の変換エンジンを統合することで、異なるターゲットプラットフォームおよびプログラミング言語に対応しています。

```text
                   VBA UserForm Transpiler
                            |
                            |
                     ワークブックの選択
                            |
                            v
                    ユーザーフォームの選択
                            |
                            v
                    変換先のプラットフォームを選択
                            |
              +-------------+-------------+
              |             |             |
              v             v             v
           Python        PowerShell      HTML/CSS
           Tkinter        WinForms        Browser
              |             |             |
              v             v             v
           output.py     output.ps1     output.html
                         output.bat
```

各変換先では、それぞれに対応する変換エンジンを使用します。VBA を単一の中間形式に変換してから各形式へ展開する方式ではありません。

この方式により、元のプロジェクトの機能や変換時の動作を維持しながら、一貫したユーザー体験を提供します。

## 変換時の制限事項

本ツールは、VBAユーザーフォームの視覚的な構造とプロパティの変換を主な目的としています。

以下の点にご注意ください。

- 生成される GUI レイアウトは、元のVBAユーザーフォームを近似的に再現したものです。
- すべての VBA コントロールやプロパティを、変換先のフレームワークで完全に同じ形で表現できるわけではありません。
- サポートされていないコントロールが含まれている場合、変換に失敗することがあります。
- 一部のプロパティは、VBA と変換先の GUI フレームワークで動作が異なります。
- VBA のイベントプロシージャやアプリケーション固有のビジネスロジックは、変換先環境で同等に動作するイベントハンドラーや実行可能コードへ自動変換されません。
- 生成された GUI を完全に機能させるには、追加のプログラミングが必要になる場合があります。

そのため、生成されたコードは、VBA ベースの GUI を移植する際の出発点として扱ってください。元の VBA アプリケーションを完全に置き換えるものではありません。

対応コントロールや変換上の制限について詳しくは、各元プロジェクトのドキュメントを参照してください。

## 元プロジェクト

本プロジェクトは、以下のオープンソースプロジェクトを統合しています。

| プロジェクト | 説明 |
|---|---|
| [VBAForm2Tkinter](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter) | VBAユーザーフォームから Python Tkinter への変換 |
| [VBAForm2PowerShell](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell) | VBAユーザーフォームから PowerShell WinForms への変換 |
| [VBAForm2HTML](https://github.com/GUI-Conversion-Tools/VBAForm2HTML) | VBAユーザーフォームから HTML/CSS への変換 |

各プロジェクトには、対応コントロール、変換時の動作、パラメーター、制限事項などを含む、独自の詳細なドキュメントが用意されています。

特定の変換先についてさらに詳しい情報が必要な場合は、各プロジェクトの元リポジトリを参照してください。
