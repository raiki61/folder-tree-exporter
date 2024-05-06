# フォルダ構造をExcelに出力するツール

このツールは、指定したディレクトリのフォルダ構造とファイルをExcelファイルに出力します。設定ファイルで、出力対象の拡張子やフォルダ、ファイルのパターンを指定できます。

## 使い方

1. `folder-tree-exporter-config.yml` という名前の設定ファイルを作成します。
2. 設定ファイルに以下の項目を記述します。
    - `fileExtensionFilters`: 出力対象の拡張子のリスト (例: `- java`, `- xml`)
    - `excludeDirectoryPatterns`: 除外するフォルダのパターンのリスト (正規表現)
    - `excludeFilePatterns`: 除外するファイルのパターンのリスト (正規表現)
3. `DirectoryToExcel.java` の `main` メソッドを実行します。
4. 現在のディレクトリに `directory_tree.xlsx` というファイル名でExcelファイルが出力されます。

## 設定ファイルの例

```yaml
fileExtensionFilters:
  - java
  - xml
excludeDirectoryPatterns:
  - .*\\/target
excludeFilePatterns:
  - .*\\.class