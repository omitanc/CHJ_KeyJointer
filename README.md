# CHJ_KeyJointer
NINJAL 日本語歴史コーパスのキー列の内容を、ページ番号列の数字を基準に結合するVBAです。
このVBAは、Mac OS版のExcelでは動作しません。

<br>

# 使用方法

1. [中納言 CHJ](https://chunagon.ninjal.ac.jp/chj/search)より、下記の検索条件で検索結果をダウンロード。
2. (作品名).csvにして任意のディレクトリに保存。
3. csvファイルと同じディレクトリ内に outputs フォルダを作成。
4. csvファイルをExcelで開く。
5. csvファイルのシートの名前を original に変更。
6. Excelの開発タブからVisual Basicを開き.basファイルをインポート
7. Ctrl+Sで保存の際、「いいえ」を選択して、.xlsmの拡張子で保存
8. VBAのスクリプトを実行。
9. Xボタンを押して保存せずにExcelを終了。
10. outputsフォルダ内に、データ整形後のcsvが生成されていれば完了。

<br>

# CHJでの検索条件式
検索には、「長単位検索」の「検索条件式で検索」を使用します
```IN subcorpusName="平安-仮名文学" AND 作品名="竹取物語"```の部分には、任意のサブコーパス名と作品名を指定してください。
例↓

```
キー: (語種="和" OR 語種="漢" OR 語種="外" OR 語種="混" OR 語種="固" OR 語種="記号")
  IN subcorpusName="平安-仮名文学" AND 作品名="竹取物語"
  WITH OPTIONS tglKugiri="" AND tglBunKugiri="#" AND limitToMainText="1" AND limitToSelfSentence="1" AND tglWords="40" AND unit="2" AND encoding="UTF-16LE" AND endOfLine="CRLF";
```
