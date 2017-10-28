# 最近書いたもの(主に拡張子が.vbs)  
特に注意書きのないものは、Windows用のInDesign上で動くVBScript

## vbs/このドキュメントの格納パスをExplorerで開く/

### このドキュメントの格納パスをExplorerで開く.vbs
app.activeDocumentがsavedである時、右ペインでそのファイルが選択された状態のエクスプローラを開く(バージョン3)
* jsx/このドキュメントの格納パスをExplorerで開く/にあるExtend Script版の次世代版
	* よい点：実行してもDOS窓が開かない
	* よい点：ない方がよいファイルI/Oはない
	* 悪い点：特にない
  
  

### activeDocumentの格納パスをExplorerで開く.bas
app.activeDocumentがsavedである時、右ペインでそのファイルが選択された状態のエクスプローラを、Excelから開く
* vbs/このドキュメントの格納パスをExplorerで開く/にあるExtend Script版のExcel VBA版
	* 実行時点でInDesignが起動しているかどうかを調べる機能は載せていないので注意
  
  

