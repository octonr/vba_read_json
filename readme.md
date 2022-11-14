# Read Json VBA

## JSONファイルの読み込み
### jsonfileのpathをどうするか。
  - fn:`read_json()`にjsonfileのpathを投げて、パースして返してもらう。</br>
    パース結果を格納した`json`は`CallByName(json, {key}, vbGet)`で取得する。

  - vbaの連想配列は以下の様な形式
    ```
    Sub macro1()
        'Dictionaryオブジェクトの宣言
        Dim myDic As Object
        Set myDic = CreateObject("Scripting.Dictionary")
        
        'Dictionaryオブジェクトの初期化、要素の追加
        myDic.Add "orange", 100
        myDic.Add "apple", 200
        myDic.Add "melon", 300
        
        'Dictionaryオブジェクトの要素の参照
        Dim str As String, i As Integer
        Dim Keys() As Variant
        Keys = myDic.Keys
        For i = 0 To 2
            str = str & Keys(i) & " : " & myDic.Item(Keys(i)) & vbCrLf
        Next i
        
        MsgBox str, vbInformation
    End Sub
    -----------------
    msgbox結果:
        orange: 100
        apple: 200
        melon: 300
    -----------------
    ```
  - 配列を取得するには以下のようにする。
    ```
    For Each item In Split(CallByName(config, "ref2", VbGet), ",")
        Debug.Print item 'String
        Debug.Print Val(item) '数値化
    Next
    ```
    
### 64bitだと`CreateObject("ScriptControl")`が使えない
  - [こちら](https://qiita.com/nukie_53/items/297e524bcc8e43f9b5d1)より`JSFunc`を使用