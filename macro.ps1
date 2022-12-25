using namespace Microsoft.VisualBasic;

$file = "c:\temp\test.xlsm"
Add-Type -AssemblyName Microsoft.VisualBasic
try {
    #COMオブジェクトの生成
    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $True
    #ブックの読み込み
    $xlBook = $xlApp.Workbooks.Open($file)
    
    $xlBook.Worksheets("Sheet1").Range("A1").Text                      #A1セルのテキストをホストに表示
    $xlBook.Worksheets("Sheet1").Range("A1").End(-4121).Text           #A1セルから下方向に移動してそのテキストをホストに表示
    $xlBook.Worksheets(1).Cells(3,3).Value  = "60"                     #Cells(3,3)のセルの値を60に設定
    $xlApp.Range("A1").Text                                            #xlApp経由でA1セルのテキストをホストに表示

    #SubShowMessageのSub関数を実行
#    $xlApp.Run("test.xlsm!SubShowMessage", "Hello world")

    #FuncShowMessageのFunction関数を実行
    $resultMessage = $xlApp.Run("FuncShowMessage", "Bye world")
    #resultMessageの表示
    Write-Host $resultMessage
    
    #FuncShowMessageのFunction関数の戻り値をパイプ処理してWrite-Hostで出力
    $xlApp.Run("FuncShowMessage", "Bye Bye") | Write-Host
    
    #FuncGetCellのFunction関数で、A1セルのRangeを取得
    $range = $xlApp.Run("FuncGetCell", "A1")
    Write-Host $range.Text
    Write-Host $xlApp.Range("A1").Text
    
    $myClass = $xlApp.Run("CreateHelloClass")
	Write-Host $myClass.ClassName
    $myClass.Hello("Japan")

	[Microsoft.VisualBasic.Interaction]::MsgBox("Test")
	[Microsoft.VisualBasic.Strings]::Len("Test")

	[Interaction]::MsgBox("Test")
	[Strings]::Len("Test")

    #ブックのクローズ
    $xlBook.Close()
} finally {
    #Excelアプリの終了
    $xlApp.Quit()
    #COMオブジェクトの解放
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlApp) | Out-Null
}