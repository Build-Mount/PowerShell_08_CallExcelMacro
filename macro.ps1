using namespace Microsoft.VisualBasic;

$file = "c:\temp\test.xlsm"
Add-Type -AssemblyName Microsoft.VisualBasic
try {
    #COM�I�u�W�F�N�g�̐���
    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $True
    #�u�b�N�̓ǂݍ���
    $xlBook = $xlApp.Workbooks.Open($file)
    
    $xlBook.Worksheets("Sheet1").Range("A1").Text                      #A1�Z���̃e�L�X�g���z�X�g�ɕ\��
    $xlBook.Worksheets("Sheet1").Range("A1").End(-4121).Text           #A1�Z�����牺�����Ɉړ����Ă��̃e�L�X�g���z�X�g�ɕ\��
    $xlBook.Worksheets(1).Cells(3,3).Value  = "60"                     #Cells(3,3)�̃Z���̒l��60�ɐݒ�
    $xlApp.Range("A1").Text                                            #xlApp�o�R��A1�Z���̃e�L�X�g���z�X�g�ɕ\��

    #SubShowMessage��Sub�֐������s
#    $xlApp.Run("test.xlsm!SubShowMessage", "Hello world")

    #FuncShowMessage��Function�֐������s
    $resultMessage = $xlApp.Run("FuncShowMessage", "Bye world")
    #resultMessage�̕\��
    Write-Host $resultMessage
    
    #FuncShowMessage��Function�֐��̖߂�l���p�C�v��������Write-Host�ŏo��
    $xlApp.Run("FuncShowMessage", "Bye Bye") | Write-Host
    
    #FuncGetCell��Function�֐��ŁAA1�Z����Range���擾
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

    #�u�b�N�̃N���[�Y
    $xlBook.Close()
} finally {
    #Excel�A�v���̏I��
    $xlApp.Quit()
    #COM�I�u�W�F�N�g�̉��
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlApp) | Out-Null
}