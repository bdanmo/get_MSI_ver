$pathToMSI = Read-Host -Prompt "Enter path to MSI file"

$windowsInstaller = New-Object -com WindowsInstaller.Installer
$database = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $windowsInstaller, @($pathToMSI, 0))

$query = "SELECT `Value` FROM `Property` WHERE `Property` = 'ProductVersion'"
$view = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $database, ($query))

$view.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $view, $Null)

$record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $Null, $view, $Null)
$version = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)

write-host $version

Read-Host -Prompt "Press Enter to exit"

