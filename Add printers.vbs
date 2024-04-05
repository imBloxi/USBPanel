strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Network = TRUE AND NOT Name LIKE '%PrinterShareNameHere%'")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next
WScript.Echo "Finished deleted all network printers."



'**** Add the printer. ****
'*                                        *
Set wshNet = CreateObject("WScript.Network")
wshnet.AddWindowsPrinterConnection "\\server\printershare"
wshnet.SetDefaultPrinter "\\server\printershare

WScript.Echo "Finished adding printer."
