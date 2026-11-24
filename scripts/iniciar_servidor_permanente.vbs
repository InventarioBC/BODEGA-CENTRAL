Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL"
WshShell.Run """C:\Users\Bodeg\AppData\Local\Programs\Python\Python312\pythonw.exe"" ""G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL\scripts\servidor_bodega.py"" 8080", 0, False
