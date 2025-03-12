' Obtener todos los discos duros del equipo y su espacio usado (en porcentaje)
' Permite pasarle como argumento la unidad o unidades de las que devolverá el espacio usado
' Si no se le pasan argumentos, devolverá todas las unidades
option explicit
on error resume next
dim oWMI, discoActual, listaDiscos, numArgumentos
dim parametros, i, espacioLibre, espacioTotal, espacioUsado

' Obtenemos el número de argumentos (pasados por parámetro en la línea de comandos)
numArgumentos = Wscript.Arguments.Count
' Obtenemos los argumentos/parámetros (pasados por parámetro en la línea de comandos)
set parametros = CreateObject("Scripting.Dictionary")
for i = 0 to numArgumentos - 1
    parametros.Add Wscript.Arguments(i), i
next

' Obtenemos la información de los discos con WMI
set oWMI = GetObject ("winmgmts:\\.\root\cimv2")
set listaDiscos = oWMI.ExecQuery ("Select * from Win32_LogicalDisk")
' Recorremos todos los discos obtenidos para mostrar su información
' Siempre que se hayan pasado por parámetro o bien no se haya pasado ningún parámetro
for each discoActual in listaDiscos
	' Si no se han pasado argumentos o si el argumento pasado 
	' coincide con el nombre de la unidad actual
	' Se obtiene su espacio usado y se muestra por consola
	if numArgumentos = 0 or parametros.Exists(discoActual.Name) then
		' Filtramos los discos duros (type 3), para no mostrar 
		' otro tipo de unidades de almacenamiento como pendrives
		if (discoActual.FreeSpace <> "") and (discoActual.DriveType = 3) then
			espacioLibre = discoActual.FreeSpace
			espacioTotal = discoActual.Size
			' Obtenemos el porcentaje de espacio usado en base al espacio libre y el espacio total
			espacioUsado = round (100 - (espacioLibre / espacioTotal) * 100, 2)
			' Mostramos el resultado por consola
			Wscript.StdOut.WriteLine discoActual.Name & " -> " & espacioUsado & "%"
			Wscript.StdOut.flush
		end if
	end if
next
WScript.Quit