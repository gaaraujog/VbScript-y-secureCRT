# VbScript-y-secureCRT
'is an script made for proccess data from a conection via ssh to a Server and then log in on a router, gets information from the router and then edit a textfile for strat a 2dn step for the scrip that is test a lot of IP adress response

#$language = "VBScript"
#$interface = "1.0"
Imports System.IO
Imports System.Text.RegularExpressions
crt.Screen.Synchronous = True

' aqui declaro el tipo de script que luego voy a editar
Set FSO = CreateObject("Scripting.FileSystemObject") 
Set logfile = FSO.CreateTextFile("C:\Users\araujo.gustavo\Documents\logfile2.txt", True) 



Sub Main
	'mensaje para seleccionar que ASR revisar
    ASR = MsgBox("quieres ver el ASR Principal primero??",3,"vamos a verificar todas las IP publicas")   
	Select case ASR 
	   case 6 'Entra en el ASR PPL y obtiene todas las destination IP de los tuneles
	    'una vez obtiene las IP de los tuneles las guarda en un archivo de texto para ser modificado y correr el script sin conexion a la red
		ASRPPAL = MsgBox("OK, revisemos",0,"ASR PRINCIPAL")
		crt.Screen.Send "telnet 172.26.30.77" & chr(13)
		crt.Screen.WaitForString "Username: "
		crt.Screen.Send "impsat" & chr(13)
		crt.Screen.WaitForString "Password: "
		crt.Screen.Send "impsat" & chr(13)
		crt.Session.Log False 'deja de grabar el log del secure
		crt.Session.LogFileName = "C:\Users\araujo.gustavo\Documents\logfile.txt" '& a(i) & ".log"  'Define log file name
		crt.Session.LogUsingSessionOptions 'comienza a grabar el log del secure
		crt.Screen.Send "show run | in  tunnel destination" & chr(13)
		crt.Screen.WaitForString " --More-- "
		crt.Screen.Send "              "   
	   Case  7 'Entra en el ASR BKP y obtiene todas las destination IP de los tuneles
	    'una vez obtiene las IP de los tuneles las guarda en un archivo de texto para ser modificado y correr el script sin conexion a la red
		ASRBCKP = MsgBox("Entendido, vamos entonces a revisar el ASRBKP",0,"ASRBCKP")
		crt.Screen.Send "telnet 172.26.30.78" & chr(13)
		crt.Screen.WaitForString "Username: "
		crt.Screen.Send "impsat" & chr(13)
		crt.Screen.WaitForString "Password: "
		crt.Screen.Send "impsat" & chr(13)
		crt.Screen.Send "show run | in  tunnel destination" & chr(13)
		crt.Screen.WaitForString " --More-- "
		crt.Screen.Send "                        " & chr(13)
		crt.Screen.Send "exit" & chr(13)
	   Case 2 'mision abortada
		MsgBox "Entendido, entoces jodete, me voy",0,"Lo Cancelaste"
	End Select
	'crt.Session.Log False 'Turn off log session
	'Set logfile = FSO.CreateTextFile(C:\Users\araujo.gustavo\Documents\logfile2.txt, True, True)  	
	'WScript.Sleep 1000
	dteWait = DateAdd("s", 4, Now())
	Do Until (Now() > dteWait)
	Loop
 fin = MsgBox("a continuacion se va a abrir la lista de IP Destination en un archivo .txr, por favor usa ctrl +r y reemplaza todo destination ip por start ping para continar" , 0, "REEMPLAZA PARA CONTINUAR")
 End Sub
