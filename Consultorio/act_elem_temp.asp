<%@ Language=VBScript %>
<% Response.Buffer = True %>

 <html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title></title>



</head>

<body>

<% 

paciente = request("paciente")
elemento = request("elemento")
extraccion = request("extraccion")
corona = request("corona")
eleccion = request("eleccion")
'response.write eleccion


cara1 = request("1")
cara2 = request("2")
cara3 = request("3")
cara4 = request("4")
cara5 = request("5")
'response.write cara1
'response.write cara2
'response.write cara3
'response.write cara4
'response.write cara5

response.write paciente
response.write elemento
response.write extraccion
response.write corona
response.write eleccion

Set oConn = Server.CreateObject("ADODB.Connection")

'..............................................................

' Si "eleccion" es ANULAR, veo de que se trata

if eleccion = "Anular" then

' Verifico Condicion 1: Extraccion
' Verifico extraccion y tipo de extraccion

if extraccion = "Si" or extraccion = "Ei" or extraccion = "ei" then
'significa que va a anular entonces anulo la marca de extraccion

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from extraccion where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

RS(elemento) = "No"
RS.Update
RS.Close
             
set RS=nothing
oConn.Close
set oConn = nothing

end if ' fin de IF extraccion

'Fin de verificacion de Condicion 1: Extraccion

'...........................................................

'Verifico Condicion 2: Corona

if corona = "Si" or corona = "Cor" then
'significa que va a anular entonces anulo la marca de Corona

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from corona where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

RS(elemento) = "No"
RS.Update
RS.Close
             
set RS=nothing


oConn.Close
set oConn = nothing
end if ' fin de IF Corona
'Fin de verificacion de Condicion 2: Corona

'.........................................................

' Condicion 3 : anulado de caras pintadas del elemento elegido

if corona = "No" and extraccion = "No" then

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from odontograma2 where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

elemcara1 = elemento & 1
elemcara2 = elemento & 2
elemcara3 = elemento & 3
elemcara4 = elemento & 4
elemcara5 = elemento & 5

response.write cara1
response.write elemcara1

if cara1 = "ON" then
RS(elemcara1) = "FFFFFF"
end if

if cara2 = "ON" then
RS(elemcara2) = "FFFFFF"
end if

if cara3 = "ON" then
RS(elemcara3) = "FFFFFF"
end if

if cara4 = "ON" then
RS(elemcara4) = "FFFFFF"
end if

if cara5 = "ON" then
RS(elemcara5) = "FFFFFF"
end if

RS.Update
RS.Close
             
set RS=nothing


oConn.Close
set oConn = nothing

end if


else ' sino anula, entonces graba nuevas pintadas o indicaciones

' Verifico si va a extraer elemento elegido

if eleccion = "ExtraccionSI" or eleccion = "ExtraccionEI" then

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from extraccion where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

if eleccion = "ExtraccionSI" then
RS(elemento) = "Si"
else
RS(elemento) = "ei"
end if

RS.Update
RS.Close
             
set RS=nothing
oConn.Close
set oConn = nothing

else

'Verifico si va a Coronar elemento elegido
if eleccion = "CoronaSi" or eleccion = "CoronaCor" then

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from corona where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

if eleccion = "CoronaSi" then
RS(elemento) = "Si"
else
RS(elemento) = "Cor"
end if

RS.Update
RS.Close             
set RS=nothing
oConn.Close
set oConn = nothing


else
'Verifico pintadas de caras de elemento elegido
if eleccion = "Rojo" or eleccion = "Azul" then

if eleccion = "Rojo" then
color = "FF0000"
else
if eleccion = "Azul" then
color = "0000FF"
end if
end if

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from odontograma2 where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

elemcara1 = elemento & 1
elemcara2 = elemento & 2
elemcara3 = elemento & 3
elemcara4 = elemento & 4
elemcara5 = elemento & 5

response.write cara1
response.write elemcara1

if cara1 = "ON" then
RS(elemcara1) = color
end if

if cara2 = "ON" then
RS(elemcara2) = color
end if

if cara3 = "ON" then
RS(elemcara3) = color
end if

if cara4 = "ON" then
RS(elemcara4) = color
end if

if cara5 = "ON" then
RS(elemcara5) = color
end if

RS.Update
RS.Close
             
set RS=nothing


oConn.Close
set oConn = nothing


end if
end if
end if

end if ' Fin IF anulacion
' Fin de actualizacion de elemento


'Actualiza Historial

accion = request.form("accion")
accion = accion & " Elemento "
accion = accion & elemento

set RS = Server.CreateObject("ADODB.Recordset")  
set oConn2 =  Server.CreateObject("ADODB.Connection")

oConn2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/historial.mdb")

             RS.Open "historial",oConn2,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("responsable") = request.form("odontologo")
             RS("accion") = accion
             RS("fecha") = request.form("fecha")                        
             RS.Update
             RS.Close
             
set RS=nothing
oConn2.Close

'........................................................


Response.Redirect "temporario.asp?elemento=" & elemento & "&paciente=" & paciente
 
%>