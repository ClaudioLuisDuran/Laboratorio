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

fecha_actual = Now()

Dia = Day(fecha_actual)
Mes = Month(fecha_actual)
Anio = Year(fecha_actual)

fecha_ok = Mes &"/"& Dia &"/"& Anio


'grabo fecha en odontograma y luego redirijo a odontograma completo

set oConn =  Server.CreateObject("ADODB.Connection")
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

SQL = "select * from corona where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

RS("fecha_consulta") = fecha_ok

RS.Update
RS.Close             
set RS=nothing


SQL2 = "select * from extraccion where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL2, oConn,3,3

RS("fecha_consulta") = fecha_ok

RS.Update
RS.Close             
set RS=nothing

SQL3 = "select * from odontograma where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL3, oConn,3,3

RS("fecha_consulta") = fecha_ok

RS.Update
RS.Close             
set RS=nothing

SQL4 = "select * from odontograma2 where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL4, oConn,3,3

RS("fecha_consulta") = fecha_ok

RS.Update
RS.Close             
set RS=nothing

oConn.Close
set oConn = nothing

' cargo fecha ultima consulta en FICHA

set oConn =  Server.CreateObject("ADODB.Connection")
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")

SQL = "select * from fichas where paciente = " & request("paciente") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL, oConn,3,3

DIM estado
estado = "En curso"

RS("fecha_consulta") = fecha_ok
RS("estado_consulta") = estado

RS.Update
RS.Close             
set RS=nothing

oConn.Close
set oConn = nothing

'........................................................


Response.Redirect "diagnostico.asp?paciente=" & paciente
 
%>