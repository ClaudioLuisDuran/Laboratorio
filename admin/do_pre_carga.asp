<% @ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

</head>

<body>

<%

DIM periodo_ok
periodo_ok = request.form("periodo")


DIM oConn
DIM rs

' grabo nombre de planilla

Set oConn = Server.CreateObject("ADODB.Connection")

set RS = Server.CreateObject("ADODB.Recordset")  

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")

             RS.Open "planillas",oConn,2,2
             
             RS.AddNew
             
             RS("periodo") = request.form("periodo")
             RS("pagada") = "No"
             
             RS.Update
             RS.Close
             
set RS=nothing

oConn.Close

  Session("allow_shopp") = True
  Session("periodo") = periodo_ok
  Session.Timeout = 600
  Session("mes") = ""
Session("anio") = ""
Session("profesional") = ""
Session("cupon") = ""
Session("afiliado") = "" 
Session("nombre") = ""
Session("valorcupon") = ""
Session("fechaosep") = ""

Response.Redirect "carga.asp"


%>