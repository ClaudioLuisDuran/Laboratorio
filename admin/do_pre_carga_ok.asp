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
periodo_ok = request("periodo")

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