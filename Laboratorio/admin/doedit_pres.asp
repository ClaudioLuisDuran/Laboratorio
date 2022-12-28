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

Id = Request("Id")



Set oConn = Server.CreateObject("ADODB.Connection")

' actualizo


oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")


 SQL = "select * from valores where Id = " & request("Id") & ""
 
set RS = Server.CreateObject("ADODB.Recordset") 
 rs.Open SQL, oConn,3,3


             RS("codigo") = request.form("codigo")
             RS("descripcion") = request.form("descripcion")
             RS("valor") = request.form("valor")
             
             RS.Update
             RS.Close
             
set RS=nothing

oConn.close

Response.Redirect "prestaciones.asp"
 
%>