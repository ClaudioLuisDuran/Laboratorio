<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>

<%

Set oConn = Server.CreateObject("ADODB.Connection")

' grabo escrito

set RS = Server.CreateObject("ADODB.Recordset")  


oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")

             RS.Open "valorespub",oConn,2,2
             
             RS.AddNew
             
                          RS("descripcion") = request.form("descripcion")
             RS("valor") = request.form("valor")

             
             RS.Update
             RS.Close
             
set RS=nothing

oConn.Close

Response.Redirect "prestaciones_pub.asp"

%>

</body>

</html>