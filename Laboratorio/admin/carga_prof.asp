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

             RS.Open "profesionales",oConn,2,2
             
             RS.AddNew
             
             RS("profesional") = request.form("profesional")
             RS("matricula") = request.form("matricula")
             RS("email") = request.form("email")
             RS("usuario") = request.form("matricula")
             RS("password") = request.form("matricula")
             RS("visitas") = 0
             RS("activo") = True
             RS("fechaalta") = now
             
             RS.Update
             RS.Close
             
set RS=nothing

oConn.Close

Response.Redirect "profesionales.asp"

%>

</body>

</html>