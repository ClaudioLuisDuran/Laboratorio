<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%
DIM UserName 
UserName = Session("usuario")
DIM Password 
Password = Session("password")
DIM uConn
DIM RSu
DIM yes
DIM error
DIM odontologo
DIM matricula

set uConn =  Server.CreateObject("ADODB.Connection")

uConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/login.mdb")
Set RSu = uConn.Execute("select * from registrados where usuario = '" & UserName & "'  and  password = '" & Password & "'  and estado = True ")

if not RSu.eof then

  odontologo = RSu("nombre")
  matricula = RSu("matricula")
  Session("allow_shopp") = True
  Session.Timeout = 600

Else
yes = "yes"
Response.Redirect "login.asp?error="&yes&""
End If

RSu.close
uConn.close

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Paciente</title>
</head>

<%  
' recepcion de paciente
' paciente tal
  paciente = request("paciente")


'set oConn =  Server.CreateObject("ADODB.Connection")
'  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
'  Set RSArt = oConn.Execute("select * from fichas  where paciente = " & paciente & "") 
'  if not RSArt.EOF then 
%>

<frameset cols="177,*" framespacing="0" border="0" frameborder="0">
  <frame name="contenido" target="principal" src="lateraldiagnost.asp?paciente=<%=paciente%>" scrolling="auto" noresize>
  <frame name="principal" src="completo.asp?paciente=<%=paciente%>">
  <noframes>
 
<%'end if
'RsArt.close
'set RsArt = nothing
'oConn.Close
'set oConn = nothing
%>
  
  <body>

  <p>Esta página usa marcos, pero su explorador no los admite.</p>

  </body>
  
  </noframes>
</frameset>

</html>