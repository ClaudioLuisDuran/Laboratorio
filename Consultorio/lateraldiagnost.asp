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
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Ficha Nº xxx</title>
<base target="principal">
</head>

<body bgcolor="#FFFFE8">
<%  
' recepcion de paciente
' paciente tal
  paciente = request("paciente")


set oConn =  Server.CreateObject("ADODB.Connection")
  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
  Set RSArt = oConn.Execute("select * from fichas where paciente = " & paciente & "") 
  if not RSArt.EOF then 
%>

<table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="339">
  <tr>
    <td width="100%" height="19">
    <p align="center"><u><font size="1" face="Verdana"><b>Ficha Nº </b>
    <font color="#000080"><b><%=paciente%></b></font></font></u></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><b>
    <font size="2" face="Verdana" color="#000080"><%=RSArt("nombre")%></font><font size="2" face="Verdana" color="#000080">&nbsp;<%=RSArt("apellido")%></font></b></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><b>
    <font size="1" face="Verdana" color="#000080"><%'=RSArt("apellido")%></font></b></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><font size="1" face="Verdana">
    <%=RSArt("obrasocial")%></font></td>
  </tr>
  <tr>
    <td width="100%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" height="19">
    <p align="center"><b><font size="1" face="Verdana" color="#000080">Opciones</font></b></td>
  </tr>
  <tr>
    <td width="100%" height="1" align="center">
    <font color="#000080" face="Verdana" size="1">
    <a href="odontograma.asp?paciente=<%=paciente%>"><font color="#000080">Ver 
    odontograma </font><span lang="es"><font color="#000080">adulto</font></span></a></font><hr color="#000080" size="1">
    </td>
  </tr>
  <tr>
    <td width="100%" height="1" align="center"><font face="Verdana" size="1">
    <a href="temporario.asp?paciente=<%=paciente%>"><font color="#000080">Ver 
    odontograma de elementos temporarios</font></a></font><hr color="#000080" size="1">
    </td>
  </tr>
  <tr>
    <td width="100%" height="3" align="center"><font size="1" face="Verdana">
    <a href="completo.asp?paciente=<%=paciente%>"><font color="#000080">Ver 
    juntos ambos Odontogramas</font></a></font><hr color="#000080" size="1"></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><span lang="es">
    <font face="Verdana" size="1">
    <a target="_top" href="verficha.asp?paciente=<%=paciente%>">
    <font color="#000080">Ver ficha completa del paciente</font></a></font></span><hr color="#000080" size="1"></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><span lang="es">
    <a target="_top" href="listado.asp"><font face="Verdana" size="1" color="#000080">Listado 
    completo de pacientes</font></a></span><hr color="#000080" size="1"></td>
  </tr>
  <tr>
    <td width="100%" height="19" align="center"><span lang="es">
    <a target="_top" href="menu.asp"><font face="Verdana" size="1" color="#000080">Volver al 
    Menú principal</font></a></span><hr color="#000080" size="1"></td>
  </tr>
</table>

<%end if
RsArt.close
set RsArt = nothing
oConn.Close
set oConn = nothing
%>

</body>

</html>