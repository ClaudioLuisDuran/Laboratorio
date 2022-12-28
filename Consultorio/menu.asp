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
<title>Consultorio de <%=odontologo%></title>
</head>

<body>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="490">
    <tr>
      <td height="134"><img border="0" src="../images/Sup_C.jpg"></td>
    </tr>
    <tr>
      <td height="23">
      <p align="center">&nbsp;</td>
    </tr>
    <tr>
      <td height="282">
      <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="677">
          <tr>
            <td width="188">
            <img border="0" src="images/sillon.jpg" width="218" height="182"></td>
            <td width="454">
            <p align="center"><font color="#2D4773" face="Verdana" size="2">[
            <span lang="es">Menú de consultorio de <%=odontologo%></span>] </font>
            </p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080">Buscar un paciente</font></span></p>
            <p align="center"><span lang="es"><font face="Verdana" size="2">
            <a href="listado.asp">
            <font color="#000080">Listado de pacientes</font></a></font></span></p>
            <p align="center"><span lang="es"><font face="Verdana" size="2">
            <a href="newpac.asp?consulta=TRUE">
            <font color="#000080">Cargar nuevo paciente (en consulta)</font></a></font></span></p>
            <p align="center"><span lang="es"><font face="Verdana" size="2">
            <a href="newpac.asp">
            <font color="#000080">Cargar nuevo paciente (de archivo)</font></a></font></span></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080"><a href="abandon.asp">
            <font color="#2D4773">Desconectarse</font></a></font></span></p>
            <p>&nbsp;</td>
          </tr>
          <tr>
            <td width="642" colspan="2">
            &nbsp;</td>
          </tr>
          </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      &nbsp;</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>