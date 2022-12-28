<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%
if Session("usuario") = "" then
response.redirect "../error_.asp"
end if
%>

<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="351">
    <tr>
      <td height="134"><img border="0" src="../../images/Sup_L.jpg"></td>
    </tr>
    <tr>
      <td height="143">
      <div align="center">
        <center>
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="333" height="211">
          <tr>
            <td width="298" height="25" align="center">&nbsp;</td>
          </tr>
          <tr>
            <td width="298" height="2" align="center" bgcolor="#2D4773"><b>
            <font color="#FFFFFF" face="Verdana" size="2"><span lang="es">
            Administración Laboratorio</span></font></b></td>
            </tr>
          <tr>
            <td width="298" height="21" align="center" bgcolor="#5F84C0"><b>
            <font color="#FFFFFF" face="Verdana" size="2"><span lang="es">Menú</span></font></b></td>
            </tr>
          <tr>
            <td width="298" height="20" align="center" bgcolor="#C2D0E7">
            <font color="#000080" face="Verdana" size="2"><span lang="es">[
            <a href="prestaciones_pub.asp"><font color="#000080">&nbsp;Precios de 
            prestaciones al público</font></a> ]</span></font></td>
            </tr>
          <tr>
            <td width="298" height="20" align="center" bgcolor="#C2D0E7">
            <font color="#000080" face="Verdana" size="2"><span lang="es">[&nbsp;&nbsp;
            <a href="prestaciones.asp"><font color="#000080">&nbsp;Precios de 
            prestaciones OSEP</font></a>&nbsp;&nbsp;&nbsp; ]</span></font></td>
            </tr>
          <tr>
            <td width="298" height="16" align="center" bgcolor="#C2D0E7">
            <font color="#000080" face="Verdana" size="2"><span lang="es">[&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <a href="profesionales.asp"><font color="#000080">Listado de 
            Profesionales</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ]</span></font></td>
            </tr>
          <tr>
            <td width="298" height="23" align="center" bgcolor="#C2D0E7">
            <font color="#000080" face="Verdana" size="2"><span lang="es">[&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; </span>
            <font color="#000080"><span lang="es"><a href="precarga.asp">
            <font color="#000080">Planillas 
            Ordenes OSEP</font></a></span></font><span lang="es">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; ]</span></font></td>
            </tr>
          <tr>
            <td width="298" height="23" align="center" bgcolor="#C2D0E7">
            <font color="#000080" face="Verdana" size="2"><span lang="es">[&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <a href="consulta.asp"><font color="#000080">Consulta de trabajos</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; 
            ]</span></font></td>
            </tr>
          <tr>
            <td width="298" height="1" align="center"></td>
            </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <span lang="es"><a href="abandon.asp"><font color="#FFFFFF">Desconectarse</font></a></span> ]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>