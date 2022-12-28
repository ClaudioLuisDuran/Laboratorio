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
        <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="333" height="48">
          <tr>
            <td width="298" height="2" align="center" bgcolor="#2D4773"><b>
            <font color="#FFFFFF" face="Verdana" size="2"><span lang="es">
            Administración Laboratorio</span></font></b></td>
            </tr>
          <tr>
            <td width="298" height="1" align="center" bgcolor="#5F84C0">
            <b><font face="Verdana" size="2" color="#FFFFFF">
            <span lang="es">Listado de Profesionales</span></font></b></td>
            </tr>
          </table>
        </center>
      </div>
<%
id_ = Request("id")

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from planillas where Id = " & id_ & "") 

if not RSArt.EOF then 
%>
             
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ Edición de datos de odontólogo ]</span></font></b></p>
             
      
      <p align="center"><span lang="es">
      <font face="Verdana" size="2" color="#2D4773">Cambiar solo lo necesario y 
      no tocar el resto</font></span></p>
      <form method="POST" action="doedit_plan.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="649">
            <tr>
              <td width="131" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Período :</font></b></span></td>
              <td width="518" colspan="3">
              <span lang="es"><font face="Verdana" size="2" color="#2D4773"><%=RSArt("periodo")%></font></span><font face="Verdana" size="2" color="#2D4773"><span lang="es">.</span></font></td>
            </tr>
            <tr>
              <td width="131" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Fecha presentación :</font></b></span></td>
              <td width="518" colspan="3">
              <input type="text" name="fecha" size="28" value="<%=RSArt("fecha")%>"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(ej.: 15/02/2009)</font></span></td>
            </tr>
            <tr>
              <td width="131" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Pagada ? :</font></b></span></td>
              <td width="158">
              <input type="text" name="pagada" size="22" value="<%=RSArt("pagada")%>"></td>
              <td width="95"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(Si ó No)</font></span></td>
              <td width="190">
              &nbsp;</td>
            </tr>
            </table>
          </center>
        </div>
        
<%else
end if
RsArt.close
set RsArt = nothing
oConn.Close
set oConn = nothing%>         
        
        <p align="center">
        <input type="submit" value="Actualizar datos" name="B1"></p>
        <input type="hidden" name="id" value="<%=Request("id")%>">
      </form>
      <p align="center"><b><font face="Verdana" size="2" color="#FFFFFF">
      <font color="#2D4773"><span lang="es">[ </span></font><span lang="es">
      <a href="javascript:history.back()"><font color="#2D4773">Volver a</font></a></span></font><font color="#2D4773" face="Verdana" size="2"><span lang="es"><a href="javascript:history.back()"><font color="#2D4773"> 
      la página anterior sin efectuar cambios</font></a> 
      ]</span></font></b></p>
      <p align="center"><b><font face="Verdana" size="2" color="#FFFFFF">
      <font color="#2D4773"><span lang="es">[ </span></font><span lang="es">
      <a href="index_adm.asp"><font color="#2D4773">Volver al Menú principal</font></a></span></font><font color="#2D4773" face="Verdana" size="2"><span lang="es"> 
      ]</span></font></b></p>
      <p align="center">&nbsp;</td>
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