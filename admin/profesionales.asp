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

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from profesionales order by id") 

if not RSArt.EOF then 
%>
      <div align="center">
        <center>
        <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="699">
          <tr>
            <td width="89" align="center">&nbsp;</td>
            <td width="285" align="center">&nbsp;</td>
            <td width="214" align="center">&nbsp;</td>
            <td width="44" align="center">&nbsp;</td>
            <td width="47" align="center">&nbsp;</td>
          </tr>
          <tr>
            <td width="89" align="center" bgcolor="#819FCD"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Matricula</font></b></span></td>
            <td width="285" align="center" bgcolor="#819FCD"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Nombre</font></b></span></td>
            <td width="214" align="center" bgcolor="#819FCD"><span lang="es"><b>
            <font face="Verdana" size="2" color="#FFFFFF">Email</font></b></span></td>
            <td width="44" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Editar</font></span></b></td>
            <td width="47" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Borrar</font></span></b></td>
          </tr>
          
 <%
 paso = 1
do while not rsart.eof
if paso = 1 then
%>
          <tr>
            <td width="89" align="center"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("matricula")%></font></span></td>
            <td width="285" align="center">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("profesional")%></font></span></td>
            <td width="214" align="center">
            <p align="left"><span lang="es">&nbsp;<font face="Verdana" size="2" color="#2D4773"> 
            <%=RSArt("email")%></font></span></td>
            <td align="center" width="44"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_prof.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center" width="47"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a onclick="return confirm('Si lo borras no podrás recuperarlo. Estás seguro?')&gt;&lt;font color=" #2D4773" href="borra_prof.asp?id=<%=RSArt("id")%>">
            <font color="#2D4773">Borrar</font></a></font></font></span></td>
          </tr>
 <%paso = 2
 else%>         
          <tr>
            <td width="89" align="center" bgcolor="#DBE3F0"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("matricula")%></font></span></td>
            <td width="285" align="center" bgcolor="#DBE3F0">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("profesional")%></font></span></td>
            <td width="214" align="center" bgcolor="#DBE3F0">
            <p align="left"><span lang="es">&nbsp;<font face="Verdana" size="2" color="#2D4773"> 
            <%=RSArt("email")%></font></span></td>
            <td align="center" bgcolor="#DBE3F0" width="44"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_prof.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center" bgcolor="#DBE3F0" width="47"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a onclick="return confirm('Si lo borras no podrás recuperarlo. Estás seguro?')&gt;&lt;font color=" #2D4773" href="borra_prof.asp?id=<%=RSArt("id")%>">
            <font color="#2D4773">Borrar</font></a></font></font></span></td>
          </tr>
  <%paso=1
  end if
  rsart.movenext
loop%>         
          
        </table>
        </center>
      </div>
  <%else
end if
RsArt.close
set RsArt = nothing
oConn.Close
set oConn = nothing%> <hr color="#2D4773" width="400" size="1">         
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ Agregar nuevo Profesional ]</span></font></b></p>
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.matricula.value == "")
  {
    alert("Escriba un valor para el campo \"matricula\".");
    theForm.matricula.focus();
    return (false);
  }

  if (theForm.profesional.value == "")
  {
    alert("Escriba un valor para el campo \"profesional\".");
    theForm.profesional.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="carga_prof.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639" bgcolor="#F4F7FB">
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Matricula</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              :</span></font></b></td>
              <td width="542"><span lang="es">&nbsp;</span><font face="Verdana" size="2" color="#2D4773"><span lang="es">Mat. 
              nro. </span></font><span lang="es">&nbsp;</span><!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="matricula" size="20"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 1023, sólo el 
              número)</font></span></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Nombre :</font></b></span></td>
              <td width="542">
              <font color="#2D4773" size="2" face="Verdana"><span lang="es">Dr.
              </span></font>
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="profesional" size="37"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(sólo el nombre y 
              apellido)</font></span></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Email :</font></b></span></td>
              <td width="542"><input type="text" name="email" size="41"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(si no lo tiene deje 
              en blanco)</font></span></td>
            </tr>
          </table>
          </center>
        </div>
        <p align="center">
        <input type="submit" value="Cargar nuevo odontólogo" name="B1"></p>
      </form>
      <hr color="#2D4773" width="400" size="1">
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