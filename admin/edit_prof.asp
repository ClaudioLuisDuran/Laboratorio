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
Set RSArt = oConn.Execute("select * from profesionales where Id = " & id_ & "") 

if not RSArt.EOF then 
%>
             
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ Edición de datos de odontólogo ]</span></font></b></p>
             
      
      <p align="center"><span lang="es">
      <font face="Verdana" size="2" color="#2D4773">Cambiar solo lo necesario y 
      no tocar el resto</font></span></p>
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

  var checkOK = "0123456789-.";
  var checkStr = theForm.visitas.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Escriba sólo dígito caracteres en el campo \"visitas\".");
    theForm.visitas.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"visitas\".");
    theForm.visitas.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="doedit_prof.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639">
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Matrícula</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              :</span></font></b></td>
              <td width="193">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="matricula" size="20" value="<%=RSArt("matricula")%>"><span lang="es">&nbsp;</span></td>
              <td width="96"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Fecha alta: </font>
              </b></span></td>
              <td width="195">
              <span lang="es"><font face="Verdana" size="2" color="#2D4773">
              <%=RSArt("fechaalta")%></font></span></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Nombre :</font></b></span></td>
              <td width="542" colspan="3">
              <span lang="es"><font face="Verdana" size="2" color="#2D4773">Dr</font></span><font face="Verdana" size="2" color="#2D4773"><span lang="es">.
              </span></font>
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="profesional" size="39" value="<%=RSArt("profesional")%>"></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Email :</font></b></span></td>
              <td width="542" colspan="3">
              <input type="text" name="email" size="43" value="<%=RSArt("email")%>"></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Usuario :</font></b></span></td>
              <td width="182">
              <input type="text" name="usuario" size="22" value="<%=RSArt("usuario")%>"></td>
              <td width="95"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Password :</font></b></span></td>
              <td width="190">
              <input type="text" name="password" size="22" value="<%=RSArt("password")%>"></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Visitas :</font></b></span></td>
              <td width="191">
              <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." --><input type="text" name="visitas" size="18" value="<%=RSArt("visitas")%>"></td>
              <td width="97"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Ultima visita:
              </font></b></span></td>
              <td width="195">
              <span lang="es"><font face="Verdana" size="2" color="#2D4773">
<%=RSArt("ultvis")%></font></span></td>
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