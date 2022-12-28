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
            <span lang="es"><b><font face="Verdana" size="2" color="#FFFFFF">
            Listado de Prestaciones</font></b></span></td>
            </tr>
          </table>
        </center>
      </div>
<%
id_ = Request("id")

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from valores where Id = " & id_ & "") 

if not RSArt.EOF then 
%>
             
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[ Edición de datos ]</span></font></b></p>
             
      
      <p align="center"><span lang="es">
      <font face="Verdana" size="2" color="#2D4773">Cambiar solo lo necesario y 
      no tocar el resto</font></span></p>
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.codigo.value == "")
  {
    alert("Escriba un valor para el campo \"codigo\".");
    theForm.codigo.focus();
    return (false);
  }

  if (theForm.descripcion.value == "")
  {
    alert("Escriba un valor para el campo \"descripcion\".");
    theForm.descripcion.focus();
    return (false);
  }

  if (theForm.valor.value == "")
  {
    alert("Escriba un valor para el campo \"valor\".");
    theForm.valor.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.valor.value;
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
    alert("Escriba sólo dígito caracteres en el campo \"valor\".");
    theForm.valor.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"valor\".");
    theForm.valor.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="doedit_pres.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639">
            <tr>
              <td width="97" align="right"><b>
            <font face="Verdana" size="2" color="#2D4773"><span lang="es">
              Código :</span></font></b></td>
              <td width="542"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">Nro</font></span><font color="#2D4773" size="2" face="Verdana"><span lang="es">
              </span></font><span lang="es">. </span>
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="codigo" size="20" value="<%=RSArt("codigo")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 40406)</font></span></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Descripción :</font></b></span></td>
              <td width="542">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="descripcion" size="60" value="<%=RSArt("descripcion")%>"></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Valor :</font></b></span></td>
              <td width="542"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">$ </font></span>
              <font color="#2D4773" size="2" face="Verdana"><span lang="es">&nbsp;</span></font><!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." b-value-required="TRUE" --><input type="text" name="valor" size="20" value="<%=RSArt("valor")%>"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(si es necesario, 
              solo 2 decimales: ej. <b>35.50</span></b><span lang="es"><span lang="es">)</span></span></font></td>
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
        
        <p align="center"><input type="submit" value="Actualizar" name="B1"></p>
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