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
            Administraci�n Laboratorio</span></font></b></td>
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

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from valores order by id") 

if not RSArt.EOF then 
%>
      <div align="center">
        <center>
        <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="699">
          <tr>
            <td width="96" align="center">&nbsp;</td>
            <td width="508" align="center">&nbsp;</td>
            <td width="100" align="center">&nbsp;</td>
            <td width="31" align="center">&nbsp;</td>
            <td width="31" align="center">&nbsp;</td>
          </tr>
          <tr>
            <td width="96" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">C�digo</font></span></b></td>
            <td width="508" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Descripci�n</font></span></b></td>
            <td width="100" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Valor</font></span></b></td>
            <td width="31" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Editar</font></span></b></td>
            <td width="31" align="center" bgcolor="#819FCD"><b><span lang="es">
            <font face="Verdana" size="2" color="#FFFFFF">Borrar</font></span></b></td>
          </tr>
          
 <%
 paso = 1
do while not rsart.eof
if paso = 1 then
%>
          <tr>
            <td width="96" align="center"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("codigo")%></font></span></td>
            <td width="508" align="center">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("descripcion")%></font></span></td>
            <td width="100" align="center">
            <p align="left"><span lang="es">&nbsp;<font face="Verdana" size="2" color="#2D4773">$ 
            <%=RSArt("valor")%></font></span></td>
            <td align="center"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_pres.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="borra_pres.asp?id=<%=RSArt("id")%>" onclick="return confirm('Si lo borras no podr�s recuperarlo. Est�s seguro?')><font color="#2D4773">
            <font color="#2D4773">Borrar</font></a></font></font></span></td>
          </tr>
 <%paso = 2
 else%>         
          <tr>
            <td width="96" align="center" bgcolor="#DBE3F0"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("codigo")%></font></span></td>
            <td width="508" align="center" bgcolor="#DBE3F0">
            <p align="left"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773"><%=RSArt("descripcion")%></font></span></td>
            <td width="100" align="center" bgcolor="#DBE3F0">
            <p align="left"><span lang="es">&nbsp;<font face="Verdana" size="2" color="#2D4773">$ 
            <%=RSArt("valor")%></font></span></td>
            <td align="center" bgcolor="#DBE3F0"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="edit_pres.asp?id=<%=RSArt("id")%>"><font color="#2D4773">Editar</font></a></font></span></td>
            <td align="center" bgcolor="#DBE3F0"><span lang="es">
            <font face="Verdana" size="2" color="#2D4773">
            <a href="borra_pres.asp?id=<%=RSArt("id")%>" onclick="return confirm('Si lo borras no podr�s recuperarlo. Est�s seguro?')><font color="#2D4773">
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
      <span lang="es">[ Agregar nueva prestaci�n ]</span></font></b></p>
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
    alert("Escriba s�lo d�gito caracteres en el campo \"valor\".");
    theForm.valor.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un n�mero v�lido en el campo \"valor\".");
    theForm.valor.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="carga_pres.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="639" bgcolor="#F0F3F9">
            <tr>
              <td width="97" align="right"><b>
            <font face="Verdana" size="2" color="#2D4773"><span lang="es">
              C�digo :</span></font></b></td>
              <td width="542"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">Nro</font></span><font color="#2D4773" size="2" face="Verdana"><span lang="es">
              </span></font><span lang="es">. </span>
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="codigo" size="20"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 40406)</font></span></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Descripci�n :</font></b></span></td>
              <td width="542">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="descripcion" size="60"></td>
            </tr>
            <tr>
              <td width="97" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Valor :</font></b></span></td>
              <td width="542"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">$ </font></span>
              <font color="#2D4773" size="2" face="Verdana"><span lang="es">&nbsp;</span></font><!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." b-value-required="TRUE" --><input type="text" name="valor" size="20"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(si es necesario, 
              solo 2 decimales: ej. <b>35.50</b></font></span><font face="Verdana" size="2" color="#2D4773"><span lang="es">)</span></font></td>
            </tr>
          </table>
          </center>
        </div>
        <p align="center"><input type="submit" value="Cargar" name="B1"></p>
      </form>
      <hr color="#2D4773" width="400" size="1">
      <p align="center"><b><font face="Verdana" size="2" color="#FFFFFF">
      <font color="#2D4773"><span lang="es">[ </span></font><span lang="es">
      <a href="index_adm.asp"><font color="#2D4773">Volver al Men� principal</font></a></span></font><font color="#2D4773" face="Verdana" size="2"><span lang="es"> 
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