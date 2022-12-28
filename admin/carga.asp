<%@ Language=VBScript %>
<% Response.Buffer = True %>

<%
if Session("usuario") = "" then
response.redirect "../error_.asp"
end if

if Session("periodo") = "" then
response.redirect "precarga.asp"
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
            Carga de ordenes de OSEP</font></b></span></td>
            </tr>
          </table>
        </center>
      </div>

 
 <hr color="#2D4773" width="400" size="1">         
      
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.periodo.value == "")
  {
    alert("Escriba un valor para el campo \"periodo\".");
    theForm.periodo.focus();
    return (false);
  }

  if (theForm.mes.selectedIndex < 0)
  {
    alert("Elija una de las opciones \"mes\".");
    theForm.mes.focus();
    return (false);
  }

  if (theForm.mes.selectedIndex == 0)
  {
    alert("La primera opción \"mes\" no es válida. Elija una de las otras opciones.");
    theForm.mes.focus();
    return (false);
  }

  if (theForm.anio.value == "")
  {
    alert("Escriba un valor para el campo \"anio\".");
    theForm.anio.focus();
    return (false);
  }

  var checkOK = "0123456789-,.";
  var checkStr = theForm.anio.value;
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
    if (ch == ",")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "." && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ".")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Escriba sólo dígito caracteres en el campo \"anio\".");
    theForm.anio.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"anio\".");
    theForm.anio.focus();
    return (false);
  }

  if (theForm.cupon.value == "")
  {
    alert("Escriba un valor para el campo \"cupon\".");
    theForm.cupon.focus();
    return (false);
  }

  var checkOK = "0123456789-,.";
  var checkStr = theForm.cupon.value;
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
    if (ch == ",")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "." && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ".")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Escriba sólo dígito caracteres en el campo \"cupon\".");
    theForm.cupon.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"cupon\".");
    theForm.cupon.focus();
    return (false);
  }

  if (theForm.profesional.selectedIndex < 0)
  {
    alert("Elija una de las opciones \"profesional\".");
    theForm.profesional.focus();
    return (false);
  }

  if (theForm.profesional.selectedIndex == 0)
  {
    alert("La primera opción \"profesional\" no es válida. Elija una de las otras opciones.");
    theForm.profesional.focus();
    return (false);
  }

  if (theForm.afiliado.value == "")
  {
    alert("Escriba un valor para el campo \"afiliado\".");
    theForm.afiliado.focus();
    return (false);
  }

  if (theForm.nombre.value == "")
  {
    alert("Escriba un valor para el campo \"nombre\".");
    theForm.nombre.focus();
    return (false);
  }

  if (theForm.codigo.selectedIndex < 0)
  {
    alert("Elija una de las opciones \"codigo\".");
    theForm.codigo.focus();
    return (false);
  }

  if (theForm.codigo.selectedIndex == 0)
  {
    alert("La primera opción \"codigo\" no es válida. Elija una de las otras opciones.");
    theForm.codigo.focus();
    return (false);
  }

  if (theForm.cantidad.value == "")
  {
    alert("Escriba un valor para el campo \"cantidad\".");
    theForm.cantidad.focus();
    return (false);
  }

  var checkOK = "0123456789-,.";
  var checkStr = theForm.cantidad.value;
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
    if (ch == ",")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "." && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ".")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Escriba sólo dígito caracteres en el campo \"cantidad\".");
    theForm.cantidad.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"cantidad\".");
    theForm.cantidad.focus();
    return (false);
  }

  var checkOK = "0123456789-,";
  var checkStr = theForm.valorcupon.value;
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
    if (ch == ",")
    {
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Escriba sólo dígito caracteres en el campo \"valorcupon\".");
    theForm.valorcupon.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Escriba un número válido en el campo \"valorcupon\".");
    theForm.valorcupon.focus();
    return (false);
  }

  if (theForm.fechaosep.value == "")
  {
    alert("Escriba un valor para el campo \"fechaosep\".");
    theForm.fechaosep.focus();
    return (false);
  }

  if (theForm.mas.selectedIndex < 0)
  {
    alert("Elija una de las opciones \"mas\".");
    theForm.mas.focus();
    return (false);
  }

  if (theForm.mas.selectedIndex == 0)
  {
    alert("La primera opción \"mas\" no es válida. Elija una de las otras opciones.");
    theForm.mas.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="do_carga_osep.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
       
 <div align="center">
          <center>
          <table border="0" cellpadding="5" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="772" bgcolor="#F4F7FB">
            <tr>
              <td width="115" align="right"><b><font face="Verdana" size="2" color="#2D4773">
              <span lang="es">Período
              :</span></font></b></td>
              <td width="607" colspan="2">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="periodo" size="44" value="<%=Session("periodo")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(Si no aparece el 
              período <a href="precarga.asp"><font color="#2D4773">haga click 
              aquí</font></a>)</font></span></td>
              </tr>
            <tr>
              <td width="115" align="right"><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              Mes:</span></font></b></td>
              <td width="294"><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              <!--webbot bot="Validation" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="mes">
              <option selected>elija</option>
              <%
              mesok = Session("mes")
              if mesok <> "" then%>
              <option selected><%=mesok%></option>
              <%end if%>
              <option value="Enero">Enero</option>
              <option value="Febrero">Febrero</option>
              <option value="Marzo">Marzo</option>
              <option value="Abril">Abril</option>
              <option value="Mayo">Mayo</option>
              <option value="Junio">Junio</option>
              <option value="Julio">Julio</option>
              <option value="Agosto">Agosto</option>
              <option value="Setiembre">Setiembre</option>
              <option value="Octubre">Octubre</option>
              <option value="Noviembre">Noviembre</option>
              <option value="Diciembre">Diciembre</option>
              </select></span></font><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: Febrero)</font></span></td>
              <td width="313"><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">Año</font></span><font face="Verdana" size="2" color="#2D4773"><span lang="es">:
              </span></font>
              <!--webbot bot="Validation" s-data-type="Number" s-number-separators=".," b-value-required="TRUE" --><input type="text" name="anio" size="20" value="<%=Session("anio")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 2009)</font></span></td>
            </tr>
            <tr>
              <td width="115" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Nro. Cupón</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es">
              :</span></font></b></td>
              <td width="622" colspan="2"><span lang="es">&nbsp;</span><!--webbot bot="Validation" s-data-type="Number" s-number-separators=".," b-value-required="TRUE" --><input type="text" name="cupon" size="20" value="<%=Session("cupon")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 1023, sólo el 
              número)</font></span></td>
            </tr>
            <tr>
              <td width="115" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Profesional :</font></b></span></td>
              <%

set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")

Set RSArt = oConn.Execute("select * from profesionales") 
if not RSArt.EOF then 
%>
              <td width="622" colspan="2">
              <font color="#2D4773" size="2" face="Verdana"><span lang="es">Dr.
              <!--webbot bot="Validation" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="profesional">
              <option selected>elegir</option>
              <%
              odont = Session("profesional")
              if odont = "" then
              else%>
              <option selected><%=odont%></option>
              <%end if%>
              <%do while not RSArt.eof%>
              <option value="<%=RSArt("profesional")%>"><%=RSArt("profesional")%></option>
              <%rsart.movenext
              loop
              %>
              </select></span></font><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">(Debe elegir uno)</span></font></td>
            </tr>
  <%else
end if
RsArt.close
set RsArt = nothing
%>
            <tr>
              <td width="115" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Nro. afiliado :</font></b></span></td>
              <td width="622" colspan="2">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="afiliado" size="33" value="<%=Session("afiliado")%>"><span lang="es"> </span></td>
            </tr>
            <tr>
              <td width="115" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Nombre:</font></b></span></td>
              <td width="622" colspan="2">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="nombre" size="33" value="<%=Session("nombre")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(Apellido y nombres)</font></span></td>
            </tr>
            
            <tr>
              <td width="115" align="left">
              <p align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Código:</font></b></span></td>
              <td width="294" align="left">
 <%
Set RSArt = oConn.Execute("select * from valores order by id") 
if not RSArt.EOF then 
%>             
              
              <font color="#2D4773" size="2" face="Verdana"><span lang="es">
               <!--webbot bot="Validation" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="codigo">
              <option selected>elegir</option>

              <%do while not RSArt.eof%>
              <option value="<%=RSArt("codigo")%>"><%=RSArt("codigo")%></option>
              <%rsart.movenext
              loop%>
<%
end if
RsArt.close
set RsArt = nothing
%>
              </select></span></font><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(Debe elegir uno) </font></span></td>
              <td width="313" align="left">
               <b><span lang="es"><font face="Verdana" size="2" color="#2D4773">
               Cantidad</font></span></b><font face="Verdana" size="2" color="#2D4773"><span lang="es"><b>:</b>
              </span></font>
               <!--webbot bot="Validation" s-data-type="Number" s-number-separators=".," b-value-required="TRUE" --><input type="text" name="cantidad" size="6"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 2 )</font></span></td>
            </tr>
            <tr>
              <td width="115" align="right"><span lang="es"><b>
              <font face="Verdana" size="2" color="#2D4773">Valor del cupón:</font></b></span></td>
              <td width="607" colspan="2">
 <span lang="es"><font color="#2D4773" size="2" face="Verdana">$</font> </span>
               <!--webbot bot="Validation" S-Data-Type="Number" S-Number-Separators="x," --><input type="text" name="valorcupon" size="11" value="<%=Session("valorcupon")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(Sólo número. Ej: 
               127,55) </span></font></td>
            </tr>
            <tr>
              <td width="115" align="left"><b><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">Fecha entrega a OSEP</font></span></b><font face="Verdana" size="2" color="#2D4773"><span lang="es"><b>:</b>
              </span></font>
              </td>
              <td width="622" align="left" colspan="2">
              <!--webbot bot="Validation" b-value-required="TRUE" --><input type="text" name="fechaosep" size="12" value="<%=Session("fechaosep")%>"><span lang="es">&nbsp;
              <font face="Verdana" size="2" color="#2D4773">(ej.: 15/03/2009 )</span></font></td>
              </tr>
            <tr>
              <td width="752" align="left" colspan="3" bgcolor="#FFFFCC">
              <p align="center"><b><span lang="es">
              <font face="Verdana" size="2" color="#2D4773">Cargará más códigos 
              a este mismo paciente ?:
              <!--webbot bot="Validation" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="mas">
              <option selected>elija</option>
              <option value="Si">Si</option>
              <option value="No">No</option>
              </select></font></span></b></td>
              </tr>
            
            <tr>
  <%
oConn.Close
set oConn = nothing%>             
              
            </tr>
          </table>
          </center>
        </div>
                <p align="center">

        <input type="submit" value="Cargar datos" name="B2"></p>
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