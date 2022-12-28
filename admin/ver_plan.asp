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
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1035" height="351">
    <tr>
      <td height="134" width="1145">
      <p align="center"><img border="0" src="../../images/Sup_L.jpg"></td>
    </tr>
    <tr>
      <td height="143" width="1145">
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
            Planillas de Ordenes OSEP</font></b></span></td>
            </tr>
          </table>
        </center>
      </div>
<%
periodo_ = Request("periodo")
set oConn =  Server.CreateObject("ADODB.Connection")

oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../../db/laboratorio.mdb")
Set RSArt = oConn.Execute("select * from prestaciones where periodo LIKE '%" & periodo_ & "%';") 

if not RSArt.EOF then 

periodo_ok = RSArt("periodo")
%>
      <div align="center">
        <center>
        <table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" width="1145" height="130">
          <tr>
            <td width="1136" align="center" colspan="12" height="49">
            <span lang="es"><b><font face="Verdana" size="2" color="#2D4773">
            Período : </font><font face="Verdana" color="#2D4773"><%=RSArt("periodo")%></font></b></span></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#819FCD" width="54" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">Nro. 
            cupón</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="98" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Profesional</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="83" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">Nro. 
            afiliado</font></b></span></td>
            <td align="left" bgcolor="#819FCD" width="149" height="26">
            <p align="center">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Apellido y nombre afiliado</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="72" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Código prestación</font></b></span></td>
            <td align="left" bgcolor="#819FCD" width="194" height="26">
            <p align="center">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Descripción</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="55" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            cantidad</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="74" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Precio unitario</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="58" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Subtotal</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="68" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Total</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="79" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Valor cupón</font></b></span></td>
            <td align="center" bgcolor="#819FCD" width="113" height="26">
            <span lang="es"><b><font face="Verdana" size="1" color="#FFFFFF">
            Fecha entrega a OSEP</font></b></span></td>
          </tr>
          
 <% 
 cupon_ant = 0
 'paso = 1
 total = 0
 color = 1
 publi = 1
 
do while not rsart.eof

cupon_new = RSArt("cupon")

if cupon_ant = 0 then
cupon_ant = cupon_new
end if

if cupon_new = cupon_ant then
paso = color
else
paso = paso
publi = 1
end if

if paso = 1 then
%>
          <tr>
            <td align="center" width="54" height="19"><span lang="es">
            <font size="1" face="Verdana">
            <a href="edit_carga.asp?id=<%=RSArt("id")%>"><%=RSArt("cupon")%></a></font></span></td>
            <td align="center" width="98" height="19"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("profesional")%></font></span></td>
            <td align="center" width="83" height="19">
            <span lang="es"><font face="Verdana" size="1"><%=RSArt("afiliado")%></font></span></td>
            <td align="center" width="149" height="19">
            <span lang="es"><font face="Verdana" size="1"><%=RSArt("nombre")%></font></span></td>
            <td align="center" width="72" height="19"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("codigo")%></font></span></td>
            <td align="left" width="194" height="19"><span lang="es">

            <font face="Verdana" size="1"><%=RSArt("descripcion")%></font></span></td>
            
            <td align="center" width="55" height="19"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("cantidad")%></font></span></td>
            <td align="center" width="74" height="19"><span lang="es">
            <% descrip_ =RSArt("descripcion")
            Set RS = oConn.Execute("select * from valores where descripcion LIKE '%" & descrip_ & "%';") 

			if not RS.EOF then %>
            <font face="Verdana" size="1">$ <%=RS("valor")%></font></span>
            <%valor_=RS("valor")
            cant_= RSART("cantidad")
            subtotal_ = valor_ * cant_
            %>
            </td>
            <td align="center" width="58" height="19"><span lang="es">
            
            <font face="Verdana" size="1">$ <%=subtotal_%></font></span>
            <%total = total + subtotal_%>
            </td>
            <td align="center" width="68" height="19"><span lang="es">
            
            <font face="Verdana" size="1">$ <%=total%></font></span>
            <%else
            end if
            Rs.close
            set Rs = nothing
            %>
            </td>
            <td align="center" width="79" height="19"><span lang="es">
            
            <font face="Verdana" size="1">
            <%if publi = 1 then%>
            $ <%=RSArt("valorcupon")%>
            <%end if%>
            
            </font></span></td>
            <td align="center" width="113" height="19"><span lang="es">
            <font face="Verdana" size="1">
            <%if publi = 1 then%>
            <%=RSArt("fechaosep")%>
             <%end if%></font></span></td>
          </tr>
 <% color = 1
 publi = 0
 paso = 2
 cupon_ant = RSArt("cupon")
 'paso = 2
 else%>         
           <tr>
            <td align="center" width="54" height="19" bgcolor="#D1DCED"><span lang="es">
            <font size="1" face="Verdana"><a href="edit_carga.asp?id=<%=RSArt("id")%>"><%=RSArt("cupon")%></a></font></span></td>
            <td align="center" width="98" height="19" bgcolor="#D1DCED"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("profesional")%></font></span></td>
            <td align="center" width="83" height="19" bgcolor="#D1DCED">
            <span lang="es"><font face="Verdana" size="1"><%=RSArt("afiliado")%></font></span></td>
            <td align="center" width="149" height="19" bgcolor="#D1DCED">
            <span lang="es"><font face="Verdana" size="1"><%=RSArt("nombre")%></font></span></td>
            <td align="center" width="72" height="19" bgcolor="#D1DCED"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("codigo")%></font></span></td>
            <td align="left" width="194" height="19" bgcolor="#D1DCED"><span lang="es">

            <font face="Verdana" size="1"><%=RSArt("descripcion")%></font></span></td>
            
            <td align="center" width="55" height="19" bgcolor="#D1DCED"><span lang="es">
            <font face="Verdana" size="1"><%=RSArt("cantidad")%></font></span></td>
            <td align="center" width="74" height="19" bgcolor="#D1DCED"><span lang="es">
            <% descrip_ = RSArt("descripcion")
            Set RS = oConn.Execute("select * from valores where descripcion LIKE '%" & descrip_ & "%';") 

			if not RS.EOF then %>
            <font face="Verdana" size="1">$ <%=RS("valor")%></font></span>
            <%valor_= RS("valor")
            cant_= RSART("cantidad")
            subtotal_ = valor_ * cant_
            %>
            </td>
            <td align="center" width="58" height="19" bgcolor="#D1DCED"><span lang="es">
            
            <font face="Verdana" size="1">$ <%=subtotal_%></font></span>
            <%total = total + subtotal_%>
            </td>
            <td align="center" width="68" height="19" bgcolor="#D1DCED"><span lang="es">
            
            <font face="Verdana" size="1">$ <%=total%></font></span>
            <%else
            end if
            Rs.close
            set Rs = nothing
            %>
            </td>
            <td align="center" width="79" height="19" bgcolor="#D1DCED"><span lang="es">
            <font face="Verdana" size="1">
            <%if publi = 1 then%>
            $ <%=RSArt("valorcupon")%>
            <%end if%>
            </font></span></td>
            <td align="center" width="113" height="19" bgcolor="#D1DCED"><span lang="es">
            <font face="Verdana" size="1">
            <%if publi = 1 then%>
            <%=RSArt("fechaosep")%>
             <%end if%>
             </font></span></td>
          </tr>  
 <%paso=1
 color = 2
 publi = 0
 cupon_ant = RSArt("cupon")
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
set oConn = nothing%> <hr color="#2D4773" width="333" size="1" align="right">
      <div align="right">
        <table border="1" cellpadding="10" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="300">
          <tr>
            <td>
            <p align="center"><span lang="es"><b>
            <font face="Verdana" size="2" color="#2D4773">Total</font></b></span><b><font face="Verdana" size="2" color="#2D4773"><span lang="es"> 
            : $ <%=total%></span> </font></b></td>
          </tr>
        </table>
      </div>
      <hr color="#2D4773" width="400" size="1">         
      
      <p align="center"><b><font color="#2D4773" size="2" face="Verdana">
      <span lang="es">[&nbsp;
      <a href="do_pre_carga_ok.asp?periodo=<%=periodo_ok%>">
      <font color="#2D4773">Continuar cargando ordenes en esta planilla </font>
      </a>]</span></font></b></p>
      <hr color="#2D4773" width="400" size="1">         
      
      <p align="center"><b><font face="Verdana" size="2" color="#FFFFFF">
      <font color="#2D4773"><span lang="es">[ </span></font><span lang="es">
      <a href="index_adm.asp"><font color="#2D4773">Volver al Menú principal</font></a></span></font><font color="#2D4773" face="Verdana" size="2"><span lang="es"> 
      ]</span></font></b></p>
      <p align="center">&nbsp;</td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773" width="1145">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <span lang="es"><a href="abandon.asp"><font color="#FFFFFF">Desconectarse</font></a></span> ]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>