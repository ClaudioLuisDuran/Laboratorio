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
<title>Pagina nueva 1</title>

<script type="text/javascript" language="javascript">
function is_loaded() { //DOM
if (document.getElementById){
document.getElementById('preloader').style.visibility='hidden';
}else{
if (document.layers){ //NS4
document.preloader.visibility = 'hidden';
}
else { //IE4
document.all.preloader.style.visibility = 'hidden';
}
}
}
//-->
</script> 


</head>

<body onload="is_loaded();" bgcolor="#FFFFE8">

<%  
' recepcion de paciente
' paciente tal
  paciente = request("paciente")
  paciente = cint(paciente)
'response.write paciente


set oConn =  Server.CreateObject("ADODB.Connection")
  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
  Set RSpac = oConn.Execute("select * from fichas where paciente = " & paciente & "")
  if not RSpac.EOF then
  nombre = RSpac("nombre")
  apellido = RSpac("apellido")
  fecha = RSpac("fecha_consulta")

Dia = Day(fecha)
Mes = Month(fecha)
Anio = Year(fecha)

fecha = Dia &"/"& Mes &"/"& Anio
%>



<div id="preloader" style="position:absolute; left:300; top:290; width:475; height:57">
<center>
<div align="center">
  <center>
<table border="0" cellpadding="6" cellspacing="0" style="font-family:Arial, Verdana; border: 2px solid #B02A38;; border-collapse:collapse" width="460">
<tr>
<td style="text-align:center; background-color:#B02A38" width="14">
<font style="font-weight:bold; color:#FFCE6F" size="6">!</font><font color="#FFCE6F" size="6">
</font>
</td>
<td style="text-align:center; background-color:#FED072" width="418">
<span lang="es"><b><font face="Georgia" size="2" color="#B02A38">Generando 
odontograma completo de <%=nombre%> <%=apellido%></font></b></span><b><font face="Georgia"><font style="color:#B02A38; text-align:center" size="2"> ...</font><font color="#B02A38" size="2">
</font></font></b>
</td>
</tr>
</table>
  </center>
</div>
</center>
</div> 

<%end if
RSpac.close
set RSpac = nothing
oConn.Close
set oConn = nothing
%>

<%' comienzo con el odontograma adulto
  
  elemento = request("elemento")
  if elemento = "" then
  elemento = 18
  end if
 if (elemento > 10) and (elemento < 19) then
 sector = 1
 end if
 
  if (elemento > 20) and (elemento < 29) then
 sector = 2
 end if
 
  if (elemento > 30) and (elemento < 39) then
 sector = 4
 end if
 
  if (elemento > 40) and (elemento < 49) then
 sector = 3
 end if
 
 'conectamos a odontograma del paciente recibido
 
  set oConn =  Server.CreateObject("ADODB.Connection")
  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")
  
  ' dibujar odontograma actual
  
%>
<p align="center"><font face="Verdana" size="2"><span lang="es">Odontograma 
completo [Orden fecha : <%=fecha%>]</span></font></p>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="718" height="11" bgcolor="#FFF4CC">
    <tr>
      <td width="320" height="40" align="center" colspan="8"><span lang="es"><b>
      <font face="Verdana">Derecha</font></b></span></td>
      <td height="40" align="center" width="3">
      </td>
      <td width="320" height="40" align="center" colspan="8"><span lang="es"><b>
      <font face="Verdana">Izquierda</font></b></span></td>
    </tr>
    <tr>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=18&paciente=<%=paciente%>"><font color="#111111">18</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=17&paciente=<%=paciente%>"><font color="#111111">17</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=16&paciente=<%=paciente%>"><font color="#111111">16</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=15&paciente=<%=paciente%>"><font color="#111111">15</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=14&paciente=<%=paciente%>"><font color="#111111">14</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=13&paciente=<%=paciente%>"><font color="#111111">13</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=12&paciente=<%=paciente%>"><font color="#111111">12</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=11&paciente=<%=paciente%>"><font color="#111111">11</font></a></font></b></td>
      <td height="40" align="center" width="3" background="images/linea-vertical.jpg" bgcolor="#000000">
      <p>&nbsp;</td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=21&paciente=<%=paciente%>"><font color="#111111">21</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=22&paciente=<%=paciente%>"><font color="#111111">22</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=23&paciente=<%=paciente%>"><font color="#111111">23</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=24&paciente=<%=paciente%>"><font color="#111111">24</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=25&paciente=<%=paciente%>"><font color="#111111">25</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=26&paciente=<%=paciente%>"><font color="#111111">26</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=27&paciente=<%=paciente%>"><font color="#111111">27</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=28&paciente=<%=paciente%>"><font color="#111111">28</font></a></font></b></td>
    </tr>
    <tr>
      
     
   <%' superior izquierdo adulto
    diente = 18
    do while diente > 10 %>
     <td width="40" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
      <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      
      else
     
      extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela si no hay corona
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            end if
             Rs.close
				set Rs = nothing
			
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
             
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente - 1
      loop%>
        
      
          
      <td height="40" align="center" width="3" background="images/linea-vertical.jpg" bgcolor="#000000">
      &nbsp;</td>
      
      <%' superior derecho adulto
    diente = 21
    do while diente < 29 %>
     <td width="40" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
           <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
             if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
                        end if
             Rs.close
				set Rs = nothing
		
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
           
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente + 1
      loop%>
      
      
    </tr>
    <tr>
      <td colspan="17" align="center" height="3">
      <hr noshade color="#000000" size="3"></td>
    </tr>
    <tr>
    
    
      <%' inferior izquierdo adulto
    diente = 48
    do while diente > 40 %>
     <td width="40" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
            <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
             end if
             Rs.close
				set Rs = nothing
	
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
 
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente - 1
      loop%>

      
      
      <td height="40" align="center" width="3" background="images/linea-vertical.jpg" bgcolor="#000000">
      <p align="center">&nbsp;</td>
      
      
<%' inferior derecho adulto
    diente = 31
    do while diente < 39 %>
     <td width="40" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
           <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
      
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
                         end if
             Rs.close
				set Rs = nothing
			
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
         
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente + 1
      loop%>
   



    </tr>
    <tr>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=48&paciente=<%=paciente%>"><font color="#111111">48</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=47&paciente=<%=paciente%>"><font color="#111111">47</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=46&paciente=<%=paciente%>"><font color="#111111">46</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=45&paciente=<%=paciente%>"><font color="#111111">45</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=44&paciente=<%=paciente%>"><font color="#111111">44</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=43&paciente=<%=paciente%>"><font color="#111111">43</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=42&paciente=<%=paciente%>"><font color="#111111">42</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=41&paciente=<%=paciente%>"><font color="#111111">41</font></a></font></b></td>
      <td height="40" align="center" width="3" background="images/linea-vertical.jpg" bgcolor="#000000">
      &nbsp;</td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=31&paciente=<%=paciente%>"><font color="#111111">31</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=32&paciente=<%=paciente%>"><font color="#111111">32</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=33&paciente=<%=paciente%>"><font color="#111111">33</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=34&paciente=<%=paciente%>"><font color="#111111">34</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=35&paciente=<%=paciente%>"><font color="#111111">35</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=36&paciente=<%=paciente%>"><font color="#111111">36</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=37&paciente=<%=paciente%>"><font color="#111111">37</font></a></font></b></td>
      <td width="40" height="40" align="center"><b><font face="Verdana">
      <a href="odontograma.asp?elemento=38&paciente=<%=paciente%>"><font color="#111111">38</font></a></font></b></td>
    </tr>
  </table>
  </center>
</div>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111" width="446" height="11" bgcolor="#CCFFCC">
    <tr>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=55&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">55</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=54&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">54</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=53&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">53</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=52&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">52</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=51&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">51</font></a></b></span></td>
      <td height="40" align="center" width="4" background="images/linea-vertical.jpg" bgcolor="#000000">
      <p>&nbsp;</td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=61&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">61</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=62&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">62</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=63&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">63</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=64&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">64</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=65&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">65</font></a></b></span></td>
    </tr>
    <tr>
      
     
   <%' superior izquierdo temporario
    diente = 55
    do while diente > 50 %>
     <td width="65" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
      <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      
      else
     
      extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela si no hay corona
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma2 where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            end if
             Rs.close
				set Rs = nothing
			
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
             
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente - 1
      loop%>
        
      
          
      <td height="40" align="center" width="65" background="images/linea-vertical.jpg" bgcolor="#000000">
      &nbsp;</td>
      
      <%' superior derecho temporario
    diente = 61
    do while diente < 66 %>
     <td width="65" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
           <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
             if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma2 where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
                        end if
             Rs.close
				set Rs = nothing
		
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
           
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente + 1
      loop%>
      
      
    </tr>
    <tr>
      <td colspan="11" align="center" height="3" width="669">
      <hr noshade color="#000000" size="3"></td>
    </tr>
    <tr>
    
    
      <%' inferior izquierdo temporario
    diente = 85
    do while diente > 80 %>
     <td width="65" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
            <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
              
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma2 where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
             end if
             Rs.close
				set Rs = nothing
	
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
 
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente - 1
      loop%>

      
      
      <td height="40" align="center" width="65" background="images/linea-vertical.jpg" bgcolor="#000000">
      <p align="center">&nbsp;</td>
      
      
<%' inferior derecho temporario
    diente = 71
    do while diente < 76 %>
     <td width="65" height="40" align="center">
     
     
     <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(diente)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
            
      %>
     
     <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     
     
           <%else
      
      if extraccion = "ei" then %>
           <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="40">
          <tr>
            <td width="40" height="40"><img border="0" src="images/aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
      
      <%
      else
     extraccion = "No"
      
      'tabla corona
      Set RSc = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      if not RSc.EOF then
      
      dientecor = Cstr(diente)
      corona = RSc(dientecor)
     
       if corona = "No" then
       borde = 0
       colorborde = "#111111"
       else
       if corona = "Si" then
       borde = 3
       colorborde = "#FF0000"
       else
       borde = 3
       colorborde = "#0000FF"
       end if
      end if
      
	  end if 
      Rsc.close
      set Rsc = nothing
 
       %>
      <div align="center">
        <center>
        <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="40" height="40">
          <tr>
            <td width="100%" height="100%">
            
            <%' tabla diente_muela
            if corona = "No" then
             Set RS = oConn.Execute("select * from odontograma2 where paciente = " & paciente & "") 
             if not RS.EOF then 
            dienteycara1 = diente & 1
            cara1 = RS(dienteycara1)
            dienteycara2 = diente & 2
            cara2 = RS(dienteycara2)
            dienteycara3 = diente & 3
            cara3 = RS(dienteycara3)
            dienteycara4 = diente & 4
            cara4 = RS(dienteycara4)
            dienteycara5 = diente & 5
            cara5 = RS(dienteycara5)
            
                         end if
             Rs.close
				set Rs = nothing
			
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
            %> <font size="1" face="Verdana">
            <span lang="es">
            <div align="center">
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40" height="18">
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara4%>" height="15">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="33%" bgcolor="#<%=cara5%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                  <td width="33%" bgcolor="#<%=cara2%>" height="10">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
            <span lang="es">
                  <td width="33%" bgcolor="#<%=cara3%>" height="10">
                  <font size="1">&nbsp;</font></td>
                </tr>
                </span>
                <tr>
                  <td width="100%" colspan="3" bgcolor="#<%=cara1%>" height="1">
                  <font size="1"><span lang="es">&nbsp;</span></font></td>
                </tr>
              </table>
              </center>
            </div>
            </font>
             <%
         
             ' fin tabla diente_muela%>
            </td>
          </tr>
        </table>
        </center>
      </div>
      <%'fin tabla corona %>
      
      
      <%   end if
      end if
      end if
      Rsx.close
      set Rsx = nothing
      
         ' fin tabla extraccion%>
        </td>
      <% diente = diente + 1
      loop%>
   



    </tr>
    <tr>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=85&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">85</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=84&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">84</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=83&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">83</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=82&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">82</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=81&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">81</font></a></b></span></td>
      <td height="40" align="center" width="4" background="images/linea-vertical.jpg" bgcolor="#000000">
      &nbsp;</td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=71&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">71</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=72&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">72</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=73&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">73</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=74&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">74</font></a></b></span></td>
      <td width="40" height="40" align="center"><span lang="es"><b>
      <a href="temporario.asp?elemento=75&paciente=<%=paciente%>">
      <font face="Verdana" color="#111111">75</font></a></b></span></td>
    </tr>
  </table>
  </center>
</div>
<%
oConn.Close
set oConn = nothing
%>
<p align="center"><font size="2" face="Verdana"><u><b><span lang="es">
      Actualizar Elemento</span></b></u><span lang="es"> (Haga 
      click en el número del elemento que desee actualizar)</span></font></p>
</body>

</html>