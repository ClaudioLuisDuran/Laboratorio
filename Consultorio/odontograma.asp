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



<div id="preloader" style="position:absolute; left:220; top:90; width:475; height:57">
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
<b><span lang="es"><font face="Georgia" size="2" color="#B02A38">Actualizando 
odontograma de <%=nombre%> <%=apellido%></font></b></span><b><font face="Georgia"><font style="color:#B02A38; text-align:center" size="2"> ...</font><font color="#B02A38" size="2">
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
adulto [Orden fecha : <%=fecha%>]</span></font></p>

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
<%' Dibujo de elemento destacado%>
      <p align="center"><font size="2" face="Verdana"><u><b><span lang="es">
      Actualizar Elemento</span></b></u><span lang="es"> (Haga 
      click en el número del elemento que desee actualizar)</span></font></p>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td>
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  var radioSelected = false;
  for (i = 0;  i < theForm.eleccion.length;  i++)
  {
    if (theForm.eleccion[i].checked)
        radioSelected = true;
  }
  if (!radioSelected)
  {
    alert("Elija una de las opciones \"eleccion\".");
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="act_elem.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <div align="center">
          <center>
          <table border="0" cellpadding="4" cellspacing="4" style="border-collapse: collapse" bordercolor="#111111" width="690" height="144" bgcolor="#FFF4CC">
            <tr>
              <td width="69" rowspan="2" height="144" valign="top">
              <p align="center"><b>
              <span lang="es">
              <font face="Verdana" size="2">Detalle</font></span></b><p align="center"><b>
              <font face="Verdana" size="2"><span lang="es">elemento</font></b></span><p align="center"><span lang="es"><b><font size="6" face="Verdana"><%=elemento%></font></b>
              </span></td>
              <td height="144" width="145" rowspan="2" align="center" valign="top" background="images/sector<%=sector%>.gif">
              
    <%'tabla extraccion
     Set RSx = oConn.Execute("select * from extraccion where paciente = " & paciente & "") 
      if not RSx.EOF then
      dienteex = Cstr(elemento)
      extraccion = RSx(dienteex)
      
      if extraccion = "Si" then
       leyenda = "Elemento ausente"
       leyenda2 = "Anular sólo si la extracción está mal marcada"     
      %>
     
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="150" height="150">
                <tr>
                  
            <td width="150" height="150"><img border="0" src="images/elem_extraido.jpg"></td>

          </tr>
        </table>
       </center>
      </div>    
     <%else
     
         if extraccion = "ei" then
          leyenda = "Extracción indicada"  
          leyenda2 = "Anular sólo si la extracción está mal indicada" 
      %>
     
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="150" height="150">
                <tr>
                  
            <td width="150" height="150">
            <img border="0" src="images/elem_aextraer.jpg"></td>

          </tr>
        </table>
       </center>
      </div>  
     
     
        
      <%else
      extraccion = "No"
             ' comienza corona elemento elegido

      Set RSce = oConn.Execute("select * from corona where paciente = " & paciente & "") 
      
      if not RSce.EOF then
      
      elemcor = Cstr(elemento)
      coronaelem = RSce(elemcor)
     
              
       if coronaelem = "No" then
       borde = 0
       colorborde = "111111"
       leyenda = "1.- Marque la o las caras a pintar con una misma opción"
       leyenda2 = "2.- Tilde sólo la opción deseada y actualice" 
       else
       if coronaelem = "Si" then
       borde = 4
       colorborde = "#FF0000"
       leyenda = "Corona realizada"
       leyenda2 = "Anular sólo si la Corona está mal marcada" 
       else
       borde = 4
       colorborde = "#0000FF"
       leyenda = "Corona a realizar"
       leyenda2 = "Anular sólo si la Corona está mal indicada" 
       end if
      end if
	  end if 
      Rsce.close
      set Rsce = nothing
 
       %>
              <table border="<%=borde%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=colorborde%>" width="150" height="150">
                <tr>
                  <td colspan="3">
              <%'comienza elemento elegido
              if coronaelem = "No" then
            Set RSelem = oConn.Execute("select * from odontograma where paciente = " & paciente & "") 
            
            if not RSelem.EOF then 
            
            elemycara1 = elemento & 1
            cara1 = RSelem(elemycara1)
            elemycara2 = elemento & 2
            cara2 = RSelem(elemycara2)
            elemycara3 = elemento & 3
            cara3 = RSelem(elemycara3)
            elemycara4 = elemento & 4
            cara4 = RSelem(elemycara4)
            elemycara5 = elemento & 5
            cara5 = RSelem(elemycara5)
            
            
             end if
             RSelem.close
				set RSelem= nothing
			
			else
			cara1 = "FFFFFF"
			cara2 = "FFFFFF"
			cara3 = "FFFFFF"
			cara4 = "FFFFFF"
			cara5 = "FFFFFF"
		    end if 
		    
		    
		    
		    if elemento < 19 then
		    casilla1 = cara1
		    c1 = 1 
		    casilla2 = cara2
		    c2 = 2
		    casilla3 = cara3
		    c3 = 3
		    casilla4 = cara4
		    c4 = 4
		    casilla5 = cara5
		    c5 = 5
		    end if
		    
		    if (elemento < 29) and (elemento >18) then
		    casilla1 = cara1 
		    c1 = 1
		    casilla2 = cara2
		    c2 = 2
		    casilla3 = cara5
		    c3 = 5
		    casilla4 = cara4
		    c4 = 4
		    casilla5 = cara3
		    c5 = 3
		    end if		    
		    
		    if elemento > 40 then
		    casilla1 = cara4 
		    c1 = 4
		    casilla2 = cara2
		    c2 = 2
		    casilla3 = cara5
		    c3 = 5
		    casilla4 = cara1
		    c4 = 1
		    casilla5 = cara3	
		    c5 = 3	    
		    end if	        
		    
		    if (elemento < 39) and (elemento >30) then
		    casilla1 = cara4 
		    c1 = 4
		    casilla2 = cara2
		    c2 = 2
		    casilla3 = cara3
		    c3 = 3
		    casilla4 = cara1
		    c4 = 1
		    casilla5 = cara5
		    c5 = 5
		    end if	    
%>
              <div align="center">
                <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="150" height="150">
                <tr>
                  <td colspan="3" bgcolor="#<%=casilla1%>">
                  <p align="center">
                  <input type="checkbox" alt="prueba" name="<%=c1%>" value="ON"></td>
                </tr>
                <tr>
                  <td bgcolor="#<%=casilla5%>">
                  <p align="center">
                  <input type="checkbox" name="<%=c5%>" value="ON"></td>
                  <td bgcolor="#<%=casilla2%>">
                  <p align="center">
                  <input type="checkbox" name="<%=c2%>" value="ON"></td>
                  <td bgcolor="#<%=casilla3%>">
                  <p align="center">
                  <input type="checkbox" name="<%=c3%>" value="ON"></td>
                </tr>
                <tr>
                  <td colspan="3" bgcolor="#<%=casilla4%>">
                  <p align="center">
                  <input type="checkbox" name="<%=c4%>" value="ON"></td>
                </tr>
              </table>
                </center>
              </div>
              <%
		'fin elemento elegido
		
		%>
		
              </td>
              </tr>
             </table>
             
       <%end if
       end if
       end if ' fin extraccion elemento elegido
       RSx.close
		set RSx = nothing
       %> 
          
              <p align="center"><font size="2" face="Verdana"><span lang="es"><%=leyenda%> </span></font>
          
             <% 
             ' fin corona elemento elegido
             %>
              
              </td>
              <td height="1" width="426" bgcolor="#FFFFE8">
              <p align="center"><span lang="es"><font size="2" face="Verdana">
              <b>Opciones</b></font></span></td>
            </tr>
            <tr>
              <td height="176" width="426" valign="top">
              <div align="center">
                <center>
                <table border="0" cellpadding="3" cellspacing="4" style="border-collapse: collapse" bordercolor="#111111" width="432" height="258">
                  
                 
                  <tr>
                    <td align="center" bgcolor="#FFFFE8" height="12" width="418">
                    <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="72%" id="AutoNumber2" height="186">
                      <% if extraccion = "No" and coronaelem = "No" then %>
                      
                      <tr>
                        <td width="51%" align="center" height="20">
                        
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
                          <tr>
                            <td width="32%" bgcolor="#FF0000">
                            <p align="center">
                    <span lang="es">
                    
                    <!--webbot bot="Validation" b-value-required="TRUE" --><input type="radio" value="Rojo" name="eleccion"></span></td>
                            <td width="68%" bgcolor="#FF0000">
                            <p align="center"><span lang="es"><b>
                    <font face="Verdana" size="2" color="#FFFFFF">Rojo</font></b></span></td>
                          </tr>
                        </table>
                        </td>
                        <td width="58%" align="center" height="20"><span lang="es">
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="103%" id="AutoNumber3">
                          <tr>
                            <td width="30%" bgcolor="#0000FF">
                            <p align="center">
                    <span lang="es">
                    
                            <font color="#0000FF">
                    <input type="radio" value="Azul" name="eleccion"></font></span></td>
                            <td width="70%" bgcolor="#0000FF">
                    <span lang="es">
                    
                            <p align="center"><b>
                    <font face="Verdana" size="2" color="#FFFFFF">Azul</font></b></span></td>
                          </tr>
                        </table>

                        </span></td>
                      </tr>
                      
                      <%end if%>
                      
                      <tr>
                        <td width="51%" align="center" height="32">
                        
                      <% if (extraccion = "No" and coronaelem = "No") or coronaelem = "Cor" then %>  
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
                          <tr>
                            <td width="32%">
                            <span lang="es">
                                 <div align="center">
                                   <center>
                                 <table border="3" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="32">
                        <tr>
                          <td width="26" bgcolor="#FFFFFF">
                          <p align="center">
                          <input type="radio" value="CoronaSi" name="eleccion"></td>
                        </tr>
                      </table>
                                   </center>
                            </div>
                            </span></td>
                            <span lang="es">
                            <td width="68%" bgcolor="#FF0000">
                            <p align="center">
                    <span lang="es">
                    
                            <b>
                    <font face="Verdana" size="2" color="#FFFFFF">
                            Corona realizada</font></b></span></td>
                          </tr>
                        </table>
                        
                        <%end if%>

                        </span></td>
                        <td width="58%" align="center" height="32">
                        
                      <% if (extraccion = "No" and coronaelem = "No") or coronaelem = "Si" then %>  
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="104%" id="AutoNumber3">
                          <tr>
                            <td width="29%">
                            <span lang="es">
                                <div align="center">
                                  <center>
                                <table border="3" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#0000FF" width="33">
                        <tr>
                          <td width="27" bgcolor="#FFFFFF">
                          <p align="center">
                          <input type="radio" value="CoronaCor" name="eleccion"></td>
                        </tr>
                      </table>
                                  </center>
                            </div>
                            </span></td>
                            <span lang="es">
                            <td width="71%" bgcolor="#0000FF">
                            <p align="center">
                    <span lang="es">
                    
                    <b><font face="Verdana" size="2" color="#FFFFFF">
                            Corona a realizar</font></b></span></td>
                          </tr>
                        </table>
                       <%end if%> 
                        

                        </span></td>
                      </tr>
                      <tr>
                        <td width="51%" align="center" height="59"><span lang="es">
                        
                         <% if (extraccion = "No" and coronaelem = "No") or extraccion = "ei" then %> 
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3" height="47">
                          <tr>
                            <td width="28%" background="images/extraido.jpg" height="40">
                            <p align="center">
                    <span lang="es">
                     
                    <input type="radio" value="ExtraccionSI" name="eleccion"></span></td>
                            <td width="72%" bgcolor="#FF0000" height="40">
                            <p align="center"><span lang="es"><b>
                    <font face="Verdana" size="2" color="#FFFFFF">Elemento ausente</font></b></span></td>
                          </tr>
                        </table>
                        <%end if%>
                        

                        </span></td>
                        <td width="58%" align="center" height="59"><span lang="es">
                        
                       <% if (extraccion = "No" and coronaelem = "No") or extraccion = "Si" then %>  
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="104%" id="AutoNumber3" height="47">
                          <tr>
                            <td width="28%" background="images/aextraer.jpg" height="40">
                            <p align="center">
                    <span lang="es">
                     
                    <input type="radio" value="ExtraccionEI" name="eleccion"></span></td>
                            <td width="83%" bgcolor="#0000FF" height="40">
                            <p align="center">
                    <span lang="es">
                     
                    <b><font face="Verdana" size="2" color="#FFFFFF">
                            Extracción indicada</font></b></span></td>
                          </tr>
                        </table>
                        <%end if%>

                        </span></td>
                      </tr>
                      <tr>
                        <td width="51%" align="center" height="48"><span lang="es">
                        <table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="138" id="AutoNumber3">
                          <tr>
                            <td width="20" bgcolor="#FFFFFF">
                            <p align="center">
                    <input type="radio" value="Anular" name="eleccion"></td>
                            <td width="101" bgcolor="#FFFFFF">
                            <p align="center"><span lang="es"><b>
                    <font face="Verdana" size="2">Anular marca errónea</font></b></span></td>
                          </tr>
                        </table>

                        </span></td>
                        <td width="58%" align="center" height="48">
                    <input type="submit" value="Actualizar" name="B1"></td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                  
                 
                   
                    
                    
                  
                    
                      
                  <tr>
                    <td align="center" bgcolor="#FFFFE8" height="38" width="418">
                    <font size="1" face="Verdana" color="#111111"><span lang="es">
                    <a href="temporario.asp?paciente=<%=paciente%>">
                    <font color="#111111">¿Desea 
                    consultar Odontograma de Elementos Temporarios de este 
                    paciente?</font></a></span></font></td>
                  </tr>
                  
                 
                  <tr>
                    <td align="center" bgcolor="#FFFFE8" height="31" width="418">
                    <span lang="es">
                    <font size="1" face="Verdana" color="#111111">
                    <a href="completo.asp?paciente=<%=paciente%>">
                    <font color="#111111">¿Desea ver 
                    juntos los Odontogramas de este paciente?</font></a></font></span></td>
                  </tr>
                  
                 
                   
                    
                    
                  
                    
                      
                </table>
                </center>
              </div>
              </td>
            </tr>
          </table>
          </center>
        </div>
        <input type="hidden" name="paciente" value="<%=paciente%>">
        <input type="hidden" name="elemento" value="<%=elemento%>">
        <input type="hidden" name="extraccion" value="<%=extraccion%>">
        <input type="hidden" name="corona" value="<%=coronaelem%>">
        <input type="hidden" name="accion" value="Actualización de Odontograma Adulto">
        <input type="hidden" name="odontologo" value="<%=odontologo%>">
        <input type="hidden" name="fecha" value="<%=fecha%>">
      </form>
      </td>
    </tr>
  </table>
  </center>
</div>

<%
oConn.Close
set oConn = nothing
%>

</body>

</html>