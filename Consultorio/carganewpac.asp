<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nuevo paciente</title>
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

<body onload="is_loaded();">

<%  
' recepcion de paciente
' paciente tal
  paciente = request("paciente")
  paciente = cint(paciente)
  consulta = request("consulta")
'response.write paciente


fecha_actual = Now()
Dia = Day(fecha_actual)
Mes = Month(fecha_actual)
Anio = Year(fecha_actual)
fecha_ok = Mes &"/"& Dia &"/"& Anio

set oConn =  Server.CreateObject("ADODB.Connection")
  oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")
  Set RSpac = oConn.Execute("select * from fichas where paciente = " & paciente & "")
  if not RSpac.EOF then
  nombre = RSpac("nombre")
  apellido = RSpac("apellido")
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
<b><span lang="es"><font face="Georgia" size="2" color="#B02A38">Cargando datos&nbsp; 
de <%=nombre%> <%=apellido%></font></b></span><b><font face="Georgia"><font style="color:#B02A38; text-align:center" size="2"> ...</font><font color="#B02A38" size="2">
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


<%

Set oConn = Server.CreateObject("ADODB.Connection")

nombre = request.form("nombre")
apellido = request.form("apellido")
paciente = request.form("paciente")

' grabo ficha

set RS = Server.CreateObject("ADODB.Recordset")  


oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/fichas.mdb")

             RS.Open "fichas",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("apellido") = request.form("apellido")
             RS("nombre") = request.form("nombre")
             RS("obrasocial") = request.form("obrasocial")
             RS("afiliadonro") = request.form("afiliadonro")
             RS("domicilio") = request.form("domicilio")
             RS("ciudad") = request.form("ciudad")
             RS("provincia") = request.form("provincia")
             RS("pais") = request.form("pais")                          
             RS("telefono") = request.form("telefono")     
             RS("email") = request.form("email")
             RS("odontologo") = request.form("odontologo")
             RS("matricula") = request.form("matricula")
             RS("observaciones") = request.form("observaciones")
             RS("fecha") = fecha_ok
                        
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close

'Incializacion del Historial

accion = "Creacion de ficha"
set RS = Server.CreateObject("ADODB.Recordset")  


oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/historial.mdb")

             RS.Open "historial",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("responsable") = request.form("odontologo")
             RS("accion") = accion
             RS("fecha") = fecha_ok
                        
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close


'creacion de odontogramas
' Coronas

set RS = Server.CreateObject("ADODB.Recordset")  
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

             RS.Open "corona",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("11") = "No"
             RS("12") = "No"
             RS("13") = "No"
             RS("14") = "No"
             RS("15") = "No"
             RS("16") = "No"
             RS("17") = "No"                          
             RS("18") = "No"     
             RS("21") = "No"
             RS("22") = "No"
             RS("23") = "No"
             RS("24") = "No"
             RS("25") = "No"
             RS("26") = "No"
             RS("27") = "No"
             RS("28") = "No"
             RS("31") = "No"
             RS("32") = "No"
             RS("33") = "No"
             RS("34") = "No"
             RS("35") = "No"
             RS("36") = "No"
             RS("37") = "No"
             RS("38") = "No" 
             RS("51") = "No"
             RS("52") = "No"
             RS("53") = "No"
             RS("54") = "No"
             RS("55") = "No"
             RS("61") = "No"
             RS("62") = "No"
             RS("63") = "No"
             RS("64") = "No"
             RS("65") = "No"
             RS("71") = "No"
             RS("72") = "No"
             RS("73") = "No"
             RS("74") = "No"
             RS("75") = "No"
             RS("81") = "No"
             RS("82") = "No"
             RS("83") = "No"
             RS("84") = "No"
             RS("85") = "No"
             RS("41") = "No"
             RS("42") = "No"
             RS("43") = "No"
             RS("44") = "No"             
             RS("45") = "No"
             RS("46") = "No"
             RS("47") = "No"
             RS("48") = "No" 
             
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close


' extracciones

set RS = Server.CreateObject("ADODB.Recordset")  
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

             RS.Open "extraccion",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("11") = "No"
             RS("12") = "No"
             RS("13") = "No"
             RS("14") = "No"
             RS("15") = "No"
             RS("16") = "No"
             RS("17") = "No"                          
             RS("18") = "No"     
             RS("21") = "No"
             RS("22") = "No"
             RS("23") = "No"
             RS("24") = "No"
             RS("25") = "No"
             RS("26") = "No"
             RS("27") = "No"
             RS("28") = "No"
             RS("31") = "No"
             RS("32") = "No"
             RS("33") = "No"
             RS("34") = "No"
             RS("35") = "No"
             RS("36") = "No"
             RS("37") = "No"
             RS("38") = "No" 
             RS("51") = "No"
             RS("52") = "No"
             RS("53") = "No"
             RS("54") = "No"
             RS("55") = "No"
             RS("61") = "No"
             RS("62") = "No"
             RS("63") = "No"
             RS("64") = "No"
             RS("65") = "No"
             RS("71") = "No"
             RS("72") = "No"
             RS("73") = "No"
             RS("74") = "No"
             RS("75") = "No"
             RS("81") = "No"
             RS("82") = "No"
             RS("83") = "No"
             RS("84") = "No"
             RS("85") = "No"
             RS("41") = "No"
             RS("42") = "No"
             RS("43") = "No"
             RS("44") = "No"             
             RS("45") = "No"
             RS("46") = "No"
             RS("47") = "No"
             RS("48") = "No" 
             
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close

'odontograma adulto

set RS = Server.CreateObject("ADODB.Recordset")  
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

             RS.Open "odontograma",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("111") = "FFFFFF"
             RS("112") = "FFFFFF"
             RS("113") = "FFFFFF"
             RS("114") = "FFFFFF"
             RS("115") = "FFFFFF"
             RS("121") = "FFFFFF"
             RS("122") = "FFFFFF"                          
             RS("123") = "FFFFFF"     
             RS("124") = "FFFFFF"
             RS("125") = "FFFFFF"
             RS("131") = "FFFFFF"
             RS("132") = "FFFFFF"
             RS("133") = "FFFFFF"
             RS("134") = "FFFFFF"
             RS("135") = "FFFFFF"
             RS("141") = "FFFFFF"
             RS("142") = "FFFFFF"
             RS("143") = "FFFFFF"
             RS("144") = "FFFFFF"
             RS("145") = "FFFFFF"
             RS("151") = "FFFFFF"
             RS("152") = "FFFFFF"
             RS("153") = "FFFFFF"
             RS("154") = "FFFFFF" 
             RS("155") = "FFFFFF"
             RS("161") = "FFFFFF"
             RS("162") = "FFFFFF"
             RS("163") = "FFFFFF"
             RS("164") = "FFFFFF"
             RS("165") = "FFFFFF"
             RS("171") = "FFFFFF"
             RS("172") = "FFFFFF"
             RS("173") = "FFFFFF"
             RS("174") = "FFFFFF"
             RS("175") = "FFFFFF"
             RS("181") = "FFFFFF"
             RS("182") = "FFFFFF"
             RS("183") = "FFFFFF"
             RS("184") = "FFFFFF"
             RS("185") = "FFFFFF"
             RS("211") = "FFFFFF"
             RS("212") = "FFFFFF"
             RS("213") = "FFFFFF"
             RS("214") = "FFFFFF"
             RS("215") = "FFFFFF"
             RS("221") = "FFFFFF"
             RS("222") = "FFFFFF"
             RS("223") = "FFFFFF"             
             RS("224") = "FFFFFF"
             RS("225") = "FFFFFF"
             RS("231") = "FFFFFF"
             RS("232") = "FFFFFF" 
             RS("233") = "FFFFFF"
             RS("234") = "FFFFFF"
             RS("235") = "FFFFFF"
             RS("241") = "FFFFFF"
             RS("242") = "FFFFFF"
             RS("243") = "FFFFFF"
             RS("244") = "FFFFFF"
             RS("245") = "FFFFFF"
             RS("251") = "FFFFFF"
             RS("252") = "FFFFFF"                          
             RS("253") = "FFFFFF"     
             RS("254") = "FFFFFF"
             RS("255") = "FFFFFF"
             RS("261") = "FFFFFF"
             RS("262") = "FFFFFF"
             RS("263") = "FFFFFF"
             RS("264") = "FFFFFF"
             RS("265") = "FFFFFF"
             RS("271") = "FFFFFF"
             RS("272") = "FFFFFF"
             RS("273") = "FFFFFF"
             RS("274") = "FFFFFF"
             RS("275") = "FFFFFF"
             RS("281") = "FFFFFF"
             RS("282") = "FFFFFF"
             RS("283") = "FFFFFF"
             RS("284") = "FFFFFF" 
             RS("285") = "FFFFFF"
             RS("311") = "FFFFFF"
             RS("312") = "FFFFFF"
             RS("313") = "FFFFFF"
             RS("314") = "FFFFFF"
             RS("315") = "FFFFFF"
             RS("321") = "FFFFFF"
             RS("322") = "FFFFFF"
             RS("323") = "FFFFFF"
             RS("324") = "FFFFFF"
             RS("325") = "FFFFFF"
             RS("331") = "FFFFFF"
             RS("332") = "FFFFFF"
             RS("333") = "FFFFFF"
             RS("334") = "FFFFFF"
             RS("335") = "FFFFFF"
             RS("341") = "FFFFFF"
             RS("342") = "FFFFFF"
             RS("343") = "FFFFFF"
             RS("344") = "FFFFFF"
             RS("345") = "FFFFFF"
             RS("351") = "FFFFFF"
             RS("352") = "FFFFFF"
             RS("353") = "FFFFFF"             
             RS("354") = "FFFFFF"
             RS("355") = "FFFFFF"
             RS("361") = "FFFFFF"
             RS("362") = "FFFFFF" 
             RS("363") = "FFFFFF"
             RS("364") = "FFFFFF"
             RS("365") = "FFFFFF"  
             RS("371") = "FFFFFF"
             RS("372") = "FFFFFF"
             RS("373") = "FFFFFF"
             RS("374") = "FFFFFF"
             RS("375") = "FFFFFF"
             RS("381") = "FFFFFF"
             RS("382") = "FFFFFF"
             RS("383") = "FFFFFF"
             RS("384") = "FFFFFF"
             RS("385") = "FFFFFF"
             RS("411") = "FFFFFF"
             RS("412") = "FFFFFF"
             RS("413") = "FFFFFF"
             RS("414") = "FFFFFF"
             RS("415") = "FFFFFF"
             RS("421") = "FFFFFF"
             RS("422") = "FFFFFF"
             RS("423") = "FFFFFF"
             RS("424") = "FFFFFF"
             RS("425") = "FFFFFF"
             RS("431") = "FFFFFF"
             RS("431") = "FFFFFF"
             RS("432") = "FFFFFF"
             RS("433") = "FFFFFF"             
             RS("434") = "FFFFFF"
             RS("435") = "FFFFFF"
             RS("441") = "FFFFFF"
             RS("442") = "FFFFFF" 
             RS("443") = "FFFFFF"
             RS("444") = "FFFFFF"
             RS("445") = "FFFFFF"              
             RS("451") = "FFFFFF"
             RS("452") = "FFFFFF"
             RS("453") = "FFFFFF"
             RS("454") = "FFFFFF"
             RS("455") = "FFFFFF"
             RS("461") = "FFFFFF"
             RS("462") = "FFFFFF"
             RS("463") = "FFFFFF"
             RS("464") = "FFFFFF"
             RS("465") = "FFFFFF"
             RS("471") = "FFFFFF"
             RS("472") = "FFFFFF"
             RS("473") = "FFFFFF"
             RS("474") = "FFFFFF"
             RS("475") = "FFFFFF"
             RS("481") = "FFFFFF"
             RS("482") = "FFFFFF"
             RS("483") = "FFFFFF"
             RS("484") = "FFFFFF"
             RS("485") = "FFFFFF"
           
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close

'odontograma temporarios

set RS = Server.CreateObject("ADODB.Recordset")  
oConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../../db/pacientes.mdb")

             RS.Open "odontograma2",oConn,2,2
             
             RS.AddNew
             
             RS("paciente") = request.form("paciente")
             RS("511") = "FFFFFF"
             RS("512") = "FFFFFF"
             RS("513") = "FFFFFF"
             RS("514") = "FFFFFF"
             RS("515") = "FFFFFF"
             RS("521") = "FFFFFF"
             RS("522") = "FFFFFF"                          
             RS("523") = "FFFFFF"     
             RS("524") = "FFFFFF"
             RS("525") = "FFFFFF"
             RS("531") = "FFFFFF"
             RS("532") = "FFFFFF"
             RS("533") = "FFFFFF"
             RS("534") = "FFFFFF"
             RS("535") = "FFFFFF"
             RS("541") = "FFFFFF"
             RS("542") = "FFFFFF"
             RS("543") = "FFFFFF"
             RS("544") = "FFFFFF"
             RS("545") = "FFFFFF"
             RS("551") = "FFFFFF"
             RS("552") = "FFFFFF"
             RS("553") = "FFFFFF"
             RS("554") = "FFFFFF" 
             RS("555") = "FFFFFF"
             RS("611") = "FFFFFF"
             RS("612") = "FFFFFF"
             RS("613") = "FFFFFF"
             RS("614") = "FFFFFF"
             RS("615") = "FFFFFF"
             RS("621") = "FFFFFF"
             RS("622") = "FFFFFF"
             RS("623") = "FFFFFF"
             RS("624") = "FFFFFF"
             RS("625") = "FFFFFF"
             RS("631") = "FFFFFF"
             RS("632") = "FFFFFF"
             RS("633") = "FFFFFF"
             RS("634") = "FFFFFF"
             RS("635") = "FFFFFF"
             RS("641") = "FFFFFF"
             RS("642") = "FFFFFF"
             RS("643") = "FFFFFF"
             RS("644") = "FFFFFF"
             RS("645") = "FFFFFF"
             RS("651") = "FFFFFF"
             RS("652") = "FFFFFF"
             RS("653") = "FFFFFF"             
             RS("654") = "FFFFFF"
             RS("655") = "FFFFFF"
             RS("711") = "FFFFFF"
             RS("712") = "FFFFFF" 
             RS("713") = "FFFFFF"
             RS("714") = "FFFFFF"
             RS("715") = "FFFFFF"
             RS("721") = "FFFFFF"
             RS("722") = "FFFFFF"
             RS("723") = "FFFFFF"
             RS("724") = "FFFFFF"
             RS("725") = "FFFFFF"
             RS("731") = "FFFFFF"
             RS("732") = "FFFFFF"                          
             RS("733") = "FFFFFF"     
             RS("734") = "FFFFFF"
             RS("735") = "FFFFFF"
             RS("741") = "FFFFFF"
             RS("742") = "FFFFFF"
             RS("743") = "FFFFFF"
             RS("744") = "FFFFFF"
             RS("745") = "FFFFFF"
             RS("751") = "FFFFFF"
             RS("752") = "FFFFFF"
             RS("753") = "FFFFFF"
             RS("754") = "FFFFFF"
             RS("755") = "FFFFFF"
             RS("811") = "FFFFFF"
             RS("812") = "FFFFFF"
             RS("813") = "FFFFFF"
             RS("814") = "FFFFFF" 
             RS("815") = "FFFFFF"
             RS("821") = "FFFFFF"
             RS("822") = "FFFFFF"
             RS("823") = "FFFFFF"
             RS("824") = "FFFFFF"
             RS("825") = "FFFFFF"
             RS("831") = "FFFFFF"
             RS("832") = "FFFFFF"
             RS("833") = "FFFFFF"
             RS("834") = "FFFFFF"
             RS("835") = "FFFFFF"
             RS("841") = "FFFFFF"
             RS("842") = "FFFFFF"
             RS("843") = "FFFFFF"
             RS("844") = "FFFFFF"
             RS("845") = "FFFFFF"
             RS("851") = "FFFFFF"
             RS("852") = "FFFFFF"
             RS("853") = "FFFFFF"
             RS("854") = "FFFFFF"
             RS("855") = "FFFFFF"
                                      
             RS.Update
             RS.Close
             
set RS=nothing
oConn.Close



'Response.Redirect "carga.asp"

%>


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="490">
    <tr>
      <td height="134"><img border="0" src="../images/Sup_C.jpg"></td>
    </tr>
    <tr>
      <td height="23">
      &nbsp;</td>
    </tr>
    <tr>
      <td height="282">
      <div align="center">
        <center>
        <table border="1" cellpadding="10" cellspacing="10" style="border-collapse: collapse" bordercolor="#2D4773" width="90%" id="AutoNumber1">
          <tr>
            <td width="100%">
            <p align="center"><font face="Verdana" color="#000080">
            <span lang="es"><br>
            [&nbsp;&nbsp;&nbsp; La ficha ha sido cargada con éxito&nbsp;&nbsp;&nbsp; 
            ]</span></font></p>
            <p align="center">&nbsp;</p>
            <p align="center"><b><span lang="es">
            <font face="Verdana" size="2" color="#000080">¿Que desea hacer 
            ahora?</font></span></b></p>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080">
            <a href="verficha.asp?paciente=<%=paciente%>"><font color="#000080">Ver ficha del paciente 
            <%=nombre%>&nbsp;<%=apellido%></font></a></font></span></p>
            
           <%if consulta = TRUE then%>
            <p align="center">
            <a href="ini_cons.asp?paciente=<%=paciente%>">
            <font face="Verdana" size="2" color="#000080"><span lang="es">
            Actualizar Odontograma para 
            <%=nombre%>&nbsp;<%=apellido%></span></font></a><span lang="es"><span lang="es"><font face="Verdana" size="2" color="#000080">
            </font></span></span></p>
            <%else%>
            <p align="center">
            <font color="#000080" face="Verdana" size="2"><span lang="es">
            <a href="diagnostico.asp?paciente=<%=paciente%>">
            <font color="#000080">Cargar odontograma del archivo del 
            paciente <%=nombre%>&nbsp;<%=apellido%></span></font></a></font></p>
            <%end if%>
            <p align="center"><span lang="es">
            <font face="Verdana" size="2" color="#000080"><a href="newpac.asp">
            <font color="#000080">Agregar fichas de nuevos pacientes de archivo</font></a></font></span></td>
          </tr>
        </table>
        </center>
      </div>
      

      <p>&nbsp;</td>
    </tr>
    <tr>
      <td height="51" bgcolor="#2D4773">
      <p align="center"><font color="#FFFFFF" face="Verdana" size="2">[
      <a href="menu.asp"><span lang="es"><font color="#FFFFFF">Volver al Menú</font></span><font color="#FFFFFF">
      </font></a>]</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>