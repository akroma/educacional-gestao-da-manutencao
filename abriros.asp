<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=abriros.asp target=_self>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
   <%
    session.LCID = 1046
    response.expires = -1000 
    response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
   %>
  </td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>ABRIR</td>
 </tr>
 <tr>
  <td align=center>
   <font face=arial size=2>
<%  if request.form("gravaros") = "" then 
     session("f5") = "" %>
     <table width=100% border=1 bordercolor=#8B8682 cellpadding=2 cellspacing=0>
      <tr>
       <td width=35%>NI do equipamento:</td>
       <td><input type=text name=nipatrfr size=10 maxlength=6></td>
      </tr>
      <tr>
       <td>Local:</td>
       <td><input type=text name=localfr size=55 maxlength=50></td>
      </tr>
      <tr>
       <td colspan=2>Serviço Solicitado:</td>
      </tr>
      <tr> 
       <td align=center colspan=2>
        <textarea name=servsofr rows=5 cols=45></textarea>
       </td>
      </tr>
     </table>
     <input type=hidden name=voltar value=nao>
     <input type=hidden name=gravaros value=sim>
<%   nros = "sim"
    else
     if request.form("localfr") = "" then
      response.write "O campo local não pode ser em branco"         
     else      
      if request.form("servsofr") = "" then 
       response.write "O campo Serviço Solicitado não pode ser em branco"   
      else
       if request.form("nipatrfr") = "" then  
        nipatrfr = 0
       else
        nipatrfr = request.form("nipatrfr")
       end if
       if isnumeric(nipatrfr) = false then
        response.write "O campo NI deve conter apenas números."
       else
        if len(request.form("servsofr")) > 225 then
         response.write "O campo Serviço Solicitado não pode ter mais de 225 caracteres."
        else 
         set conexao = server.createobject("adodb.connection")
         conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
           
           
             
          
         if session("f5") = "sim" then  ' Testando para impedir que o F5 gere outro número de OS
           


           response.redirect("menu.asp")
              
         else
          set bdmax = conexao.execute("select max(nros) as expr1 from os")
          nros = bdmax("expr1") + 1
          set bdmax = nothing 
          session("f5") = "sim"               
              
          if session("cargo") = "1" or session("cargo") = "2" or session("cargo") = "3" then
           set bdabros = conexao.execute("insert into os (nros, codfunc, nipatr, local, servso, dtapr, codsitua) values (" & nros & ",'" & session("ni") & "'," & nipatrfr & ",'" & request.form("localfr") & "','" & request.form("servsofr") & "', now(), 13)")
           mensagem = "A OS " & nros & " foi aberta com sucesso!<br><br><a href=abriros.asp>Deseja abrir outra OS</a>"
          else 
           set bdabros = conexao.execute("insert into os (nros, codfunc, nipatr, local, servso) values (" & nros & ",'" & session("ni") & "'," & nipatrfr & ",'" & request.form("localfr") & "','" & request.form("servsofr") & "')")
           mensagem = "A OS " & nros & " foi aberta com sucesso! <br><br>Anote o número e peça para o seu coordenador aprovar!!!<br><br><a href=abriros.asp>Deseja abrir outra OS</a>"
           
              
          'Envia um email para o coordenador aprovar a OS
              
'           set buscaemail = conexao.execute("select email from func where codsitua = 7 and codcarg = 2 and codnucl = " & session("nucleo"))
'               
'           assunto = "Por favor aprovar a OS " & nros & " aberta por " & session("func")
'           while not buscaemail.eof 
'            set correio = Server.CreateObject("CDONTS.NewMail") 
'            correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2
'            buscaemail.movenext
'           wend
'            
'           set correio = nothing
'           set buscaemail = nothing
          end if                
          set bdabros = Nothing    

            
         end if
             

         conexao.close
         set conexao = nothing
         response.write mensagem
        end if
       end if
      end if
     end if  
    end if %> 
   </font>
  </td>
 </tr>
 <tr>
  <td valign=bottom>
   <!--#include file=botoes.asp-->
  </td>
 </tr>
</table>
</form>
</body>
</html>
