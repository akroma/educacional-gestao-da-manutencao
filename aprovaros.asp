<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=aprovaros.asp target=_self>
 <table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
  <tr height=15>
   <td align=center>
<%  set conexao = server.createobject("adodb.connection")
    conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"  
   if session("cargo") = 4 or session("cargo") = 5 or session("cargo") = 6 then 
    response.redirect("acessorestrito.asp")
   end if
    session.LCID = 1046
    response.expires = -1000
    response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%> </td>
  </tr>
  <tr bgcolor=#8B8682 height=15>
   <td align=center>APROVAR</td>
  </tr>
  <tr>
   <td align=center>
     <table width=100% border=0>
      <tr> 
       <td align=center>
        <font face=arial size=2>  
<%       if request.form("confbusca") = "" then 'Verifica para buscar OS aberta e não aprovadas
          set bdapros = conexao.execute("select nros, codfunc from oscon where codsitua = 12 order by nros")
          if not bdapros.eof then
           response.write "<select name=nrosapfr>"
           while not bdapros.eof
            response.write "<option value=" & bdapros("nros") & ">" 
            set bdsolic = conexao.execute("select func from func where codfunc = '" & bdapros("codfunc") & "'")
            solicitante = bdsolic("func")
            response.write bdapros("nros") & " - " & solicitante & "</option>" & "<br>"
            set bdsolic = nothing
            bdapros.movenext
           wend
           response.write "</select><input type=hidden name=confbusca value=sim>"
          else
           response.write "Não existe OS para aprovar"
          end if
          set bdapros = nothing 
         else
          if request.form("confgr") = "" then
           set bdapros = conexao.execute("select codfunc, nros, nipatr, local, servso, dtabe from oscon where nros =" & request.form("nrosapfr"))%>
           <table width=100% border=1 bordercolor=#8B8682 cellpadding=2 cellspacing=0>
            <tr>
             <td width=25%><font face=Arial size=2>OS:</font></td>
             <td><font face=Arial size=2><%=bdapros("nros")%></font></td>
            </tr>
<%         if bdapros("nipatr") <> 0 then %>
           <tr>
            <td width=33%><font face=Arial size=2>NI:</font></td>
             <td><font face=Arial size=2><%=bdapros("nipatr")%></font></td>
           </tr>
<%         end if %>
            <tr>
             <td><font face=Arial size=2>Local:</font></td>
             <td><font face=Arial size=2><%=bdapros("local")%></font></td>
            </tr>
            <tr>
             <td><font face=Arial size=2>Solicitante:</font></td>
             <td><font face=Arial size=2>
<%            set bdsolic = conexao.execute("select func from func where codfunc = '" & bdapros("codfunc") & "'")
              response.write bdsolic("func") 
              set bdsolic = nothing
%>          </tr>
            <tr>
             <td valign=top height=50><font face=Arial size=2>Dados Pedido:</font></td>
             <td valign=top><font face=Arial size=2><p align=justify><%=bdapros("servso")%></p></font></td>
            </tr>
            <tr>
             <td><font face=Arial size=2>Data abertura:</font></td>
             <td><font face=Arial size=2><%=bdapros("dtabe")%></font></td>
            </tr>
            <tr>
             <td><font face=Arial size=2>Direcionar para:</font></td>
             <td>
              <select name=direcfr>
               <option selected value=1></option>
               <option value=2>SOLICITANTE</option>
               <option value=3>MANUTENÇÃO</option>
               <option value=4>REPROVADA</option>
              </select>
             </td>
            </tr>
            <tr>
             <td colspan=2><font face=Arial size=2>Observação:</font></td>
            </tr>
            <tr> 
             <td align=center colspan=2>
              <textarea name=obsfr rows=4 cols=45></textarea>
             </td>
            </tr>
            <tr> 
             <td align=center colspan=2>
              <input type=hidden name=confbusca value=sim>
              <input type=hidden name=confgr value=sim>
              <input type=hidden name=nrosapfr value=<%=bdapros("nros")%>>
              <input type=hidden name=solicfr value=<%=bdapros("codfunc")%>>
             </td>
            </tr>
           </table>
<%         set bdapros = nothing   
          else  
           if len(request.form("obsfr")) < 225 then
            if request.form("direcfr") = "2" then
    	     set bdapros = conexao.execute("update os set dtapr = now(), codsitua = 1, executor = '" & request.form("solicfr") & "' where nros = " & request.form("nrosapfr")) 
             response.write "A OS " & request.form("nrosapfr") & " foi direcionada para o docente!<br><a href=aprovaros.asp>Aprovar outra OS</a>"
             set bdapros = nothing 
            end if
            if request.form("direcfr") = "3" then
             set bdapros = conexao.execute("update os set dtapr = now(), codsitua = 13 where nros = " & request.form("nrosapfr"))
             response.write "A OS " & request.form("nrosapfr") & " foi direcionada para a manutenção!<br><a href=aprovaros.asp>Aprovar outra OS</a>"
             set bdapros = nothing 
            end if
            if request.form("direcfr") = "4" then

            if request.form("obsfr") <> "" then            
             set bdapros = conexao.execute("update os set dtapr = now(), codsitua = 3 where nros = " & request.form("nrosapfr"))
             response.write "A OS " & request.form("nrosapfr") & " foi reprovada!<br><a href=aprovaros.asp>Aprovar outra OS</a>"
             set bdapros = nothing 


set buscaemail = conexao.execute("select email from func where codfunc ='" & request.form("solicfr") & "'")

assunto = "A OS " & request.form("nrosfr") & " foi reprovada por " & session("func")

corpoemail = request.form("obsfr") & "- Este correio foi enviado automaticamente pelo programa de OS, por favor não responder."

set correio = Server.CreateObject("CDONTS.NewMail") 

correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, corpoemail, 2

set correio = nothing

set buscaemail = nothing


            else
             response.write "O campo observações não pode ser em branco quando a OS é reprovada.<br><a href=javascript:history.go(-1)>Voltar</a>"
          
            end if
            end if


            if request.form("direcfr") = "1" then
             response.write "O campo direcionar não pode ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>"
            end if
            if request.form("obsfr") <> "" then
             set bdapros = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosapfr") & ",'" & request.form("obsfr") & "','" & session("ni") & "')")
            end if
           else
             response.write "O campo observações deve conter menos de 225 caracteres<br><a href=javascript:history.go(-1)>Voltar</a>"
           end if
          end if 
         end if
         conexao.close
         set conexao = nothing 
%>      </font>
       </td>
      </tr>
     </table>
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
