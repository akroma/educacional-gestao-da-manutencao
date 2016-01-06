<%
set conexao = server.createobject("adodb.connection")
conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") '& ";uid=sa;pwd=;"
set bdcarg = conexao.execute("select codcarg from func where func = '" & session("func") & "'")
cargo = bdcarg("codcarg")
set bdcarg = nothing
conexao.close
set conexao = nothing
%>
<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>


<script language=JavaScript>
function abrepagina(arquivo)
{
open(arquivo,'POPUP','toolbar=no,location=no,status=no,scrollbars=yes,resizable=no,menubar=no,height=500,width=600,screenX=0,screenY=0,top=0,left=0');
}
</script>

<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
   <% response.write "<i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" %>
  </td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td><center><b>ORDEM DE SERVIÇO MANUTENÇÃO</b><center></td>
 <tr>
  <td>
   <table width=100% height=235>
    <tr>
     <td width=40%>
     </td>
     <td>
      <font face=verdana, arial size=1>
   <%  if cargo = "1" then 'Administrador
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=aprovaros.asp>APROVAR OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=direcionaros.asp>DIRECIONAR OS</a><br><br>"
        response.write "<a href=redirecionaros.asp>REDIRECIONAR OS</a><br><br>"
        response.write "<a href=reprovaros.asp>REPROVAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorioadmin.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=manutencao.asp>MANUTENÇÃO</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"
		response.write "<a href=funcionariocadastro.asp>CADASTRAR NOVO FUNCIONÁRIO</a><br><br>"

       else
        if cargo = "2" then 'Coordenador
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=aprovaros.asp>APROVAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorio.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"
        response.write "<a href=ajuda.asp target=_blank>AJUDA</a><br><br>"
 else 
  if cargo = "4" then 'Executor
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorio.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"
        response.write "<a href=ajuda.asp target=_blank>AJUDA</a><br><br>"
  else
   if cargo = "5" then 'Solicitante      
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorio.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"
        response.write "<a href=ajuda.asp target=_blank>AJUDA</a><br><br>"
else
  if cargo = "6" then 'Zelador
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=direcionaros.asp>DIRECIONAR OS</a><br><br>"
        response.write "<a href=redirecionaros.asp>REDIRECIONAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorioadmin.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"
   else
 if cargo = "3" then 'Diretor
        response.write "<a href=abriros.asp>ABRIR OS</a><br><br>"
        response.write "<a href=fecharos.asp>FECHAR OS</a><br><br>"
        response.write "<a href=andamentoos.asp>ANDAMENTO OS</a><br><br>"
        response.write "<a href=aprovaros.asp>APROVAR OS</a><br><br>"
        response.write "<a href=avaliaros.asp>AVALIAR OS</a><br><br>"
        response.write "<a href=reprovaros.asp>REPROVAR OS</a><br><br>"
        response.write "<a href=consultaos.asp>CONSULTAR OS</a><br><br>"
        response.write "<a href=relatorioadmin.asp>RELATÓRIOS</a><br><br>"
        response.write "<a href=senhaalterar.asp>ALTERAR DADOS</a><br><br>"

       else
    response.redirect("acessorestrito.asp")   
   end if
  end if
 end if
end if
end if
end if
%>
 </font>
     </td>
    </tr>
   </table>  
  </td>
 </tr>
 <tr>
  <td>
 <center><a href=index.asp><img src=sair.jpg border=0></a></center>
  </td>
 </tr>
 <tr>
  <td valign=bottom>
    <font face=Arial size=1>Copyright© SENAI </font>
  </td>
 </tr>
</table>
</body>
</html>
