<html>
<head>
	<title>CFP 1.15 | Ordem de Servi&ccedil;o</title>
</head> 
<body bgcolor=#F0FFF0><br><br>
	<center>
	<% response.write  WeekDayName(WeekDay(now()), False) & ", " & day(now()) & " de " & MonthName(Month(now()), False) & " de " & Year(date) & "<br>"
	response.expires = -1000
	%> <table align=center bgcolor=#C1CDC1 border=4 bordercolor=#8B8682 width=300 height=150 cellspacing=2 cellpadding=2>
		<tr>
			<td>
				<table align=center border=0>
				<%     if request.form("login") = "" or request.form("sen") = "" then
					session("func") = ""
					%>      <form method=post action=index.asp target=_self>   
						<tr>
							<td align=center>
								<h3>INFORME SEU LOGIN E SENHA</h>
								</td>  
								<tr>
									<td align=center><h5>LOGIN:</h>&nbsp;&nbsp;&nbsp;<input type=text name=login maxlength=10></td>
									</tr>   
									<tr>      
										<td align=center><h5>SENHA:</h>&nbsp;&nbsp;&nbsp;<input type=password name=sen maxlength=30></td>
										</tr>
										<tr>      
											<td colspan=2 align=center>
												<input type=image src=botao.gif>
											</td>
										</tr>
									</form>
									<%     else 
									set conexao = server.createobject("adodb.connection")
									' conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") '& ";uid=sa;pwd=;"
									conexao.open = "Data Source=" & Server.MapPath("oscadastroteste.mdb") & ";Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB;"
									
									set bdusuario = conexao.execute("select codfunc, func, sen, email, codcarg, codnucl from func where codfunc = '" & Request.Form("login") & "' and sen ='" & Request.Form("sen") & "'")
									if not bdusuario.eof then    'encontrou o registro procurado
										session("func") = bdusuario("func")
										session("ni") = bdusuario("codfunc")
										session("email") = bdusuario("email")
										session("cargo") = bdusuario("codcarg")
										session("nucleo") = bdusuario("codnucl")
										if bdusuario("sen") = "senaisp" then
											response.redirect("senhaalterar.asp")
										else 
											response.redirect("menu.asp")
										end if
									else 
										%>       <tr>
											<td align=center>
												<font face=Arial size=3 color=#ff0000>Login Inválido!</font><br>
												<font face=Arial size=2><a href=javascript:history.go(-1)>voltar</a></font>
											</td>
										</tr>
										<%      end if
										set bdusuario = nothing
										conexao.close
										set conexao = nothing
									end if
									%>    </table>
								</td>
							</tr>
						</table>
						<!-- <font face=Arial size=2><b>Ordem de Serviço <br>Desenvolvido na Escola SENAI Mario Amato</b><br>Versão 3 - Outubro/09</font></br>
						<font face=Arial size=1>Copyright© SENAI </font> -->
					</center>
				</bodY>
				</html>