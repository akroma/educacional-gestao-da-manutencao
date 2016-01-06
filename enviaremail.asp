<%
'crio o objeto correio'
'Set Mail = Server.CreateObject("SMTPsvg.Mailer")
set mail = server.createObject("Persits.MailSender") 
'configuro a mensagem 
'assinalo o servidor de saída para enviar o correio 
mail.host = "sn116anotes1.cfp116.sp.senai.br" 
'indico o endereço de correio do remitente 
mail.from = "suporteinf116@sp.senai.br" 
'indico o endereço do destinatário da mensagem 
mail.addAddress "edilsonfsp@ig.com.br" 
'indico o corpo da mensagem 
mail.body = "Se você recebeu funcionou" 
'o envio 
'certifico-me que não se apresentem erros na página se se produzem 


On Error Resume Next 
mail.send 
if Err <> 0 then 
response.write "Erro, não pode completar a operação" 
else 
response.write "Obrigado por preencher o formulário. Foi enviado corretamente." end if 
%>