<%

dim strcon 'String de conex�o OLE DB

dim conexao 'Objeto de conex�o

set conexao = server.createobject("adodb.connection")

strcon = "DRIVER={microsoft access driver (*.mdb)};dbq=" & server.mappath("oscadastroteste.mdb") & ";uid=sa;pwd=;"

sub abreconexao()

	conexao.open strcon

end sub

sub fechaconexao()

	conexao.close
	
end sub


%>