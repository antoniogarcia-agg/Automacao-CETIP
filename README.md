# Automatização de Atualização de PU de eventos

O código é baseado no documento anexado da CETIP "ENVIAR ARQUIVOS (2)".

# Ferramentas
1. Python

# Cuidados
Deve haver, na pasta com os códigos, uma planilha com colunas ['CRI / CRA', 'IF', 'Data de pagamento', 'Juros PU', 'Amortização PU', 'Saldo Devedor PU', 'Indexador'], que são puxadas pelo código e tem as informações necessárias para gerar o arquivo txt.
O script original.py gera um executável que seleciona essa planilha e a pasta de destino dos arquivos (um para CRI's e outro para CRA's).
Deve-se alterar o nome da planilha na linha 18 e o Nome Simplificado do Participante que gerou o arquivo nas linhas 36 e 227.
