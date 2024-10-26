import pandas as pd
from datetime import datetime

# Carregar o arquivo Excel com pandas
def carregar_dados_excel(caminho_excel):
    df = pd.read_excel(caminho_excel)
    return df

# Função para criar o arquivo OFX manualmente
def criar_arquivo_ofx(transacoes, arquivo_saida):
    # Cabeçalho do arquivo OFX
    ofx_content = """
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
  <SIGNONMSGSRSV1>
    <SONRS>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <DTSERVER>{data_atual}</DTSERVER>
      <LANGUAGE>POR</LANGUAGE>
    </SONRS>
  </SIGNONMSGSRSV1>
  <BANKMSGSRSV1>
    <STMTTRNRS>
      <TRNUID>1</TRNUID>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <STMTRS>
        <CURDEF>BRL</CURDEF>
        <BANKACCTFROM>
          <BANKID>001</BANKID>
          <ACCTID>123456</ACCTID>
          <ACCTTYPE>CHECKING</ACCTTYPE>
        </BANKACCTFROM>
        <BANKTRANLIST>
          <DTSTART>{data_inicio}</DTSTART>
          <DTEND>{data_fim}</DTEND>
    """.format(
        data_atual=datetime.now().strftime('%Y%m%d%H%M%S'),
        data_inicio=transacoes['Data'].min().strftime('%Y%m%d'),
        data_fim=transacoes['Data'].max().strftime('%Y%m%d')
    )

    # Adicionar transações
    for idx, transacao in transacoes.iterrows():
        tipo_transacao = 'DEBIT' if transacao['Valor'] < 0 else 'CREDIT'
        ofx_content += """
          <STMTTRN>
            <TRNTYPE>{tipo}</TRNTYPE>
            <DTPOSTED>{data}</DTPOSTED>
            <TRNAMT>{valor}</TRNAMT>
            <FITID>ID{idx}</FITID>
            <MEMO>{descricao}</MEMO>
          </STMTTRN>
        """.format(
            tipo=tipo_transacao,
            data=transacao['Data'].strftime('%Y%m%d'),
            valor=transacao['Valor'],
            idx=idx,
            descricao=transacao['Descrição']
        )

    # Fechar o conteúdo OFX
    ofx_content += """
        </BANKTRANLIST>
      </STMTRS>
    </STMTTRNRS>
  </BANKMSGSRSV1>
</OFX>
    """

    # Escrever o arquivo OFX
    with open(arquivo_saida, "w") as f:
        f.write(ofx_content)

# Exemplo de como usar
if __name__ == "__main__":
    caminho_excel = "02Fevereiro.xlsx"
    arquivo_saida_ofx = "Fevereiro.ofx"

    # Carregar as transações do Excel
    df_transacoes = carregar_dados_excel(caminho_excel)

    # Criar arquivo OFX
    criar_arquivo_ofx(df_transacoes, arquivo_saida_ofx)

    print(f"Arquivo {arquivo_saida_ofx} criado com sucesso!")
