import pandas
import requests
import gspread
import datetime
from google.colab import auth
from google.auth import default

auth.authenticate_user()
creds, _ = default()
gc = gspread.authorize(creds)

# Consulta de ativo no Opções.net.br
def listar_opcoes (ativo_obj, vencimento):
    url=f'https://opcoes.net.br/listaopcoes/completa?idAcao={ativo_obj}&listarVencimentos=true&cotacoes=true&vencimentos={vencimento}'
    r = requests.get(url).json()
    l = [[ativo_obj, vencimento, i[0].split('_')[0], i[2], i[3], i[5], i[8]] for i in r['data']['cotacoesOpcoes']]
    return pandas.DataFrame(l, columns = ['ativo_obj', 'vencimento', 'ativo', 'tipo', 'modelo', 'strike', 'preco'])

def listar_tudo (ativo_obj):
    url=f'https://opcoes.net.br/listaopcoes/completa?idAcao={ativo_obj}&listarVencimentos=true&cotacoes=true'
    r = requests.get(url).json()
    vencimentos = [i['value'] for i in r['data']['vencimentos']]
    df=pandas.concat([listar_opcoes(ativo_obj, vencimento) for vencimento in vencimentos])
    return df

# Varre a lista de Ações e busca a opcao que estiver com strike a cerca de 5% do valor atual da ação
s = gc.open('XM');
page = s.worksheet('template');

venc_30d = '2023-08-18';

xm = pandas.DataFrame(page.get_all_records());

for i, acao in xm.iterrows():
    print (f'{acao.Nome} - {acao.Cotacao}');
    if acao.Nome != "":
      acao.Cotacao = float(acao.Cotacao.replace(',', '.'));
      opcoes=listar_opcoes(acao.Nome, venc_30d);
      for j, opcao in opcoes.iterrows():
        perc = ((opcao.strike - acao.Cotacao)/acao.Cotacao)*100;
        opcao = opcao.fillna(0);
        if opcao.preco != 0:
          if abs(perc) < 2:
            if opcao.tipo == 'CALL':
              acao.CallName = opcao.ativo;
              acao.CallPremio = opcao.preco;
              acao.CallStrike = opcao.strike;
              acao.CallMargem = (opcao.preco/acao.Cotacao);
            else:
              acao.PutName = opcao.ativo;
              acao.PutPremio = opcao.preco;
              acao.PutStrike = opcao.strike;
              acao.PutMargem = (opcao.preco/acao.Cotacao);

timeStamp = str(datetime.datetime.strptime(datetime.datetime.now().strftime("%A %d %B %y %I:%M"),"%A %d %B %y %I:%M"));
novo = s.duplicate_sheet(page.id, insert_sheet_index=None, new_sheet_id=None, new_sheet_name=timeStamp);
novo.update([xm.columns.values.tolist()] + xm.values.tolist());
