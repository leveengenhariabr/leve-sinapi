#!/usr/bin/env python3
"""Atualiza os arquivos do SINAPI no repositório (qualquer estado, qualquer mês).
Uso:
  python3 sinapi_atualizar.py                        -> MT: baixa o mês mais recente publicado pela Caixa (tenta o mês passado e o retrasado)
  python3 sinapi_atualizar.py --mes 12meses          -> MT: traz de uma vez os últimos 12 meses (pula os que já existem)
  python3 sinapi_atualizar.py --uf GO --mes 2024-05  -> um estado e um mês específicos (ficam "fixos": a limpeza dos 12 meses não apaga)
Também processa os zips da pasta entrada/ (sem precisar de download):
  - meses de 2025 em diante (zip nacional):  SINAPI-2025-03-formato-xlsx.zip           -> MT;  GO_SINAPI-2025-03-formato-xlsx.zip -> GO
  - meses até 2024 (um zip por estado/regime): SINAPI_ref_Insumos_Composicoes_MT_202305_NaoDesonerado[_Retificacao01].zip  e  ..._Desonerado...zip"""
import argparse, datetime, glob, http.cookiejar, io, json, os, re, shutil, sys, time, urllib.request, urllib.error, zipfile
from sinapi_gerar import gerar
from sinapi_gerar_antigo import gerar_regime

UF_PADRAO = 'MT'
RAIZ = os.path.dirname(os.path.abspath(__file__))
DADOS = os.path.join(RAIZ, 'dados')
MANTER_MESES = 12
PRIMEIRO_NOVO = '2025-01'          # a partir daqui o zip é nacional e único; antes disso é um zip por estado e por regime
UA = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36', 'Accept': '*/*'}
UFS = 'AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO'.split()

_cookies = http.cookiejar.CookieJar()
_abridor = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_cookies))

def _baixar_urls(urls):
    # a Caixa às vezes responde com redirecionamento (302) em laço (proteção contra robôs): tenta de novo com pausa
    try: _abridor.open(urllib.request.Request('https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx', headers=UA), timeout=60).read(2000)
    except Exception: pass
    for u in urls:
        for tentativa in range(3):
            try:
                req = urllib.request.Request(u, headers={**UA, 'Referer': 'https://www.caixa.gov.br/site/paginas/downloads.aspx'})
                with _abridor.open(req, timeout=180) as r: dados = r.read()
                if zipfile.is_zipfile(io.BytesIO(dados)):
                    print('baixado:', u, len(dados), 'bytes'); return dados
                print('resposta não é zip:', u); break
            except urllib.error.HTTPError as e:
                print('falhou', e.code, u, f'(tentativa {tentativa + 1})')
                if e.code == 404: break
                time.sleep(6 * (tentativa + 1))
            except Exception as e:
                print('falhou', type(e).__name__, u); time.sleep(4)
    return None

def urls_novo(ref):
    a, m = ref.split('-')
    nomes = [f'SINAPI-{a}-{m}-formato-xlsx.zip', f'SINAPI-{a}-{m}-formato-xlsx_Retificacao01.zip', f'SINAPI-{a}-{m}-formato-xlsx_Retificacao02.zip', f'SINAPI_{a}_{m}_formato_xlsx.zip']
    return ['https://www.caixa.gov.br/Downloads/sinapi-relatorios-mensais/' + n for n in nomes]

def urls_antigo(ref, uf, regime):
    a, m = ref.split('-'); reg = 'Desonerado' if regime == 'CD' else 'NaoDesonerado'
    base = f'https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{uf.lower()}/SINAPI_ref_Insumos_Composicoes_{uf}_{a}{m}_{reg}'
    return [base + suf + '.zip' for suf in ('', '_Retificacao01', '_Retificacao02')]

# ---------- índice ----------
def ler_indice(uf):
    p = os.path.join(DADOS, uf, 'indice.json')
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {'uf': uf, 'meses': []}

def tem_arquivo(uf, ref, regime):
    return os.path.exists(os.path.join(DADOS, uf, f'{ref}-{regime}.json'))

def mes_ja_existe(uf, ref):
    return tem_arquivo(uf, ref, 'SD') and tem_arquivo(uf, ref, 'CD')

def atualizar_indice(uf, ref, fixo=False):
    if not mes_ja_existe(uf, ref):                       # só entra no índice quando os dois regimes estão prontos
        print('índice: aguardando o outro regime de', ref, 'para', uf); return
    ind = ler_indice(uf)
    atual = {m['ref']: m for m in ind['meses']}
    atual[ref] = {'ref': ref, **({'fixo': True} if fixo or atual.get(ref, {}).get('fixo') else {})}
    meses = sorted(atual.values(), key=lambda m: m['ref'], reverse=True)
    recentes = {m['ref'] for m in meses[:MANTER_MESES]}
    manter = []
    for m in meses:
        if m['ref'] in recentes or m.get('fixo'): manter.append(m)
        else:
            for reg in ('SD', 'CD'):
                f = os.path.join(DADOS, uf, f"{m['ref']}-{reg}.json")
                if os.path.exists(f): os.remove(f); print('removido (mais de 12 meses):', f)
    ind = {'uf': uf, 'atualizado': datetime.date.today().isoformat(), 'meses': manter}
    os.makedirs(os.path.join(DADOS, uf), exist_ok=True)
    with open(os.path.join(DADOS, uf, 'indice.json'), 'w', encoding='utf-8') as f: json.dump(ind, f, ensure_ascii=False, indent=1)
    print('índice', uf + ':', [m['ref'] for m in manter])

def indice_geral():
    """dados/estados.json: quais estados têm dados (o app usa para montar o seletor)"""
    ufs = sorted(d for d in os.listdir(DADOS) if os.path.isdir(os.path.join(DADOS, d)) and os.path.exists(os.path.join(DADOS, d, 'indice.json'))) if os.path.isdir(DADOS) else []
    with open(os.path.join(DADOS, 'estados.json'), 'w', encoding='utf-8') as f: json.dump({'ufs': ufs}, f)

# ---------- processamento ----------
def processar_novo(dados, ref, uf, fixo):
    tmp = os.path.join(RAIZ, '_tmp'); shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        # a planilha "Referência" é a maior do zip (as outras são manutenções, famílias e mão de obra); o nome pode vir com acento quebrado
        xl = [i for i in z.infolist() if i.filename.lower().endswith('.xlsx')]
        if not xl: sys.exit('ERRO: o zip não tem nenhuma planilha .xlsx.')
        alvo = max(xl, key=lambda i: i.file_size)
        caminho = os.path.join(tmp, 'referencia.xlsx')
        with z.open(alvo) as origem, open(caminho, 'wb') as destino: shutil.copyfileobj(origem, destino)
    gerar(caminho, uf, DADOS, ref)
    shutil.rmtree(tmp, ignore_errors=True)
    atualizar_indice(uf, ref, fixo)

def processar_antigo(dados, ref, uf, regime, fixo):
    tmp = os.path.join(RAIZ, '_tmp_' + regime); shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(io.BytesIO(dados)) as z: z.extractall(tmp)
    gerar_regime(tmp, uf, ref, regime, DADOS)
    shutil.rmtree(tmp, ignore_errors=True)
    atualizar_indice(uf, ref, fixo)

def obter_mes(uf, ref, fixo):
    """baixa e processa um mês (formato novo ou antigo conforme o ano). Devolve True se conseguiu."""
    if ref >= PRIMEIRO_NOVO:
        d = _baixar_urls(urls_novo(ref))
        if not d: return False
        processar_novo(d, ref, uf, fixo); return True
    ok = True
    for regime in ('SD', 'CD'):
        if tem_arquivo(uf, ref, regime): continue
        d = _baixar_urls(urls_antigo(ref, uf, regime))
        if not d: ok = False; continue
        processar_antigo(d, ref, uf, regime, fixo)
    return ok

def prim_recente(uf):
    ms = ler_indice(uf)['meses']
    return ms[min(len(ms), MANTER_MESES) - 1]['ref'] if len(ms) >= MANTER_MESES else '0000-00'

def processar_entrada():
    feito = False
    for z in sorted(glob.glob(os.path.join(RAIZ, 'entrada', '*.zip'))):
        nome = os.path.basename(z)
        mo = re.search(r'Insumos_Composicoes_([A-Z]{2})_(\d{4})(\d{2})_(Nao)?Desonerado', nome)
        print('processando envio manual:', nome)
        if mo:                                              # formato antigo: um zip por estado e regime
            uf, ref, regime = mo.group(1), f'{mo.group(2)}-{mo.group(3)}', ('SD' if mo.group(4) else 'CD')
            processar_antigo(open(z, 'rb').read(), ref, uf, regime, fixo=True)
        else:
            m = re.search(r'(\d{4})[-_](\d{2})', nome)
            if not m: print('ignorado (sem AAAA-MM no nome):', nome); continue
            ref = f'{m.group(1)}-{m.group(2)}'
            pu = re.match(r'^([A-Z]{2})[_-]', nome)
            uf = pu.group(1) if pu and pu.group(1) in UFS else UF_PADRAO
            processar_novo(open(z, 'rb').read(), ref, uf, fixo=(uf != UF_PADRAO or ref < prim_recente(uf)))
        os.remove(z); feito = True
    return feito

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--mes'); ap.add_argument('--uf', default=UF_PADRAO); a = ap.parse_args()
    uf = a.uf.upper()
    if uf not in UFS: sys.exit('UF inválida: ' + uf)
    os.makedirs(os.path.join(DADOS, uf), exist_ok=True)
    feito = processar_entrada()
    hoje = datetime.date.today()
    def meses_atras(n):
        for k in range(1, n + 1):
            a_, m_ = hoje.year, hoje.month - k
            while m_ < 1: m_ += 12; a_ -= 1
            yield f'{a_}-{m_:02d}'
    if a.mes == '12meses':                                 # os últimos 12 meses publicados, de uma vez
        falhas = 0
        for ref in meses_atras(MANTER_MESES):
            if mes_ja_existe(uf, ref): print('já existe:', ref); continue
            ok = obter_mes(uf, ref, fixo=(uf != UF_PADRAO)); time.sleep(3)
            if not ok: falhas += 1; print('NÃO consegui baixar', ref)
        if falhas: print(f'{falhas} mês(es) não baixaram. Baixe o zip no site da Caixa e envie para a pasta "entrada".')
    elif a.mes:
        if not re.fullmatch(r'\d{4}-\d{2}', a.mes): sys.exit('Mês inválido. Use AAAA-MM, ex.: 2025-03')
        if mes_ja_existe(uf, a.mes): print('já existe:', uf, a.mes)
        elif not obter_mes(uf, a.mes, fixo=True):
            print(f'ERRO: não consegui baixar {uf} {a.mes} da Caixa (o site às vezes bloqueia o download automático). Baixe o zip no site da Caixa e envie para a pasta "entrada" do repositório.')
    elif not feito:
        for ref in meses_atras(2):                         # mês passado, depois o retrasado
            if mes_ja_existe(uf, ref): print('já existe:', ref); break
            if obter_mes(uf, ref, fixo=False): break
        else: print('Nada novo publicado pela Caixa ainda (ou o site bloqueou o download). Se precisar, envie o zip pela pasta "entrada".')
    indice_geral()

if __name__ == '__main__':
    main()
