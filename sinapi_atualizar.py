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

def atualizar_indice_fonte(pasta, uf, ref):
    """índice das outras bases (GOINFRA, SICRO…): dados/<pasta>/<uf>/indice.json; só entra o mês com SD e CD prontos; guarda todos os meses"""
    d = os.path.join(DADOS, pasta, uf)
    if not all(os.path.exists(os.path.join(d, f'{ref}-{r}.json')) for r in ('SD', 'CD')):
        print('índice:', pasta, uf, ref, '- aguardando o outro regime'); return
    pi = os.path.join(d, 'indice.json')
    ind = json.load(open(pi, encoding='utf-8')) if os.path.exists(pi) else {'uf': uf, 'meses': []}
    refs = {m['ref'] for m in ind['meses']} | {ref}
    ind = {'uf': uf, 'atualizado': datetime.date.today().isoformat(), 'meses': [{'ref': r} for r in sorted(refs, reverse=True)]}
    with open(pi, 'w', encoding='utf-8') as f: json.dump(ind, f, ensure_ascii=False, indent=1)
    print('índice', pasta, uf + ':', [m['ref'] for m in ind['meses']])

def processar_planilhas():
    """GOINFRA (3 planilhas por regime: composições, materiais, mão de obra) e SICRO (relatórios do DNIT por estado/mês):
    os arquivos são reconhecidos pelo CONTEÚDO, não pelo nome. Só processa quando o conjunto está completo; o resto espera na pasta."""
    arqs = sorted(glob.glob(os.path.join(RAIZ, 'entrada', '*.xls')) + glob.glob(os.path.join(RAIZ, 'entrada', '*.xlsx')))
    if not arqs: return False
    import sinapi_gerar_goinfra as G, sinapi_gerar_sicro as S
    goi, sic = {}, {}
    for a in arqs:
        nome = os.path.basename(a)
        try:
            info = G.identificar(G.linhas(a))
            if info['tipo'] and info['ref'] and info['regime']:
                goi.setdefault((info['ref'], info['regime']), {})[info['tipo']] = a; continue
        except Exception as e: print('não é GOINFRA:', nome, '-', type(e).__name__)
        try:
            rows = S._linhas(a); t = S.identificar(rows)
        except Exception as e: print('ignorado (não consegui ler):', nome, e); continue
        pu = re.match(r'^([A-Z]{2})[_\s-](\d{2})-(\d{4})', nome)
        if not t:
            if pu and re.search(r'relat.rio', nome.lower()): sic.setdefault((pu.group(1), f'{pu.group(3)}-{pu.group(2)}'), {'arquivos': []})['arquivos'].append(a)   # relatório do SICRO que não usamos: sai junto
            else: print('ignorado (formato desconhecido):', nome)
            continue
        if not pu: print('SICRO sem “UF_MM-AAAA” no nome, ignorado:', nome); continue
        cd = bool(re.search(r'com[ _]desonera', nome.lower()))
        chave = (pu.group(1), f'{pu.group(3)}-{pu.group(2)}')
        grupo = sic.setdefault(chave, {'arquivos': []}); grupo['arquivos'].append(a)
        grupo['comp' if t == 'comp' else 'mat' if t == 'mat' else 'mo_an' if (t == 'mo_an' and not cd) else (t + '_' + ('cd' if cd else 'sd')) if t in ('eq', 'mo') else '_x'] = rows
    feito = False
    por_mes = {}
    for (ref, reg), g in goi.items():
        if all(k in g for k in ('comp', 'mat', 'mdo')):
            r, rg = G.gerar_goinfra(g, DADOS); por_mes.setdefault(r, []).append(rg)
            for k in g.values(): os.remove(k)
            feito = True
        else: print('GOINFRA', ref, reg, 'incompleto, faltam:', [k for k in ('comp', 'mat', 'mdo') if k not in g])
    for r in por_mes: atualizar_indice_fonte('goinfra', 'GO', r)
    for (uf, ref), g in sic.items():
        falta = [k for k in ('comp', 'mat', 'eq_sd', 'eq_cd', 'mo_sd', 'mo_cd') if k not in g]
        if falta: print('SICRO', uf, ref, 'incompleto, faltam:', falta); continue
        r, _ = S.gerar_sicro(g, DADOS, uf=uf)
        atualizar_indice_fonte('sicro', uf, r)
        for a in g['arquivos']: os.remove(a)
        feito = True
    return feito

# ---------- download automático da GOINFRA e do SICRO (se o site bloquear, vale o envio manual para "entrada") ----------
import urllib.parse
GOINFRA_PAGINAS = ['https://goias.gov.br/goinfra/tabela-de-composicao114/', 'https://goias.gov.br/goinfra/custo-referencial-de-materiais120/',
                   'https://goias.gov.br/goinfra/custo-referencial-de-mao-de-obra-121/']
SICRO_ESTADOS = {'MT': 'mato-grosso', 'GO': 'goias'}
SICRO_BASE = 'https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro/centro-oeste'
SICRO_MESES = {1: 'janeiro', 4: 'abril', 7: 'julho', 10: 'outubro'}

def _get(url):
    req = urllib.request.Request(urllib.parse.quote(url, safe=":/%?=&"), headers=UA)
    with _abridor.open(req, timeout=180) as r: return r.read()

def baixar_goinfra():
    """acha nas páginas da GOINFRA as planilhas (.xls) da tabela Obras Civis mais recente (pastas T###_S e T###_C) e salva em entrada/"""
    dest = os.path.join(RAIZ, 'entrada'); os.makedirs(dest, exist_ok=True)
    pasta_est = os.path.join(DADOS, 'goinfra', 'GO'); os.makedirs(pasta_est, exist_ok=True)
    pe = os.path.join(pasta_est, 'baixados.json')
    ja = json.load(open(pe)) if os.path.exists(pe) else []
    links = set()
    for pag in GOINFRA_PAGINAS:
        try: html = _get(pag).decode('utf-8', 'replace')
        except Exception as e: print('GOINFRA: não abriu', pag, '-', e); continue
        for m in re.findall(r'https?://cdn\.goinfra\.go\.gov\.br/[^"\'<>]+?\.xls', html): links.add(urllib.parse.unquote(m))
    pastas = {}
    for l in links:
        m = re.search(r'/Obras Civis/(T(\d+)_([SC]))/', l)
        if m: pastas.setdefault(int(m.group(2)), {}).setdefault(m.group(3), []).append(l)
    # a tabela mais recente tem uma pasta S e uma C; os números ficam lado a lado (ex.: T348_C e T349_S)
    todas = sorted(pastas)
    if not todas: print('GOINFRA: nenhum link .xls encontrado (site mudou ou bloqueou). Envie as planilhas para a pasta entrada.'); return False
    alvo = [n for n in todas[-2:]]
    if all(str(n) in ja for n in alvo): print('GOINFRA: já baixado', alvo); return False
    baixou = 0
    for n in alvo:
        for reg, lista in pastas[n].items():
            for l in lista:
                nome = f'GOINFRA_T{n}_{reg}_' + re.sub(r'[^A-Za-z0-9._-]+', '_', os.path.basename(l))
                try:
                    dados = _get(l)
                    if dados[:4] in (b'<!DO', b'<htm', b'<HTM'): raise ValueError('veio uma página em vez da planilha')
                    open(os.path.join(dest, nome), 'wb').write(dados); baixou += 1; print('GOINFRA baixado:', nome, len(dados))
                except Exception as e: print('GOINFRA: falhou', l, '-', e)
    if baixou: json.dump(sorted(set(ja) | {str(n) for n in alvo}), open(pe, 'w'))
    return baixou > 0

def baixar_sicro(hoje=None, so_mes=None):
    """SICRO sai em jan/abr/jul/out por estado: tenta o trimestre mais recente (e o anterior) de MT e GO; guarda o .7z, abre e salva as planilhas em entrada/"""
    hoje = hoje or datetime.date.today(); dest = os.path.join(RAIZ, 'entrada'); os.makedirs(dest, exist_ok=True)
    cand = []; a, m = hoje.year, hoje.month
    if so_mes: cand = [so_mes]
    while len(cand) < 2 and not so_mes:
        mm = max(k for k in SICRO_MESES if k <= m) if m >= 1 and any(k <= m for k in SICRO_MESES) else None
        if mm is None: a, m = a - 1, 12; continue
        cand.append((a, mm)); m = mm - 1
        if m < 1: a, m = a - 1, 12
    feito = False
    for uf, est in SICRO_ESTADOS.items():
        for a, mm in cand:
            ref = f'{a}-{mm:02d}'
            if os.path.exists(os.path.join(DADOS, 'sicro', uf, f'{ref}-SD.json')): print('SICRO já existe:', uf, ref); break
            url = f"{SICRO_BASE}/{est}/{a}/{SICRO_MESES[mm]}/{uf.lower()}-{mm:02d}-{a}.7z"
            try:
                dados = _get(url)
                if dados[:2] != b'7z': raise ValueError('não é um arquivo .7z')
            except Exception as e: print('SICRO: não consegui', uf, ref, '-', e); continue
            try:
                import py7zr, tempfile
                with tempfile.TemporaryDirectory() as t:
                    pz = os.path.join(t, 'x.7z'); open(pz, 'wb').write(dados)
                    with py7zr.SevenZipFile(pz) as z: z.extractall(os.path.join(t, 'x'))
                    for raiz, _, arqs in os.walk(os.path.join(t, 'x')):
                        for f in arqs:
                            if f.lower().endswith('.xlsx'): shutil.copy(os.path.join(raiz, f), os.path.join(dest, f if re.match(r'^[A-Z]{2}[_ -]\d{2}-\d{4}', f) else f'{uf}_{mm:02d}-{a}_' + f))
                print('SICRO baixado:', uf, ref); feito = True; break
            except Exception as e: print('SICRO: falhou ao abrir o .7z', uf, ref, '-', e)
    return feito

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--mes'); ap.add_argument('--uf', default=UF_PADRAO); a = ap.parse_args()
    uf = a.uf.upper()
    if uf not in UFS: sys.exit('UF inválida: ' + uf)
    os.makedirs(os.path.join(DADOS, uf), exist_ok=True)
    feito = processar_entrada()
    forcado = None                                         # --mes 2026-04 também pede ao SICRO aquele trimestre (jan/abr/jul/out)
    if a.mes and re.fullmatch(r'\d{4}-(01|04|07|10)', a.mes): forcado = (int(a.mes[:4]), int(a.mes[5:]))
    for baixar in (baixar_goinfra, lambda: baixar_sicro(so_mes=forcado)):          # GOINFRA e SICRO: tenta baixar sozinho; falhou, não derruba o resto
        try: baixar()
        except Exception as e: print('download automático falhou:', e)
    processar_planilhas()
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
