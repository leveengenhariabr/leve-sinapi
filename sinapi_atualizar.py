#!/usr/bin/env python3
"""Atualiza os arquivos do SINAPI de Mato Grosso no repositório.
Uso:
  python3 sinapi_atualizar.py                  -> baixa o mês mais recente publicado pela Caixa (tenta o mês passado e o retrasado)
  python3 sinapi_atualizar.py --mes 2025-03    -> baixa e guarda um mês específico (fica "fixo": a limpeza dos 12 meses não apaga)
Também processa qualquer zip da pasta entrada/ (nome com AAAA-MM, ex.: SINAPI-2026-10-formato-xlsx.zip), sem precisar de download."""
import argparse, datetime, glob, io, json, os, re, shutil, sys, urllib.request, urllib.error, zipfile
from sinapi_gerar import gerar

UF = 'MT'
RAIZ = os.path.dirname(os.path.abspath(__file__))
DADOS = os.path.join(RAIZ, 'dados')
MANTER_MESES = 12
UA = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36', 'Accept': '*/*'}

def urls_candidatas(ref):
    a, m = ref.split('-')
    nomes = [f'SINAPI-{a}-{m}-formato-xlsx.zip']
    bases = ['https://www.caixa.gov.br/Downloads/sinapi-relatorios-mensais/', 'https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-mt/']
    return [b + n for b in bases for n in nomes]

def baixar(ref):
    for u in urls_candidatas(ref):
        try:
            req = urllib.request.Request(u, headers=UA)
            with urllib.request.urlopen(req, timeout=180) as r:
                dados = r.read()
            if zipfile.is_zipfile(io.BytesIO(dados)):
                print('baixado:', u, len(dados), 'bytes'); return dados
            print('resposta não é zip:', u)
        except urllib.error.HTTPError as e:
            print('falhou', e.code, u)
        except Exception as e:
            print('falhou', type(e).__name__, u)
    return None

def processar_zip(dados, ref, fixo):
    tmp = os.path.join(RAIZ, '_tmp'); shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        # a planilha "Referência" é a maior do zip (as outras são manutenções, famílias e mão de obra); o nome pode vir com acento quebrado
        xl = [i for i in z.infolist() if i.filename.lower().endswith('.xlsx')]
        if not xl: sys.exit('ERRO: o zip não tem nenhuma planilha .xlsx.')
        alvo = max(xl, key=lambda i: i.file_size)
        caminho = os.path.join(tmp, 'referencia.xlsx')
        with z.open(alvo) as origem, open(caminho, 'wb') as destino: shutil.copyfileobj(origem, destino)
    gerar(caminho, UF, DADOS, ref)
    shutil.rmtree(tmp, ignore_errors=True)
    atualizar_indice(ref, fixo)

def ler_indice():
    p = os.path.join(DADOS, UF, 'indice.json')
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {'uf': UF, 'meses': []}

def atualizar_indice(ref, fixo=False):
    ind = ler_indice()
    atual = {m['ref']: m for m in ind['meses']}
    atual[ref] = {'ref': ref, **({'fixo': True} if fixo or atual.get(ref, {}).get('fixo') else {})}
    meses = sorted(atual.values(), key=lambda m: m['ref'], reverse=True)
    recentes = {m['ref'] for m in meses[:MANTER_MESES]}
    manter = []
    for m in meses:
        if m['ref'] in recentes or m.get('fixo'): manter.append(m)
        else:
            for reg in ('SD', 'CD'):
                f = os.path.join(DADOS, UF, f"{m['ref']}-{reg}.json")
                if os.path.exists(f): os.remove(f); print('removido (mais de 12 meses):', f)
    ind = {'uf': UF, 'atualizado': datetime.date.today().isoformat(), 'meses': manter}
    with open(os.path.join(DADOS, UF, 'indice.json'), 'w', encoding='utf-8') as f: json.dump(ind, f, ensure_ascii=False, indent=1)
    print('índice:', [m['ref'] for m in manter])

def mes_ja_existe(ref):
    return all(os.path.exists(os.path.join(DADOS, UF, f'{ref}-{r}.json')) for r in ('SD', 'CD'))

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--mes'); a = ap.parse_args()
    os.makedirs(os.path.join(DADOS, UF), exist_ok=True)
    feito = False
    for z in sorted(glob.glob(os.path.join(RAIZ, 'entrada', '*.zip'))):   # envio manual
        m = re.search(r'(\d{4})[-_](\d{2})', os.path.basename(z))
        if not m: print('ignorado (sem AAAA-MM no nome):', z); continue
        ref = f'{m.group(1)}-{m.group(2)}'
        print('processando envio manual:', z)
        processar_zip(open(z, 'rb').read(), ref, fixo=bool(a.mes) or ref < prim_recente())
        os.remove(z); feito = True
    if a.mes:
        if not re.fullmatch(r'\d{4}-\d{2}', a.mes): sys.exit('Mês inválido. Use AAAA-MM, ex.: 2025-03')
        if not mes_ja_existe(a.mes) or not feito:
            d = baixar(a.mes)
            if not d: sys.exit(f'ERRO: não consegui baixar {a.mes} da Caixa. Baixe o zip no site da Caixa e envie para a pasta "entrada" do repositório.')
            processar_zip(d, a.mes, fixo=True)
        return
    if feito: return
    hoje = datetime.date.today()
    for k in (1, 2):                                       # mês passado, depois o retrasado
        a_, m_ = hoje.year, hoje.month - k
        while m_ < 1: m_ += 12; a_ -= 1
        ref = f'{a_}-{m_:02d}'
        if mes_ja_existe(ref): print('já existe:', ref); return
        d = baixar(ref)
        if d: processar_zip(d, ref, fixo=False); return
    print('Nada novo publicado pela Caixa ainda (ou o site bloqueou o download). Se precisar, envie o zip pela pasta "entrada".')

def prim_recente():
    ms = ler_indice()['meses']
    return ms[min(len(ms), MANTER_MESES) - 1]['ref'] if len(ms) >= MANTER_MESES else '0000-00'

if __name__ == '__main__':
    main()
