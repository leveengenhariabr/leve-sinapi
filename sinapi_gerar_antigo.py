#!/usr/bin/env python3
"""Lê os zips ANTIGOS do SINAPI (até 2024: um zip por estado e por regime, ex.: SINAPI_ref_Insumos_Composicoes_MT_202305_NaoDesonerado.zip)
e grava no MESMO formato dos meses novos (dados/UF/AAAA-MM-SD|CD.json).
Aproveita a divisão oficial da Caixa (mão de obra / material / equipamento / serviços / outros) que já vem no analítico."""
import os, re, json, glob, datetime, unicodedata
import openpyxl

def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn').upper().strip()

def num(v):
    """número no padrão brasileiro ('1.234,56' ou '0,0099') ou já numérico"""
    if v is None or v == '': return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace('%', '')
    if ',' in s: s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except ValueError: return None

def cod(v):
    if v is None or v == '': return None
    if isinstance(v, (int, float)): return str(int(v))
    s = str(v).strip()
    try: return str(int(float(s)))
    except ValueError: return s or None

BALDE = lambda c: ('mat' if c.startswith('MATERIAL') else 'mo' if c.startswith(('MAO DE OBRA', 'ENCARGOS')) else
                   'eq' if c.startswith('EQUIPAMENTO') else 'out')

def achar(pasta, trecho):
    r = [f for f in glob.glob(os.path.join(pasta, '**', '*.xlsx'), recursive=True) if trecho in sem_acento(os.path.basename(f))]
    return sorted(r)[0] if r else None

def linhas(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    return list(wb.worksheets[0].iter_rows(values_only=True))

def ler_insumos(pasta):
    arq = achar(pasta, 'PRECO_REF_INSUMOS')
    if not arq: raise SystemExit('Nao achei a planilha de precos de insumos na pasta ' + pasta)
    classes = {}
    fam = achar(pasta, 'FAMILIA')
    if fam:
        for r in linhas(fam)[3:]:
            if r and len(r) > 6 and cod(r[1]): classes[cod(r[1])] = str(r[6] or '').strip()
    ins = {}
    rows = linhas(arq)
    ini = next((i for i, r in enumerate(rows) if r and sem_acento(r[0]).startswith('CODIGO')), 6) + 1
    for r in rows[ini:]:
        c = cod(r[0]) if r else None
        if not c or not str(r[0]).strip().isdigit(): continue
        cls = classes.get(c, '')
        ins[c] = [c, str(r[1] or '').strip(), str(r[2] or '').strip(), BALDE(sem_acento(cls)), cls or None, num(r[4]), str(r[3] or '').strip() or None]
    return list(ins.values())

def ler_comps(pasta):
    arq = achar(pasta, 'COMPOSICOES_ANALITICO')
    if not arq: raise SystemExit('Nao achei a planilha analitica de composicoes na pasta ' + pasta)
    rows = linhas(arq)
    h = next(i for i, r in enumerate(rows) if r and sem_acento(r[0]) == 'DESCRICAO DA CLASSE')
    cab = [sem_acento(x) for x in rows[h]]
    def col(nome, n=0):
        idx = [i for i, x in enumerate(cab) if x == nome]
        return idx[n] if len(idx) > n else None
    c_g, c_cod, c_desc, c_un, c_tot = col('DESCRICAO DA CLASSE'), col('CODIGO DA COMPOSICAO'), col('DESCRICAO DA COMPOSICAO'), col('UNIDADE'), col('CUSTO TOTAL', 0)
    c_tipo, c_ic, c_coef = col('TIPO ITEM'), col('CODIGO ITEM'), col('COEFICIENTE')
    c_mo, c_mat, c_eq, c_st, c_ou = col('CUSTO MAO DE OBRA'), col('CUSTO MATERIAL'), col('CUSTO EQUIPAMENTO'), col('CUSTO SERVICOS TERCEIROS'), col('CUSTO OUTROS')
    c_as = next((i for i, x in enumerate(cab) if x.startswith('% ATRIBUIDO')), None)
    g = lambda r, i: (r[i] if i is not None and i < len(r) else None)
    comps, ordem = {}, []
    for r in rows[h + 1:]:
        c = cod(g(r, c_cod))
        if not c: continue
        tipo = g(r, c_tipo)
        if not tipo:                               # linha de cabeçalho da composição (traz o custo e a divisão oficial)
            tot = num(g(r, c_tot))
            if tot:
                sp = [num(g(r, c_mat)) or 0.0, num(g(r, c_mo)) or 0.0, num(g(r, c_eq)) or 0.0, (num(g(r, c_st)) or 0.0) + (num(g(r, c_ou)) or 0.0)]
                split = [round(x, 4) for x in sp]
            else: tot, split = None, None
            pa = num(g(r, c_as)); pa = round(pa / 100, 6) if pa is not None else None
            comps[c] = [c, str(g(r, c_desc) or '').strip(), str(g(r, c_un) or '').strip(), str(g(r, c_g) or '').strip(), tot, split, pa, 0, []]
            ordem.append(c)
        elif c in comps:
            comps[c][8].append(['I' if str(tipo).upper().startswith('INSUMO') else 'C', cod(g(r, c_ic)), num(g(r, c_coef))])
    return [comps[c] for c in ordem]

def gerar_regime(pasta, uf, mes, regime, saida_dir):
    insumos, comps = ler_insumos(pasta), ler_comps(pasta)
    dados = {'ref': mes, 'uf': uf, 'regime': regime, 'desonerado': regime == 'CD', 'gerado': datetime.date.today().isoformat(),
             'cab_comp': ['cod', 'descricao', 'un', 'grupo', 'total', '[material,mao_de_obra,equipamento,outros]', 'pct_as', 'incompleta', 'itens[tipo I|C, cod, coef]'],
             'cab_ins': ['cod', 'descricao', 'un', 'balde', 'classificacao', 'preco', 'origem'],
             'insumos': insumos, 'comps': comps}
    os.makedirs(os.path.join(saida_dir, uf), exist_ok=True)
    arq = os.path.join(saida_dir, uf, f'{mes}-{regime}.json')
    with open(arq, 'w', encoding='utf-8') as f: json.dump(dados, f, ensure_ascii=False, separators=(',', ':'))
    sem = sum(1 for c in comps if c[4] is None)
    print(regime, arq, os.path.getsize(arq), 'bytes;', len(comps), 'composicoes;', sem, 'sem custo;', len(insumos), 'insumos (formato antigo)')

if __name__ == '__main__':
    import sys
    pasta, uf, mes, regime = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]   # regime SD ou CD
    gerar_regime(pasta, uf, mes, regime, sys.argv[5] if len(sys.argv) > 5 else 'dados')
