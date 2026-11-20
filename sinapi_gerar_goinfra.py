#!/usr/bin/env python3
"""Lê os relatórios da GOINFRA (Obras Civis, sistema SCO) e grava no MESMO formato do SINAPI (dados/goinfra/GO/AAAA-MM-SD|CD.json).
São 3 arquivos por regime (com e sem desoneração = 6 no total):
  • Relatório de Composição do Serviço  (analítico: mão de obra, materiais... por serviço)
  • Custo Referencial de Materiais
  • Custo Referencial de Mão de Obra (horistas e mensalistas, com e sem encargos)
Aceita .xls (precisa do pacote xlrd) e .xlsx (openpyxl). O mês, o regime e o tipo vêm do cabeçalho de cada arquivo."""
import os, re, json, datetime, unicodedata

def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s or '')) if unicodedata.category(c) != 'Mn').upper().strip()

def texto(v):
    if v is None: return ''
    if isinstance(v, float): return str(int(v)) if v == int(v) else repr(v)
    return str(v).strip()

def linhas(caminho):
    """todas as linhas da primeira planilha, como listas de texto"""
    if caminho.lower().endswith('.xls'):
        import xlrd
        wb = xlrd.open_workbook(caminho)
        sh = wb.sheet_by_index(0)
        return [[texto(sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(sh.nrows)]
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    return [[texto(x) for x in r] for r in wb.worksheets[0].iter_rows(values_only=True)]

def num(v):
    s = texto(v).replace('%', '').replace('R$', '').strip()
    if not s: return None
    if ',' in s: s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except ValueError: return None

def cod(v, prefixo=''):
    s = texto(v)
    if not s: return None
    s = s.strip()
    return prefixo + (str(int(s)) if s.isdigit() else s)

def identificar(rows):
    """devolve {tipo: comp|mat|mdo, ref: AAAA-MM, regime: SD|CD, tabela: texto} lendo o cabeçalho do arquivo"""
    topo = [' '.join(c for c in r if c) for r in rows[:12]]
    tit = ' '.join(topo)
    t = sem_acento(tit)
    tipo = 'comp' if 'COMPOSICAO DO SERVICO' in t else 'mat' if 'REFERENCIAL DE MATERIAIS' in t else 'mdo' if 'REFERENCIAL DE MAO DE OBRA' in t else None
    m = re.search(r'DATA BASE:?\s*(\d{2})/(\d{2})/(\d{4})', t)
    ref = f'{m.group(3)}-{m.group(2)}' if m else None
    regime = 'SD' if 'SEM DESONERACAO' in t else 'CD' if 'COM DESONERACAO' in t else None
    fam = 'civil' if 'OBRAS CIVIS' in t else 'rod' if 'RODOVIAR' in t else None
    mt = re.search(r'TABELA[^0-9]{0,20}(\d+)', t)
    return {'tipo': tipo, 'ref': ref, 'regime': regime, 'familia': fam, 'tabela': mt.group(1) if mt else None}

def ler_materiais(rows):
    ins, secao = {}, ''
    for r in rows:
        nz = [x for x in r if x]
        if not nz: continue
        if re.match(r'^\d\d - ', nz[0]) and len(nz) == 1: secao = nz[0].split(' - ', 1)[1].strip(); continue
        if nz[0].startswith(('Código', 'Tabela', 'GOINFRA', 'Custo Ref', 'SCO ')): continue
        c = cod(r[0])
        if not c or len(nz) < 4: continue
        un = next((x for x in r[1:] if x and x in nz[-3:-1]), '')
        p = num(nz[-1]); d = nz[1]
        un = nz[-2]
        if p is None: continue
        ins[c] = [c, d, un, 'mat', secao or None, p, None]
    return ins

def ler_mao_de_obra(rows):
    ins, secao = {}, ''
    for r in rows:
        nz = [x for x in r if x]
        if not nz: continue
        if re.match(r'^\d\d - ', nz[0].strip()) and len(nz) == 1: secao = nz[0].split(' - ', 1)[1].strip(); continue
        if nz[0].startswith(('Código', 'Tabela', 'GOINFRA', 'Custo Ref', 'SCO ')) or len(nz) < 6: continue
        c = cod(r[0], 'MO')
        if not c or not texto(r[0]).isdigit(): continue
        enc, sal, com = num(nz[-3]), num(nz[-2]), num(nz[-1])
        if com is None: continue
        ins[c] = [c, nz[1], nz[2], 'mo', secao or None, com, None, [sal, enc]]    # 8º campo: salário sem encargos e % de encargos
    return ins

SECAO_BALDE = {'A': 'eq', 'B': 'mo', 'C': 'mat', 'D': 'out', 'E': 'out'}
def ler_composicoes(rows, mat, mdo):
    comps, ordem = {}, []
    novos = {}
    atual = None; secao = None; cols = {}
    for r in rows:
        nz = [x for x in r if x]
        if not nz: continue
        a = nz[0]
        if a.startswith('Serviço:'):
            m = re.match(r'^Serviço:\s*(\S+)\s*-\s*(.*)$', a)
            un = next((x.split(':', 1)[1].strip() for x in nz[1:] if x.startswith('Unidade:')), '')
            c = m.group(1) if m else a
            atual = [c, (m.group(2) if m else a).strip(), un, 'Grupo ' + c[:2], None, None, None, 0, []]
            comps[c] = atual; ordem.append(c); secao = None; novos.setdefault(c, {})
            atual.append({'A': 0.0, 'B': 0.0, 'C': 0.0, 'D': 0.0, 'E': 0.0})   # [9] = totais por seção
            continue
        if atual is None: continue
        if a.startswith('Código Auxiliar'):
            h = next((x for x in nz[1:] if re.match(r'^\([A-E]\)', x)), None)
            secao = h[1] if h else None
            cols = {x: j for j, x in enumerate(r) if x}
            cols['_desc'] = next((j for j, x in enumerate(r) if x and re.match(r'^\([A-E]\)', x)), 3)
            continue
        mt = re.match(r'^\(([A-E])\) Total:', a)
        if mt: atual[9][mt.group(1)] = num(nz[-1]) or 0.0; continue
        if a.startswith('Custo direto total'): atual[4] = num(nz[-1]); continue
        if a.startswith('Preço unitário total'): atual[4] = num(nz[-1]); continue
        if secao and a[:1].isdigit() or (secao and re.match(r'^[A-Za-z]?\d+$', a)):
            g = lambda nome: (r[cols[nome]] if nome in cols and cols[nome] < len(r) else '')
            consumo = num(g('Consumo'))
            if consumo is None: continue
            desc = r[cols['_desc']] if cols['_desc'] < len(r) else ''
            if secao == 'B':
                c = cod(a, 'MO'); preco = num(g('Sal/Hora')); un = 'h'
                if c not in mdo: mdo[c] = [c, desc, 'h', 'mo', 'MÃO DE OBRA', preco, None, [num(g('Eq. Salarial')), num(g('Encargos(%)'))]]
            else:
                c = cod(a); preco = num(g('Valor unitário')) or num(g('Custo Horário')); un = g('Unidade')
                if c not in mat: mat[c] = [c, desc, un, SECAO_BALDE.get(secao, 'out'), None, preco, None]
            atual[8].append(['I', c, consumo])
    out = []
    for c in ordem:
        k = comps[c]; t = k[9]
        total = k[4]
        if total is None or total == 0:
            k[4] = None; k[5] = None
        else:
            mat_, mo_, eq_ = t['C'], t['B'], t['A']
            outros = max(0.0, round(total - mat_ - mo_ - eq_, 4))
            k[5] = [round(mat_, 4), round(mo_, 4), round(eq_, 4), outros]
        out.append(k[:9])
    return out

def gerar_goinfra(arqs, saida_dir, uf='GO', pasta='goinfra'):
    """arqs = {'comp': caminho, 'mat': caminho, 'mdo': caminho} do MESMO mês e regime. Devolve (ref, regime)."""
    info = {t: identificar(linhas(arqs[t])) for t in ('comp', 'mat', 'mdo')}
    refs = {(i['ref'], i['regime']) for i in info.values()}
    if len(refs) != 1 or None in next(iter(refs)): raise SystemExit(f'Os 3 arquivos não são do mesmo mês/regime: {info}')
    ref, regime = next(iter(refs))
    mat = ler_materiais(linhas(arqs['mat'])); mdo = ler_mao_de_obra(linhas(arqs['mdo']))
    comps = ler_composicoes(linhas(arqs['comp']), mat, mdo)
    insumos = [*mat.values(), *mdo.values()]
    dados = {'ref': ref, 'uf': uf, 'regime': regime, 'desonerado': regime == 'CD', 'fonte': pasta, 'gerado': datetime.date.today().isoformat(),
             'cab_comp': ['cod', 'descricao', 'un', 'grupo', 'total', '[material,mao_de_obra,equipamento,outros]', 'pct_as', 'incompleta', 'itens[tipo I|C, cod, coef]'],
             'cab_ins': ['cod', 'descricao', 'un', 'balde', 'classificacao', 'preco', 'origem', '[salario_sem_encargos, encargos_pct]'],
             'insumos': insumos, 'comps': comps}
    os.makedirs(os.path.join(saida_dir, pasta, uf), exist_ok=True)
    arq = os.path.join(saida_dir, pasta, uf, f'{ref}-{regime}.json')
    with open(arq, 'w', encoding='utf-8') as f: json.dump(dados, f, ensure_ascii=False, separators=(',', ':'))
    sem = sum(1 for c in comps if c[4] is None)
    print(regime, arq, os.path.getsize(arq), 'bytes;', len(comps), 'composicoes;', sem, 'sem custo;', len(insumos), 'insumos (GOINFRA)')
    return ref, regime

if __name__ == '__main__':
    import sys
    c, m, d = sys.argv[1:4]
    gerar_goinfra({'comp': c, 'mat': m, 'mdo': d}, sys.argv[4] if len(sys.argv) > 4 else 'dados')
