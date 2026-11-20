"""SICRO (DNIT) -> JSON no mesmo formato do SINAPI. Entrada: relatórios .xlsx do SICRO (por estado/mês).
Lê o Analítico de Composições (vem só num regime) e RECALCULA o outro regime com os preços de mão de obra e equipamentos dele.
Saída: dados/sicro/<UF>/AAAA-MM-SD.json e -CD.json"""
import json, os, re
from datetime import datetime

MESES = {'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12}

def _num(v):
    if v in (None, ''): return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).strip().replace('.', '').replace(',', '.') if ',' in str(v) else str(v).strip())
    except ValueError: return None

def _linhas(caminho):
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    return [[('' if x is None else (x.strip() if isinstance(x, str) else x)) for x in r] for r in wb.worksheets[0].iter_rows(values_only=True)]

def identificar(rows):
    """-> {tipo: comp|mat|mo|eq|enc, desonerado: True|False|None, uf, ref}"""
    cab = ' '.join(str(c) for r in rows[:6] for c in r if c != '').lower()
    h = [str(c).lower() for c in rows[0]]
    t = None
    if rows[0] and str(rows[0][0]).startswith('SISTEMA DE CUSTOS'): t = 'comp'
    elif h[:2] == ['código', 'descrição'] and any('aquisi' in x for x in h): t = 'eq'
    elif h[:3] == ['código', 'descrição', 'unidade'] and any('preço unit' in x for x in h): t = 'mat'
    elif h[0] == 'código' and any(x.startswith('salário') for x in h): t = 'mo_an'
    elif h[:3] == ['código', 'descrição', 'unidade'] and any(x.startswith('custo') for x in h) and len(rows) > 1 and str(rows[1][0]).startswith('P'): t = 'mo'
    return t

def ler_equip(rows):
    h = [str(c).lower() for c in rows[0]]
    ip = next(i for i, x in enumerate(h) if 'custo produtivo' in x); ii = next(i for i, x in enumerate(h) if 'custo improdutivo' in x)
    return {str(r[0]): (r[1], _num(r[ip]), _num(r[ii])) for r in rows[1:] if r[0]}

def ler_mat(rows):
    return {str(r[0]): (r[1], r[2], _num(r[3])) for r in rows[1:] if r[0]}

def ler_mo(rows):  # sintético: custo/h já com encargos
    return {str(r[0]): (r[1], r[2], _num(r[3])) for r in rows[1:] if r[0] and _num(r[3]) is not None}

def ler_salarios(rows):  # analítico: código, descrição, unidade, salário
    out = {}; ic = next(i for i, x in enumerate(rows[0]) if str(x).lower().startswith('salário'))
    for r in rows[1:]:
        if re.match(r'^P\d+', str(r[0])): out[str(r[0])] = _num(r[ic])
    return out

def ler_composicoes(rows):
    comps = {}; fic_dem = {}
    st = [i for i, r in enumerate(rows) if str(r[0]).startswith('SISTEMA DE CUSTOS')]
    mes = None
    for a, b in zip(st, st[1:] + [len(rows)]):
        blk = rows[a:b]
        fic = _num(blk[0][7]) or 0.0
        m = re.search(r'([A-Za-zçÇ]+)/(\d{4})', str(blk[1][3]))
        if m and not mes and m.group(1).lower() in MESES: mes = f"{m.group(2)}-{MESES[m.group(1).lower()]:02d}"
        prod = _num(blk[1][7]) or 1.0; un = str(blk[1][8])
        cod, desc = str(blk[2][0]), str(blk[2][1])
        sec = None; eq = []; mo = []; mat = []; aux = []; tf = 0.0; direto = None; desp = 0.0
        for r in blk[3:]:
            t = str(r[0])
            m2 = re.match(r'^([A-F]) - ', t)
            if m2: sec = m2.group(1); continue
            if str(r[4]).startswith('Custo unitário direto total'): direto = _num(r[8]); continue
            if sec == 'B' and str(r[6]).startswith('Despesas diversas'): desp += _num(r[8]) or 0.0; continue
            if t.startswith('Obs') or t == '' : continue
            if sec == 'A' and _num(r[2]) is not None and re.match(r'^E\d+', t): eq.append((t, _num(r[2]), _num(r[3]) or 0.0, _num(r[4]) or 0.0))
            elif sec == 'B' and re.match(r'^P\d+', t): mo.append((t, _num(r[2])))
            elif sec == 'C' and _num(r[2]) is not None: mat.append((t, _num(r[2])))
            elif sec == 'D' and _num(r[2]) is not None: aux.append((t, str(r[1]), _num(r[2]), str(r[3])))
            elif sec == 'E' and _num(r[8]) is not None and not str(r[2]).startswith('Custo'): tf += _num(r[8])
        comps[cod] = dict(cod=cod, desc=desc, un=un, prod=prod, fic=fic, eq=eq, mo=mo, mat=mat, aux=aux, tf=tf, desp=desp, direto=direto)
    return comps, mes

def calcular(comps, eq_p, mo_p, mat_p):
    """-> {cod: (total, [mat,mo,eq,out])} recursivo (auxiliares)"""
    memo = {}
    def tot(cod, pilha=()):
        if cod in memo: return memo[cod]
        c = comps.get(cod)
        if c is None or cod in pilha: return None
        e = sum(q * (op * eq_p.get(k, (0, 0, 0))[1] + im * eq_p.get(k, (0, 0, 0))[2]) for k, q, op, im in c['eq']) / c['prod']
        m = sum(q * mo_p.get(k, (0, 0, 0))[2] for k, q in c['mo']) / c['prod']
        dd = c['desp'] / c['prod']
        exe = e + m + dd
        # FIC incide sobre a mão de obra + o custo improdutivo (cheio) dos equipamentos
        fic = c['fic'] * (sum(q * eq_p.get(k, (0, 0, 0))[2] for k, q, op, im in c['eq']) + sum(q * mo_p.get(k, (0, 0, 0))[2] for k, q in c['mo'])) / c['prod']
        ma = sum(q * (mat_p.get(k, (0, 0, 0))[2] or 0) for k, q in c['mat'])
        ax = 0.0
        for k, _, q, _u in c['aux']:
            sx = tot(k, pilha + (cod,))
            ax += q * (round(sx[0], 2) if sx else 0)
        total = exe + fic + ma + ax + c['tf']
        # parte de material/mo/eq dentro das auxiliares entra proporcionalmente
        sp = [ma, m, e, fic + c['tf'] + dd]
        for k, _, q, _u in c['aux']:
            s = tot(k, pilha + (cod,))
            if s:
                for i in range(4): sp[i] += q * s[1][i]
        memo[cod] = (total, sp)
        return memo[cod]
    for k in comps: tot(k)
    return memo

def gerar_sicro(arq, saida_dir, uf='MT', pasta='sicro'):
    """arq: dict com 'comp','mat','eq_sd','eq_cd','mo_sd','mo_cd' (linhas) e 'sal' (analítico MO) — qualquer um dos regimes presentes"""
    comps, ref = ler_composicoes(arq['comp'])
    mat = ler_mat(arq['mat'])
    sal = ler_salarios(arq['mo_an']) if arq.get('mo_an') else {}
    # a composição analítica vem num regime; descobrimos comparando o custo da mão de obra no arquivo
    saidas = {}
    for reg in ('SD', 'CD'):
        eqr = ler_equip(arq['eq_' + reg.lower()]); mor = ler_mo(arq['mo_' + reg.lower()])
        r = calcular(comps, eqr, mor, mat)
        insumos = []
        for k, (d, un, p) in sorted(mat.items()): insumos.append([k, d, un, 'mat', 'Material', p, 'C'])
        for k, (d, un, p) in sorted(mor.items()):
            s = sal.get(k); enc = round((p / s - 1) * 100, 2) if s else None
            insumos.append([k, d, un, 'mo', 'Mão de obra', p, 'C', [s, enc]])
        for k, (d, pp, pi) in sorted(eqr.items()):
            insumos.append([k, d + ' (produtivo)', 'h', 'eq', 'Equipamento', pp, 'C'])
            insumos.append([k + '-I', d + ' (improdutivo)', 'h', 'eq', 'Equipamento', pi, 'C'])
        lista = []
        for cod, c in comps.items():
            t, sp = r[cod]
            itens = []
            for k, q, op, im in c['eq']:
                if op: itens.append(['I', k, round(q * op / c['prod'], 8)])
                if im: itens.append(['I', k + '-I', round(q * im / c['prod'], 8)])
            for k, q in c['mo']: itens.append(['I', k, round(q / c['prod'], 8)])
            for k, q in c['mat']: itens.append(['I', k, q])
            for k, d, q, u in c['aux']: itens.append(['C', k, q])
            base = (sum(q * eqr.get(k, (0, 0, 0))[2] for k, q, op, im in c['eq']) + sum(q * mor.get(k, (0, 0, 0))[2] for k, q in c['mo'])) / c['prod']
            if c['fic'] and base:
                ic = 'FIC.' + cod; insumos.append([ic, 'FIC — fator de influência de chuvas (' + cod + ')', 'un', 'out', 'FIC', round(base * c['fic'], 6), 'C']); itens.append(['I', ic, 1])
            if c['desp']:
                idd = 'DD.' + cod; insumos.append([idd, 'Despesas diversas e manutenção (' + cod + ')', 'un', 'out', 'Despesas', round(c['desp'] / c['prod'], 6), 'C']); itens.append(['I', idd, 1])
            if c['tf']:
                it = 'TF.' + cod; insumos.append([it, 'Tempo fixo de transporte de material (' + cod + ')', 'un', 'out', 'Transporte', round(c['tf'], 6), 'C']); itens.append(['I', it, 1])
            inc = False; tt = round(t, 2)
            if reg == 'SD' and c['direto'] is not None and abs(tt - c['direto']) > 0.02:   # raras instalações: vale o total publicado
                sp[3] += c['direto'] - t; tt = c['direto']; inc = True
            lista.append([cod, c['desc'], c['un'], 'SICRO ' + cod[:2], tt, [round(x, 4) for x in sp], None, inc, itens])
        saidas[reg] = {'ref': ref, 'uf': uf, 'regime': reg, 'desonerado': reg == 'CD', 'fonte': 'sicro', 'gerado': datetime.now().strftime('%Y-%m-%d %H:%M'),
                       'cab_comp': ['Código', 'Descrição', 'Unid.', 'Grupo', 'Total'], 'cab_ins': ['Código', 'Descrição', 'Unid.', 'Tipo', 'Classificação', 'Preço', 'Origem'],
                       'insumos': insumos, 'comps': lista}
    pasta_uf = os.path.join(saida_dir, pasta, uf); os.makedirs(pasta_uf, exist_ok=True)
    for reg, d in saidas.items():
        with open(os.path.join(pasta_uf, f'{ref}-{reg}.json'), 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
    return ref, saidas

def verificar(comps, saidas):
    """compara o recalculado do regime em que o arquivo analítico foi emitido com o 'Custo unitário direto total' do arquivo"""
    res = {}
    for reg, d in saidas.items():
        dif = [abs(c[4] - comps[c[0]]['direto']) for c in d['comps'] if comps[c[0]]['direto'] is not None]
        res[reg] = (len(dif), max(dif) if dif else None, sum(1 for x in dif if x > 0.02))
    return res
