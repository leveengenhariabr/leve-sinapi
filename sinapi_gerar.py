#!/usr/bin/env python3
"""Lê a planilha SINAPI_Referência_AAAA_MM.xlsx da Caixa e gera, para uma UF, um JSON por regime
(sem desoneração = SD, com desoneração = CD) com insumos, composições e a decomposição
em material / mão de obra / equipamento / outros (explodindo as composições auxiliares)."""
import sys, re, json, os, datetime
import openpyxl

UF_PADRAO = 'MT'
CLASSE = {  # classificação do insumo na planilha -> balde
    'MATERIAL': 'mat', 'MAO DE OBRA': 'mo', 'ENCARGOS COMPLEMENTARES': 'mo',
    'EQUIPAMENTO (AQUISIÇÃO)': 'eq', 'EQUIPAMENTO (SEM CRÉDITO IVA)': 'eq', 'EQUIPAMENTO (LOCAÇÃO)': 'eq',
    'SERVIÇOS': 'out', 'ESPECIAIS': 'out',
}
RE_COD = re.compile(r'\)\s*,\s*(\d+)\s*\)\s*$')

def cod(v):
    if v is None or v == '': return None
    if isinstance(v, (int, float)): return str(int(v))
    s = str(v).strip()
    m = RE_COD.search(s)
    if s.startswith('=') and m: return m.group(1)
    try: return str(int(float(s)))
    except ValueError: return s

def num(v):
    if v is None or v == '': return None
    try: return float(v)
    except (TypeError, ValueError): return None

def col_uf(linha, uf):
    for i, x in enumerate(linha):
        if x == uf: return i
    raise SystemExit(f'UF {uf} nao encontrada no cabecalho')

def ler_aba(wb, nome, uf, formulas):
    ws = wb[nome]
    rows = list(ws.iter_rows(values_only=True))
    return rows

def gerar(caminho, uf, saida_dir, mes):
    wb_v = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    wb_f = openpyxl.load_workbook(caminho, read_only=True, data_only=False)
    # estrutura (igual nos dois regimes): analítico
    an = list(wb_v['Analítico'].iter_rows(min_row=11, values_only=True))
    comps_cab, itens = {}, {}
    ordem = []
    for r in an:
        g, c, tipo, ci, desc, un, coef, sit = (list(r) + [None] * 8)[:8]
        c = cod(c)
        if not c: continue
        if not tipo:
            comps_cab[c] = {'g': g, 'd': desc, 'u': un, 'sit': sit}
            ordem.append(c); itens[c] = []
        else:
            itens[c].append(('I' if tipo == 'INSUMO' else 'C', cod(ci), num(coef)))
    # códigos das composições sintéticas (vêm em fórmula)
    resultado = {}
    for regime, ins_aba, comp_aba in (('SD', 'ISD', 'CSD'), ('CD', 'ICD', 'CCD')):
        rows = list(wb_v[ins_aba].iter_rows(values_only=True))
        cab_ins = rows[3] if rows[3][4] and 'DESONERA' in str(rows[3][4]) else rows[3]
        ci = col_uf(rows[3], uf)
        insumos = {}
        for r in rows[10:]:
            if r[1] is None: continue
            c = cod(r[1]); cl = CLASSE.get(r[0], 'out')
            insumos[c] = {'d': r[2], 'u': r[3], 'cl': cl, 'cls': r[0], 'o': r[4], 'p': num(r[ci])}
        rv = list(wb_v[comp_aba].iter_rows(values_only=True))
        rf = list(wb_f[comp_aba].iter_rows(min_row=11, values_only=True))
        cc = col_uf(rv[8], uf)  # linha com siglas repetidas; coluna do custo
        custo, asp, grupo_sint = {}, {}, {}
        for vv, ff in zip(rv[10:], rf):
            c = cod(ff[1])
            if not c: continue
            custo[c] = num(vv[cc]); asp[c] = num(vv[cc + 1]); grupo_sint[c] = vv[0]
        # decomposição: usa o custo OFICIAL de cada composição auxiliar (já arredondado pela Caixa) e reparte
        # em material / mão de obra / equipamento / outros conforme os insumos de cada nível.
        # Insumos sem preço na UF (ex.: valor de aquisição de equipamento) entram como "resto" do nível:
        # em composições de depreciação/juros/manutenção o resto é equipamento; nos demais fica em "outros".
        memo = {}
        def eh_equip(c):
            return str(comps_cab[c]['g'] or '').lower().startswith(('depreciação', 'depreciacao'))
        def explode(c, pilha=()):
            """devolve (baldes em R$ por unidade da composição, houve_item_sem_preco)"""
            if c in memo: return memo[c]
            if c in pilha: return ({'mat': 0.0, 'mo': 0.0, 'eq': 0.0, 'out': 0.0}, True)
            b = {'mat': 0.0, 'mo': 0.0, 'eq': 0.0, 'out': 0.0}; faltou = False; sem_preco = {}
            for t, ic, coef in itens.get(c, []):
                if coef is None: continue
                if t == 'I':
                    i = insumos.get(ic)
                    if not i or i['p'] is None:
                        faltou = True
                        if i: sem_preco[i['cl']] = sem_preco.get(i['cl'], 0) + 1
                        continue
                    b[i['cl']] += coef * i['p']
                else:
                    t_aux = custo.get(ic)
                    if ic not in comps_cab or not t_aux: faltou = True; continue
                    sb, sf = explode(ic, pilha + (c,))
                    soma = sum(sb.values())
                    if soma > 0:
                        for k in b: b[k] += coef * t_aux * sb[k] / soma
                    else:
                        b['eq' if eh_equip(ic) else 'out'] += coef * t_aux
            tot = custo.get(c)
            if tot:
                soma = sum(b.values()); resto = tot - soma
                if resto > max(0.02, 0.01 * tot):          # parte relevante sem preço de insumo
                    if sem_preco:                           # reparte o resto entre as classes dos insumos sem preço
                        n = sum(sem_preco.values())
                        for k, q in sem_preco.items(): b[k] += resto * q / n
                    else:
                        b['eq' if eh_equip(c) else 'out'] += resto
                elif soma > 0:                              # só arredondamento: ajusta ao total oficial
                    f = tot / soma
                    for k in b: b[k] *= f
            memo[c] = (b, faltou)
            return memo[c]
        comps = []
        dif_max = 0; sem_custo = 0
        for c in ordem:
            cab = comps_cab[c]
            tot = custo.get(c)
            if not tot:                       # sem custo nesta UF/mês (a planilha traz 0 ou vazio)
                comps.append([c, cab['d'], cab['u'], cab['g'], None, None, asp.get(c), 0, [[t, ic, coef] for t, ic, coef in itens[c]]])
                sem_custo += 1; continue
            b, faltou = explode(c)
            soma = sum(b.values()); dif_max = max(dif_max, abs(soma - tot))
            comps.append([c, cab['d'], cab['u'], cab['g'], tot,
                          [round(b['mat'], 4), round(b['mo'], 4), round(b['eq'], 4), round(b['out'], 4)],
                          asp.get(c), 1 if faltou else 0,
                          [[t, ic, coef] for t, ic, coef in itens[c]]])
        ins_lista = [[c, i['d'], i['u'], i['cl'], i['cls'], i['p'], i['o']] for c, i in insumos.items()]
        dados = {'ref': mes, 'uf': uf, 'regime': regime, 'desonerado': regime == 'CD',
                 'gerado': datetime.date.today().isoformat(),
                 'cab_comp': ['cod', 'descricao', 'un', 'grupo', 'total', '[material,mao_de_obra,equipamento,outros]', 'pct_as', 'incompleta', 'itens[tipo I|C, cod, coef]'],
                 'cab_ins': ['cod', 'descricao', 'un', 'balde', 'classificacao', 'preco', 'origem'],
                 'insumos': ins_lista, 'comps': comps}
        os.makedirs(os.path.join(saida_dir, uf), exist_ok=True)
        arq = os.path.join(saida_dir, uf, f'{mes}-{regime}.json')
        with open(arq, 'w', encoding='utf-8') as f: json.dump(dados, f, ensure_ascii=False, separators=(',', ':'))
        print(regime, arq, os.path.getsize(arq), 'bytes;', len(comps), 'composicoes;', sem_custo, 'sem custo;', len(ins_lista), 'insumos; maior diferenca soma x oficial: R$', round(dif_max, 4))

if __name__ == '__main__':
    arq, mes = sys.argv[1], sys.argv[2]            # mes = AAAA-MM
    uf = sys.argv[3] if len(sys.argv) > 3 else UF_PADRAO
    saida = sys.argv[4] if len(sys.argv) > 4 else 'dados'
    gerar(arq, uf, saida, mes)
