"""ORSE (Orçamentação de Obras de Sergipe) -> arquivos do LeveApp.
O ORSE é um sistema: a composição só sai em PDF ("Relatório de composições"), e os preços dos insumos e dos serviços em Excel.
Entradas: pdf das composições + .xls/.xlsx de insumos + .xls/.xlsx de serviços (colunas: código, descrição, unidade, preço).
Saída: dados/orse/SE/AAAA-MM-SD.json (o ORSE só traz um regime de encargos; o que o relatório informa é o % de encargos sociais)."""
import os, re, json, datetime, subprocess, shutil, unicodedata
NUM = re.compile(r'^-?[\d\.]*\d,\d+$')
CAB = re.compile(r'^\s*(\d+)\s*/\s*(\w+)\s+(.*)$')
MESES = {'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12}
SECOES = [('MÃO-DE-OBRA', 'mo'), ('MAO-DE-OBRA', 'mo'), ('MATERIAL', 'mat'), ('SERVIÇO AUXILIAR', 'aux'), ('SERVICO AUXILIAR', 'aux'), ('EQUIPAMENTO', 'eq'), ('SERVIÇO DE TERCEIRO', 'out'), ('SERVICO DE TERCEIRO', 'out')]
RUIDO = ('RELATÓRIO DE COMPOSIÇÕES', 'Ref :', 'ORSE - Orçamentação', 'EQUIPAMENTO        MÃO-DE-OBRA', 'RESUMO DA COMPOSIÇÃO', 'COMPOSIÇÃO SINTÉTICA', 'COMPOSIÇÃO PADRÃO')

def n(s):
    if s is None or s == '': return None
    if isinstance(s, (int, float)): return float(s)
    s = str(s).strip().replace('R$', '').strip()
    if not s: return None
    if ',' in s: s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except ValueError: return None

def texto_pdf(pdf):
    if not shutil.which('pdftotext'):
        subprocess.run('sudo apt-get install -y poppler-utils >/dev/null 2>&1', shell=True)
    return subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout

def referencia(texto):
    m = re.search(r'Ref\s*:\s*([a-zçã]+)\s*/\s*(\d{4})', texto[:5000], re.I)
    if not m or m.group(1).lower() not in MESES: return None
    return f'{m.group(2)}-{MESES[m.group(1).lower()]:02d}'

def eh_pdf_orse(caminho):
    try: return 'RELATÓRIO DE COMPOSIÇÕES ( ORSE )' in subprocess.run(['pdftotext', '-l', '2', '-layout', caminho, '-'], capture_output=True, text=True).stdout
    except Exception: return False

def linhas_excel(caminho):
    if caminho.lower().endswith('.xls'):
        try:
            from python_calamine import CalamineWorkbook
            wb = CalamineWorkbook.from_path(caminho); return wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        except ImportError:
            import xlrd
            s = xlrd.open_workbook(caminho).sheet_by_index(0); return [s.row_values(i) for i in range(s.nrows)]
    import openpyxl
    return [list(r) for r in openpyxl.load_workbook(caminho, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)]

def ler_precos(caminho):
    """planilha do ORSE sem cabeçalho: código, descrição, unidade, preço (e às vezes um '#')"""
    out = {}
    for r in linhas_excel(caminho):
        v = [x for x in r if x not in ('', None)]
        if len(v) < 4: continue
        try: cod = str(int(float(v[0]))).zfill(5)
        except (ValueError, TypeError): continue
        p = n(v[3])
        if p is None: continue
        out[cod] = (str(v[1]).strip(), str(v[2]).strip(), p)
    return out

ROT_RESUMO = ('EQUIPAMENTO', 'MÃO-DE-OBRA', 'ENCARCO SOCIAL', 'MATERIAL', 'SERV. TERCEIRO', 'CUSTO TOTAL')

def so_codigo(l):
    """linha "NNNNN / ORSE ... [prod] UN" sem descrição = a linha do código de uma composição"""
    m = CAB.match(l)
    if not m or m.group(2) != 'ORSE': return None
    t = m.group(3).split()
    if not t or NUM.match(t[-1]) or len(t) > 2 or (len(t) == 2 and not NUM.match(t[0])): return None
    return m.group(1).zfill(5), t[-1], (n(t[0]) if len(t) == 2 else None)

def _centros(l):
    return [(m.start() + m.end()) / 2 for m in re.finditer(r'\S+', l)]

def ler_composicoes(texto):
    """devolve lista de blocos: {cod, desc, un, prod, tipo, enc, itens:[(secao, fonte, cod, desc, un, nums)], resumo=[eq, mo, enc, mat, terc, total]}"""
    L = texto.split('\n'); blocos = []; b = None; secao = None; ult = None; esp_resumo = False; rot_pos = None
    for i, l in enumerate(L):
        s = l.strip()
        if not s: continue
        c = so_codigo(l)
        if c:
            k = i - 1; desc = ''
            while k >= max(0, i - 8):
                if L[k].strip().startswith('CÓDIGO'): desc = re.sub(r'\s+(PROD\.?)?\s*(EQUIPE)?\s*(UNID)?\s*$', '', L[k].strip()[6:].strip()).strip(); break
                k -= 1
            desc = re.sub(r'\s+PROD\.?$', '', desc)
            b = {'cod': c[0], 'desc': desc, 'un': c[1], 'prod': c[2], 'tipo': '', 'enc': None, 'itens': [], 'resumo': None}
            blocos.append(b); secao = None; ult = None; esp_resumo = False; continue
        if s.startswith('CÓDIGO'):
            resto = s[6:].strip()
            sc = next((c2 for k2, c2 in SECOES if resto.startswith(k2)), None)
            if sc and b is not None:
                secao = sc; ult = None
                me = re.search(r'Encargo Social:\s*([\d\.,]+)\s*%', resto)
                if me: b['enc'] = n(me.group(1))
            continue
        if b is None: continue
        if s.startswith('COMPOSIÇÃO'): b['tipo'] = s; continue
        if s.startswith('EQUIPAMENTO') and 'ENCARCO' in s: rot_pos = [l.find(r) + len(r) for r in ROT_RESUMO] if all(r in l for r in ROT_RESUMO) else None; continue
        if s.startswith('RESUMO DA COMPOSIÇÃO'): esp_resumo = True; continue
        if esp_resumo:
            tk = s.split()
            if all(NUM.match(x) for x in tk):
                fins = [m.end() for m in re.finditer(r'\S+', l)]; res = [0.0] * 6
                for x, fx in zip(tk, fins):
                    if rot_pos and len(rot_pos) == 6: res[min(range(6), key=lambda q: abs(fx - rot_pos[q]))] = n(x)
                    elif len(tk) == 6: res = [n(v) for v in tk]
                b['resumo'] = res if (rot_pos and len(rot_pos) == 6) or len(tk) == 6 else [0.0] * 5 + [n(tk[-1])]; esp_resumo = False
            continue
        sec = next((c2 for k2, c2 in SECOES if s.startswith(k2)), None)
        if sec: secao = sec; ult = None; continue
        m = CAB.match(l)
        if m:
            toks = m.group(3).split(); nums = []
            while toks and NUM.match(toks[-1]): nums.append(toks.pop())
            nums.reverse()
            if len(nums) in (3, 4, 6) and len(toks) >= 1:
                un = toks.pop(); ult = [secao, m.group(2), m.group(1).zfill(5), ' '.join(toks), un, [n(x) for x in nums]]
                b['itens'].append(ult); continue
        if ult is not None and not any(x in s for x in RUIDO) and not NUM.match(s.split()[0]) and not s.startswith(('CÓDIGO', 'ORSE -')) and not re.match(r'^\d{2}/\d{2}/\d{4}', s):
            ult[3] += ' ' + s     # descrição que continua na linha de baixo
    return blocos

def gerar_orse(pdf, arq_ins, arq_serv, saida_dir, uf='SE', pasta='orse'):
    texto = texto_pdf(pdf); ref = referencia(texto)
    if not ref: raise SystemExit('Não achei o mês (Ref : agosto/2026) no PDF do ORSE.')
    blocos = ler_composicoes(texto)
    precos = {}
    for a in (arq_ins, arq_serv):
        if a: precos.update({k: v for k, v in ler_precos(a).items() if k not in precos})
    tot_comp = {b['cod']: (b['resumo'][5] if b['resumo'] else None) for b in blocos}
    insumos = {}; comps = []; sem_preco = 0
    def put(chave, d, u, balde, pr, fonte, extra=None, rot=None):
        """o mesmo código pode aparecer com preços diferentes (ex.: equipamento em composição sintética x padrão DNIT): vira outra linha"""
        k = chave
        while k in insumos and abs((insumos[k][5] or 0) - (pr or 0)) > 0.015: k = chave + '/' + ('%.2f' % pr).replace('.', ',')
        if k not in insumos: insumos[k] = [k, d, u, balde, rot, pr, fonte] + ([extra] if extra else [])
        return k
    for b in blocos:
        itens = []; enc = b['enc'] or 0.0; sint = b['tipo'].endswith('SINTÉTICA'); p = b['prod'] if (b['prod'] and not sint) else None
        div = (1.0 / p) if p else 1.0     # padrão DNIT: equipamento e mão de obra são divididos pela produção da equipe
        for sec, fonte, cod, desc, un, nums in b['itens']:
            qtd = nums[0]
            if sec == 'eq' and len(nums) == 4:
                if not nums[3]: continue     # equipamento sem uso na composição
                nums = [nums[0], 1.0, 0.0, nums[1], nums[2], nums[3]]
            if sec == 'eq' and len(nums) == 6:
                _, prd, imp, cp, ci, ct = nums; q = qtd * div
                d0 = precos.get(cod, (desc,))[0] if False else desc
                if prd: k1 = put('E' + cod, d0, 'h', 'eq', cp, fonte); itens.append(['I', k1, round(q * prd, 7)])
                if imp: k2 = put('E' + cod + '-I', d0 + ' (improdutivo)', 'h', 'eq', ci, fonte); itens.append(['I', k2, round(q * imp, 7)])
                continue
            if len(nums) == 4: nums = nums[:1] + nums[-2:]
            unit = nums[-2]
            if sec in ('aux', 'out') and fonte == 'ORSE' and tot_comp.get(cod) is not None and cod != b['cod'] and abs(tot_comp[cod] - unit) <= 0.02:
                itens.append(['C', cod, qtd]); continue
            if sec == 'mo':
                # sintético: valor unitário SEM encargos; guardamos COM encargos (como GOINFRA/SICRO) e o % em extra
                tot_i = nums[-1] or 0.0; tol = max(0.02, 0.03 * tot_i); pc = unit * (1 + enc / 100) if enc else unit
                if not sint: pr = unit                                        # padrão DNIT: o valor já traz encargos
                elif abs(tot_i - qtd * pc) <= tol: pr = round(pc, 2)          # sintético: unitário sem encargos; o total já os soma
                elif abs(tot_i - qtd * unit) <= tol: pr = unit                # unidade de serviço (sem encargos)
                else: pr = round(tot_i / qtd, 2) if qtd else unit
                com_enc = bool(enc) and (not sint or abs(pr - unit) > 0.005)     # padrão DNIT: o valor já traz encargos; sintético: só se o total difere do unitário
                chave = ('S' if fonte != 'ORSE' else 'MO') + cod
                chave = put(chave, desc, un, 'mo', pr, fonte, [round(pr / (1 + enc / 100), 2) if com_enc else pr, enc if com_enc else None], 'MÃO DE OBRA')
                itens.append(['I', chave, round(qtd * div, 7)]); continue
            balde = {'mat': 'mat', 'eq': 'eq', 'out': 'out', 'aux': 'out'}.get(sec, 'out')
            chave = ('S' if fonte != 'ORSE' else {'mat': '', 'eq': 'E', 'out': 'T', 'aux': 'T'}.get(sec, 'T')) + cod
            d, u = desc or cod, un
            x = precos.get(cod) if (fonte == 'ORSE' and sec in ('mat', 'out', 'aux')) else None
            if x and (abs(x[2] - unit) <= 0.02 or not desc): d, u = x[0], x[1]
            elif fonte == 'ORSE' and sec in ('mat', 'out', 'aux'): sem_preco += 1
            chave = put(chave, d, u, balde, unit, fonte)
            itens.append(['I', chave, round(qtd * (div if sec == 'eq' else 1), 7)])
        r = b['resumo']
        if r:
            eq, mo, es, mat, terc, total = r
            split = [round(mat, 4), round(mo + es, 4), round(eq, 4), round(terc, 4)]
        else: total = None; split = None
        comps.append([b['cod'], b['desc'], b['un'], 'ORSE', total, split, None, 0 if r else 1, itens])
    dados = {'ref': ref, 'uf': uf, 'regime': 'SD', 'desonerado': False, 'fonte': pasta, 'gerado': datetime.date.today().isoformat(),
             'cab_comp': ['cod', 'descricao', 'un', 'grupo', 'total', '[material,mao_de_obra,equipamento,outros]', 'pct_as', 'incompleta', 'itens[tipo I|C, cod, coef]'],
             'cab_ins': ['cod', 'descricao', 'un', 'balde', 'classificacao', 'preco', 'origem', '[salario_sem_encargos, encargos_pct]'],
             'insumos': list(insumos.values()), 'comps': comps}
    os.makedirs(os.path.join(saida_dir, pasta, uf), exist_ok=True)
    arq = os.path.join(saida_dir, pasta, uf, f'{ref}-SD.json')
    with open(arq, 'w', encoding='utf-8') as f: json.dump(dados, f, ensure_ascii=False, separators=(',', ':'))
    print('ORSE', ref, arq, os.path.getsize(arq), 'bytes;', len(comps), 'composicoes;', len(insumos), 'insumos;', sem_preco, 'sem preço na planilha')
    return ref

if __name__ == '__main__':
    import sys
    gerar_orse(*sys.argv[1:5])
