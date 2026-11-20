#!/usr/bin/env python3
"""Baixa um zip antigo do SINAPI (antes de 2025) e imprime a estrutura das planilhas (abas, colunas, primeiras linhas). Só para diagnóstico."""
import io, sys, zipfile, http.cookiejar, urllib.request, urllib.error, time
import openpyxl
UF = sys.argv[1] if len(sys.argv) > 1 else 'MT'
REF = sys.argv[2] if len(sys.argv) > 2 else '202412'
REG = sys.argv[3] if len(sys.argv) > 3 else 'NaoDesonerado'
SUF = sys.argv[4] if len(sys.argv) > 4 else ''
UA = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'}
ab = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
class _Sem(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k): return None
dados = None
import http.client, urllib.parse
def passo_a_passo(u):
    jar = {}
    for n in range(8):
        pu = urllib.parse.urlparse(u)
        c = http.client.HTTPSConnection(pu.netloc, timeout=60)
        h = dict(UA); h['Accept'] = 'text/html,application/zip,*/*'; h['Accept-Language'] = 'pt-BR,pt;q=0.9'
        if jar: h['Cookie'] = '; '.join(f'{k}={v}' for k, v in jar.items())
        c.request('GET', pu.path + (('?' + pu.query) if pu.query else ''), headers=h)
        r = c.getresponse(); corpo = r.read(300)
        print(' passo', n, r.status, r.getheader('Content-Type'), 'Location=', r.getheader('Location'), 'Set-Cookie=', [v[:60] for k, v in r.getheaders() if k.lower() == 'set-cookie'], 'corpo=', corpo[:80])
        for k, v in r.getheaders():
            if k.lower() == 'set-cookie':
                kv = v.split(';')[0]; a, _, b = kv.partition('='); jar[a] = b
        if r.status in (301, 302, 303, 307, 308) and r.getheader('Location'):
            u = urllib.parse.urljoin(u, r.getheader('Location')); continue
        return r.status
import subprocess
def com_curl(u):
    for extra in ([], ['--http1.1'], ['--http2']):
        cmd = ['curl', '-sS', '-L', '--max-redirs', '6', '--compressed', '-o', '/tmp/t.zip', '-w', '%{http_code} %{size_download} %{num_redirects} %{url_effective}', '-c', '/tmp/jar.txt', '-b', '/tmp/jar.txt',
               '-A', UA['User-Agent'], '-e', 'https://www.caixa.gov.br/site/paginas/downloads.aspx',
               '-H', 'Accept: text/html,application/xhtml+xml,application/zip,*/*;q=0.8', '-H', 'Accept-Language: pt-BR,pt;q=0.9,en;q=0.8',
               '-H', 'Upgrade-Insecure-Requests: 1', '-H', 'Sec-Fetch-Dest: document', '-H', 'Sec-Fetch-Mode: navigate', '-H', 'Sec-Fetch-Site: same-origin'] + extra + [u]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=200)
        print('CURL', extra, 'saida=', r.stdout[-200:], 'erro=', r.stderr[-150:])
        if r.stdout.startswith('200') and zipfile.is_zipfile('/tmp/t.zip'): return open('/tmp/t.zip', 'rb').read()
    return None
_z = com_curl(f'https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{UF.lower()}/SINAPI_ref_Insumos_Composicoes_{UF}_{REF}_{REG}_Retificacao01.zip')
print('CURL RESULTADO:', 'ZIP OK' if _z else 'sem zip')
passo_a_passo(f'https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{UF.lower()}/SINAPI_ref_Insumos_Composicoes_{UF}_{REF}_{REG}_Retificacao01.zip')
for suf in ('', '_Retificacao01', '_Retificacao02'):
    u = f'https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{UF.lower()}/SINAPI_ref_Insumos_Composicoes_{UF}_{REF}_{REG}{suf}.zip'
    try:
        urllib.request.build_opener(_Sem).open(urllib.request.Request(u, headers=UA), timeout=60)
        print('sem redirect, ok:', u)
    except urllib.error.HTTPError as e:
        print('TESTE', e.code, 'Location=', e.headers.get('Location'), '|', u); print('CABECALHOS:', dict(e.headers))
    except Exception as e:
        print('TESTE', type(e).__name__, u)
    for t in range(2):
        try:
            with ab.open(urllib.request.Request(u, headers=UA), timeout=180) as r: dados = r.read()
            break
        except Exception as e:
            print('falhou', type(e).__name__, getattr(e, 'code', ''), u); time.sleep(4)
    if dados and zipfile.is_zipfile(io.BytesIO(dados)): break
    dados = None
if not dados or not zipfile.is_zipfile(io.BytesIO(dados)): sys.exit('sem zip: ' + u)
print('baixado', u, len(dados), 'bytes')
z = zipfile.ZipFile(io.BytesIO(dados))
for i in z.infolist():
    print('ARQUIVO', i.filename, i.file_size)
    if not i.filename.lower().endswith('.xlsx'): continue
    wb = openpyxl.load_workbook(io.BytesIO(z.read(i)), read_only=True, data_only=False)
    for ws in wb.worksheets:
        print('  ABA', repr(ws.title), 'linhas~', ws.max_row, 'colunas~', ws.max_column)
        for n, row in enumerate(ws.iter_rows(values_only=True)):
            if n >= 11: break
            print('   ', n, [str(c)[:38] if c is not None else None for c in row[:32]])
