// Teste no SITE DE VERDADE, com o Firebase de verdade (não é o Firebase falso dos outros testes).
// Só lê: confere login, regras do Firestore, telas do menu (erro de script, permissão negada, largura) e sai.
// Não cria nem altera obra, pedido, entrega etc. O login registra a sessão, como qualquer entrada no sistema.
//
// Uso (no computador ou no GitHub Actions):
//   LEVE_LOGIN=teste.robo LEVE_SENHA='...' node site-real.mjs            # Full HD
//   LEVE_LOGIN=teste.robo LEVE_SENHA='...' node site-real.mjs celular    # celular
// Variáveis:
//   LEVE_URL              endereço do app (padrão https://leveengenhariabr.com/leveapp/)
//   LEVE_LOGIN/LEVE_SENHA conta de TESTE com cargo comum (não use a conta de administrador)
//   LEVE_APPCHECK_DEBUG   token de depuração do App Check (Firebase → App Check → Apps → Gerenciar tokens de depuração);
//                         sem ele, o reCAPTCHA pode barrar o navegador automático.
import { pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

export async function rodar({ url = process.env.LEVE_URL || 'https://leveengenhariabr.com/leveapp/', login = process.env.LEVE_LOGIN, senha = process.env.LEVE_SENHA,
  tokenAppCheck = process.env.LEVE_APPCHECK_DEBUG || '', celular = false, prepararContexto = null, esperarRegras = true } = {}) {
  if (!login || !senha) throw new Error('Defina LEVE_LOGIN e LEVE_SENHA (conta de teste).');
  let pw; try { pw = await import('playwright'); } catch { pw = createRequire(import.meta.url)('playwright'); } // instalado na pasta ou global
  const { chromium } = pw.chromium ? pw : pw.default;
  const base = url.endsWith('/') ? url : url + '/';
  const nav = await chromium.launch();
  const ctx = await nav.newContext({ viewport: celular ? { width: 390, height: 844 } : { width: 1920, height: 1080 }, locale: 'pt-BR', timezoneId: 'America/Sao_Paulo' });
  if (tokenAppCheck) await ctx.addInitScript(t => { self.FIREBASE_APPCHECK_DEBUG_TOKEN = t; }, tokenAppCheck);
  if (prepararContexto) await prepararContexto(ctx);
  const pg = await ctx.newPage();
  const falhas = [], erros = [], negadas = [];
  const ok = m => console.log('ok - ' + m), falha = m => { falhas.push(m); console.log('FALHA - ' + m); };
  let onde = 'carga';
  pg.on('pageerror', e => erros.push(`${onde}: ${e.message}`));
  pg.on('console', m => { const t = m.text(); if (m.type() !== 'error') return; if (/permission|insufficient permissions/i.test(t)) negadas.push(`${onde}: ${t.slice(0, 160)}`); else if (/app.?check|recaptcha/i.test(t)) erros.push(`${onde}: App Check — ${t.slice(0, 160)}`); else if (!/Failed to load resource|favicon/i.test(t)) erros.push(`${onde}: ${t.slice(0, 200)}`); });
  pg.on('dialog', d => d.dismiss());
  // regras: pedido feito pelo próprio módulo de dados do app (mesmo endereço = mesma conexão já logada)
  const tentar = (op, col, id) => pg.evaluate(async ({ op, col, id }) => {
    const { Dados } = await import(new URL('js/dados.js', location.href).href);
    try { const r = op === 'obter' ? await Dados.obter(col, id) : await Dados.listar(col); return { ok: true, n: Array.isArray(r) ? r.length : (r ? 1 : 0) }; }
    catch (e) { return { ok: false, codigo: e.code || '', msg: String(e.message || e).slice(0, 160) }; }
  }, { op, col, id });
  const negado = r => !r.ok && /permission|insufficient/i.test(r.codigo + ' ' + r.msg);
  try {
    // 0) versão publicada
    const v = await (await ctx.request.get(base + 'versao.json?t=' + Date.now())).json().catch(() => null);
    v?.versao ? ok(`site no ar, versão ${v.versao}`) : falha('não consegui ler versao.json do site');
    await pg.goto(base + 'index.html', { waitUntil: 'domcontentloaded' });
    await pg.waitForSelector('#form-login', { state: 'visible', timeout: 30000 });
    // 1) sem login, o banco não pode responder
    if (esperarRegras) { const r = await tentar('obter', 'config', 'geral'); negado(r) ? ok('regras: sem login, o banco recusa a leitura') : falha(`regras: sem login a leitura de config não foi recusada (${JSON.stringify(r)})`); }
    // 2) login
    onde = 'login';
    await pg.fill('[name=login]', login); await pg.fill('[name=senha]', senha); await pg.click('#form-login button[type=submit]');
    const entrou = await pg.waitForSelector('#sair:visible, #dlg-troca-senha[open]', { timeout: 30000 }).then(() => true).catch(() => false);
    await pg.waitForTimeout(1500);
    if (!entrou) { const aviso = await pg.locator('#aviso-login').innerText().catch(() => ''); throw new Error('não entrou no sistema' + (aviso ? `: "${aviso}"` : '') + '. Confira login/senha da conta de teste e, se o aviso falar de verificação/robô, o token do App Check.'); }
    if (/Crie uma senha forte/.test(await pg.locator('#dlg-troca-senha[open]').innerText({ timeout: 1000 }).catch(() => ''))) throw new Error('a conta de teste está pedindo troca de senha: entre uma vez com ela no navegador, troque a senha e atualize o segredo LEVE_SENHA');
    ok('login da conta de teste');
    // 3) regras com login de cargo comum
    if (esperarRegras) {
      const uid = await pg.evaluate(async () => { const { Auth } = await import(new URL('js/dados.js', location.href).href); return Auth.usuarioAtual()?.uid || ''; });
      const meu = await tentar('obter', 'acessos', uid); meu.ok ? ok('regras: lê o próprio cadastro') : falha(`regras: não lê o próprio cadastro (${meu.msg})`);
      const todos = await tentar('listar', 'acessos'); negado(todos) ? ok('regras: não lista os cadastros dos outros') : falha(`regras: cargo comum LISTOU os cadastros (${JSON.stringify(todos)}) — use uma conta que não seja administrador`);
      const logins = await tentar('listar', 'loginUsuarios'); negado(logins) ? ok('regras: lista de logins fechada') : falha(`regras: lista de logins aberta (${JSON.stringify(logins)})`);
    }
    // 4) todas as telas do menu da conta
    const hashes = await pg.evaluate(() => [...new Set([...document.querySelectorAll('a[href^="#"]')].map(a => a.getAttribute('href')).filter(h => h && h.length > 1))]);
    let estouro = [];
    for (const h of hashes) {
      onde = h; await pg.evaluate(x => { location.hash = x; }, h); await pg.waitForTimeout(2500);
      const [w, j] = await pg.evaluate(() => [document.documentElement.scrollWidth, innerWidth]); if (w > j + 2) estouro.push(`${h} (${w}>${j})`);
    }
    ok(`${hashes.length} telas do menu abertas`);
    estouro.length ? falha('telas que estouram a largura: ' + estouro.join(', ')) : ok('nenhuma tela estoura a largura');
    const unicos = a => [...new Set(a)];
    unicos(negadas).length ? falha('permissão negada em tela que o menu mostra (cargo × regras do Firestore): ' + unicos(negadas).slice(0, 8).join(' | ')) : ok('nenhuma permissão negada nas telas do menu');
    unicos(erros).length ? falha('erros de script: ' + unicos(erros).slice(0, 8).join(' | ')) : ok('nenhum erro de script');
    // 5) sair
    onde = 'sair'; await pg.click('#sair').catch(() => {});
    (await pg.waitForSelector('#form-login', { state: 'visible', timeout: 15000 }).then(() => true).catch(() => false)) ? ok('sair volta para o login') : falha('sair não voltou para o login');
  } catch (e) { falha(e.message); }
  finally { await nav.close(); }
  return falhas;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const celular = process.argv.includes('celular');
  console.log(`Teste no site de verdade — ${celular ? 'celular' : 'Full HD'}`);
  const falhas = await rodar({ celular }).catch(e => [e.message]);
  if (falhas.length) { console.error('\nFALHAS:\n- ' + falhas.join('\n- ')); process.exit(1); }
  console.log('\nsite de verdade: tudo certo');
}
