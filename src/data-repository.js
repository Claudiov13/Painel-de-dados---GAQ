/* Classic scripts in an isolated frame preserve file:// compatibility.
 * Nothing reaches the active application until every source matches the manifest.
 */
(function (global) {
  const sources = { dados: '__PAINEL_DADOS__', tags: '__PAINEL_TAGS__', servicedesk: '__SERVICE_DESK__', fracionamento: '__FRACIONAMENTO__' };
  let current = null, pending = null;
  function loadScript(doc, file) {
    return new Promise((resolve, reject) => {
      const script = doc.createElement('script');
      const timer = setTimeout(() => finish(new Error('Tempo esgotado ao ler ' + file)), 30000);
      function finish(error) { clearTimeout(timer); script.onload = script.onerror = null; error ? reject(error) : resolve(); }
      script.src = new URL(file + '?v=' + Date.now(), document.baseURI).href;
      script.onload = () => finish();
      script.onerror = () => finish(new Error('Arquivo indisponível: ' + file + '. Aguarde a sincronização do OneDrive.'));
      doc.head.appendChild(script);
    });
  }
  function validate(manifest, values, meta) {
    if (!manifest || manifest.schemaVersion !== 1 || !manifest.revision || !manifest.sources) throw new Error('Publicação de dados inválida.');
    for (const key of Object.keys(sources)) {
      const expected = manifest.sources[key];
      if (!expected || !meta[key] || meta[key].revision !== expected.revision) throw new Error('A base ' + key + ' ainda não sincronizou com esta publicação. Tente novamente.');
      const value = values[key];
      const count = key === 'dados' ? (Array.isArray(value) ? value.length : -1)
        : key === 'fracionamento' ? (Array.isArray(value?.itens) ? value.itens.length : -1)
        : value && typeof value === 'object' && !Array.isArray(value) ? Object.keys(value).length : -1;
      if (count < 0 || count !== expected.count || (key === 'dados' && !count)) throw new Error('Conteúdo inválido na base ' + key + '.');
    }
  }
  async function read(force) {
    const frame = document.createElement('iframe');
    frame.hidden = true; frame.title = 'Leitura de dados';
    document.body.appendChild(frame);
    try {
      const doc = frame.contentDocument, win = frame.contentWindow;
      await loadScript(doc, 'publicacao.js');
      const manifest = win.__PAINEL_MANIFEST__;
      if (!manifest || manifest.schemaVersion !== 1 || !manifest.revision) throw new Error('Manifesto de publicação inválido.');
      // A new day needs recalculation of aging even when files are unchanged.
      const day = new Date().toLocaleDateString('en-CA');
      if (!force && current && current.revision === manifest.revision && current.day === day) return { ...current, changed: false };
      for (const key of Object.keys(sources)) {
        const entry = manifest.sources[key];
        const files = { dados: 'dados.js', tags: 'tags.js', servicedesk: 'chamados_servicedesk.js', fracionamento: 'base_fracionamento.js' };
        if (!entry || entry.file !== files[key]) throw new Error('Arquivo não reconhecido no manifesto.');
      }
      // Sequential within the frame because legacy scripts assign their metadata object.
      const values = {}, meta = {};
      for (const [key, name] of Object.entries(sources)) {
        await loadScript(doc, manifest.sources[key].file);
        values[key] = win[name];
        meta[key] = win.__PAINEL_SOURCE_META__?.[key];
      }
      validate(manifest, values, meta);
      return { ...values, manifest, revision: manifest.revision, day, changed: true };
    } finally { frame.remove(); }
  }
  global.PainelRepository = {
    read(options = {}) {
      if (!pending) pending = read(Boolean(options.force)).finally(() => { pending = null; });
      return pending;
    },
    commit(snapshot) {
      for (const [key, name] of Object.entries(sources)) global[name] = snapshot[key];
      current = snapshot;
    },
    validate,
    metadata(snapshot, total) {
      const source = snapshot.manifest.sources.dados;
      return { nomeArq: 'dados.js', total, atualizadoEm: source.generatedAt ? new Date(source.generatedAt).toLocaleString('pt-BR') : 'Geração não informada',
        geradoEm: source.generatedAt, publicadoEm: snapshot.manifest.publishedAt, verificadoEm: new Date().toISOString(), versao: snapshot.revision };
    }
  };
})(window);
