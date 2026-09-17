(function (global) {
  const key = 'painel_acessos_locais_v1';
  const max = 500, retention = 90 * 86400000;
  function list() {
    try {
      const rows = JSON.parse(localStorage.getItem(key) || '[]');
      return Array.isArray(rows) ? rows.filter(r => r && Date.parse(r.at) >= Date.now() - retention).slice(-max) : [];
    } catch { return []; }
  }
  global.PainelAccess = {
    list,
    record(profile, type, event) {
      const row = { at: new Date().toISOString(), profile: String(profile || 'Administrador'), type, event };
      try { localStorage.setItem(key, JSON.stringify([...list(), row].slice(-max))); return true; } catch { return false; }
    },
    download() {
      const cell = value => '"' + String(value ?? '').replace(/^[=+@-]/, "'$&").replaceAll('"', '""') + '"';
      const text = '\uFEFF' + ['Data/hora;Perfil informado;Tipo;Evento', ...list().map(r => [r.at, r.profile, r.type, r.event].map(cell).join(';'))].join('\r\n');
      const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
      const link = document.createElement('a'); link.href = url; link.download = 'acessos-deste-navegador.csv'; link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  };
})(window);
