import { ModalOverlay } from './ModalOverlay.jsx';
export function AccessPanel({ onClose }) {
  const rows = window.PainelAccess.list().slice().reverse();
  return <ModalOverlay onDismiss={onClose} onClick={onClose} aria-label="Acessos deste navegador" className="gaq-access-overlay">
    <section className="gaq-access-panel" onClick={event => event.stopPropagation()}>
      <header><h2>Acessos deste navegador</h2><button className="gaq-btn" onClick={onClose}>Fechar</button></header>
      <p>Histórico local dos últimos 90 dias, limitado a 500 entradas. Não reúne acessos de outros computadores. Perfis com senha compartilhada não identificam uma pessoa individualmente.</p>
      <button className="gaq-btn" onClick={() => window.PainelAccess.download()}>Exportar CSV</button>
      <div className="gaq-table-scroll"><table className="gaq-table"><thead><tr><th>Quando</th><th>Perfil</th><th>Tipo</th><th>Evento</th></tr></thead><tbody>
        {rows.map((row, i) => <tr key={row.at + ':' + i}><td>{new Date(row.at).toLocaleString('pt-BR')}</td><td>{row.profile}</td><td>{row.type}</td><td>{row.event}</td></tr>)}
      </tbody></table></div>
      {!rows.length && <p>Nenhum acesso registrado neste navegador.</p>}
    </section>
  </ModalOverlay>;
}
