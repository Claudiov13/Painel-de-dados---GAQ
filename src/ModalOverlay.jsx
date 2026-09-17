const modalStack = [];
export function ModalOverlay({ children, onDismiss, ...props }) {
  const node = React.useRef(null);
  const dismiss = React.useRef(onDismiss);
  dismiss.current = onDismiss;
  React.useEffect(() => {
    const previous = document.activeElement;
    const element = node.current;
    modalStack.push(element);
    const focusables = () => [...element.querySelectorAll('button, input, select, textarea, a[href], [tabindex="0"]')].filter(el => !el.disabled && el.getClientRects().length);
    (focusables()[0] || element).focus();
    const handle = event => {
      if (modalStack[modalStack.length - 1] !== element) return;
      if (event.key === 'Escape') { event.preventDefault(); event.stopImmediatePropagation(); dismiss.current?.(); }
      if (event.key === 'Tab') {
        const items = focusables(), first = items[0] || element, last = items.at(-1) || element;
        if (event.shiftKey && (document.activeElement === first || !element.contains(document.activeElement))) { event.preventDefault(); last.focus(); }
        else if (!event.shiftKey && (document.activeElement === last || !element.contains(document.activeElement))) { event.preventDefault(); first.focus(); }
      }
    };
    document.addEventListener('keydown', handle, true);
    const oldOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      modalStack.splice(modalStack.indexOf(element), 1);
      document.removeEventListener('keydown', handle, true);
      document.body.style.overflow = oldOverflow;
      if (previous?.isConnected) previous.focus();
    };
  }, []);
  return <div {...props} ref={node} role="dialog" aria-modal="true" aria-label={props['aria-label'] || 'Detalhes'} tabIndex={-1}>{children}</div>;
}
