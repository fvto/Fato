'use strict';
// Presentation only: no network requests, execution events or uploaded data.
(() => {
  const menu = document.getElementById('workspace-menu');
  const trigger = document.getElementById('site-menu-open');
  const close = document.getElementById('site-menu-close');
  if (menu && trigger && close) {
    let closing = false;
    function closeMenu() {
      if (!menu.open || closing) return;
      if (window.matchMedia('(prefers-reduced-motion: reduce)').matches || !menu.animate) {
        menu.close();
        return;
      }
      closing = true;
      const exit = menu.animate([
        {opacity: 1, clipPath: 'inset(0 0 0 0 round 10px)'},
        {opacity: 0, clipPath: 'inset(0 0 80% 0 round 10px)'}
      ], {duration: 160, easing: 'cubic-bezier(.4,0,.7,1)'});
      exit.finished.catch(() => {}).then(() => {
        menu.close();
        closing = false;
      });
    }
    trigger.addEventListener('click', () => {
      if (!menu.open) menu.showModal();
      trigger.setAttribute('aria-expanded', 'true');
    });
    close.addEventListener('click', closeMenu);
    menu.addEventListener('cancel', event => { event.preventDefault(); closeMenu(); });
    menu.addEventListener('close', () => {
      trigger.setAttribute('aria-expanded', 'false');
      trigger.focus({preventScroll: true});
    });
    menu.addEventListener('click', event => {
      if (event.target === menu) {
        const bounds = menu.getBoundingClientRect();
        if (event.clientX < bounds.left || event.clientX > bounds.right ||
            event.clientY < bounds.top || event.clientY > bounds.bottom) closeMenu();
      }
    });
  }
  const animations = new WeakMap();
  window.ChingluhUI = {
    transition(element) {
      if (!element || window.matchMedia('(prefers-reduced-motion: reduce)').matches || !element.animate) return;
      animations.get(element)?.cancel();
      animations.set(element, element.animate([
        {opacity: .65, transform: 'translateY(8px)'},
        {opacity: 1, transform: 'translateY(0)'}
      ], {duration: 240, easing: 'cubic-bezier(.2,.75,.25,1)'}));
    }
  };
})();
