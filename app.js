(function(){
  function ready(fn){ if(document.readyState!=='loading') fn(); else document.addEventListener('DOMContentLoaded', fn); }
  ready(function(){
    const menu = document.getElementById('gear-menu');
    const btn = document.getElementById('btn-gear');
    const act = document.getElementById('action-refresh');
    const inline = document.getElementById('btn-refresh-inline');
    const env = document.getElementById('env');
    env.textContent = 'URL: ' + location.href + ' · Cargado: ' + new Date().toLocaleString();
    btn.addEventListener('click', function(ev){
      const open = menu.getAttribute('aria-hidden') === 'false';
      menu.setAttribute('aria-hidden', open ? 'true' : 'false');
      btn.setAttribute('aria-expanded', open ? 'false' : 'true');
      ev.stopPropagation();
    });
    document.addEventListener('click', function(ev){
      if(!menu.contains(ev.target) && ev.target !== btn){
        menu.setAttribute('aria-hidden','true');
        btn.setAttribute('aria-expanded','false');
      }
    });
    function hardRefresh(){
      const url = new URL(window.location.href);
      url.searchParams.set('_r', String(Date.now()));
      window.location.replace(url.toString());
    }
    act.addEventListener('click', hardRefresh);
    inline.addEventListener('click', hardRefresh);
  });
})();
