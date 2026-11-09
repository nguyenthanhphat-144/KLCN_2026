/* Progressive enhancements; all original server forms remain functional. */
document.addEventListener('DOMContentLoaded', () => {
    const body = document.body;
    const toggle = document.querySelector('.ui-menu-toggle');
    const backdrop = document.querySelector('.ui-sidebar-backdrop');
    const setMenu = open => {
        body.classList.toggle('ui-sidebar-open', open);
        toggle?.setAttribute('aria-expanded', String(open));
        toggle?.setAttribute('aria-label', open ? 'Đóng menu' : 'Mở menu');
        if (backdrop) backdrop.hidden = !open;
        const sidebar=document.getElementById('workspaceSidebar');
        if(sidebar) sidebar.inert=innerWidth<992 && !open;
    };
    toggle?.addEventListener('click', () => setMenu(!body.classList.contains('ui-sidebar-open')));
    backdrop?.addEventListener('click', () => {setMenu(false);toggle?.focus();});
    document.addEventListener('keydown', event => {if(event.key === 'Escape'){setMenu(false);}});
    window.addEventListener('resize', () => {if(innerWidth >= 992)setMenu(false);});
    setMenu(false);

    function setupTabs(bar, panels, labels) {
        if (!bar || !panels.length) return;
        bar.hidden = false;
        bar.setAttribute('role','tablist');
        const buttons = labels.map((label,index) => {
            const button = document.createElement('button');
            button.type='button';button.textContent=label;
            button.setAttribute('role','tab');button.id=bar.id+'-tab-'+index;
            panels[index].id=panels[index].id || bar.id+'-panel-'+index;
            panels[index].setAttribute('role','tabpanel');
            panels[index].setAttribute('aria-labelledby',button.id);
            button.setAttribute('aria-controls',panels[index].id);
            bar.appendChild(button);return button;
        });
        const select = index => {
            buttons.forEach((button,i) => {button.classList.toggle('active',i===index);button.setAttribute('aria-selected',String(i===index));button.tabIndex=i===index?0:-1;panels[i].hidden=i!==index;});
            window.dispatchEvent(new Event('resize'));
        };
        buttons.forEach((button,i) => {
            button.addEventListener('click',()=>select(i));
            button.addEventListener('keydown',event=>{let next=i;if(event.key==='ArrowRight')next=(i+1)%buttons.length;else if(event.key==='ArrowLeft')next=(i-1+buttons.length)%buttons.length;else if(event.key==='Home')next=0;else if(event.key==='End')next=buttons.length-1;else return;event.preventDefault();select(next);buttons[next].focus();});
        });
        select(0);
        // Bring a tab containing a server-side validation error into view.
        document.addEventListener('invalid',event=>{const index=panels.findIndex(p=>p.contains(event.target));if(index>=0)select(index);},true);
    }
    const catalogs=document.getElementById('catalogPanels');
    if(catalogs){const panels=[...catalogs.children].filter(node=>[...node.classList].some(c=>c.startsWith('col-')));panels.forEach(p=>p.classList.add('ui-catalog-panel'));setupTabs(document.getElementById('catalogTabs'),panels,['Lĩnh vực nghiên cứu','Loại công trình','Tạp chí / Hội thảo','Tác giả'].slice(0,panels.length));}
    const aiTabs=document.getElementById('aiTabs');
    if(aiTabs){aiTabs.replaceChildren();setupTabs(aiTabs,[...document.querySelectorAll('[data-ai-section]')],['Phân loại văn bản','Hàng đợi & cảnh báo','Mô hình bất thường']);}

    const userSearch=document.getElementById('userSearch');
    const roleFilter=document.getElementById('userRoleFilter');
    if(userSearch){
        const rows=[...document.querySelectorAll('main table tbody tr')].filter(r=>r.cells.length>1);
        const normalize=value=>value.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').toLowerCase();
        const filter=()=>{let count=0;rows.forEach(row=>{const role=row.cells[4]?.textContent.trim()||'';const text=[...row.cells].slice(0,5).map(cell=>cell.textContent).join(' ');const visible=normalize(text).includes(normalize(userSearch.value.trim()))&&(!roleFilter.value||role===roleFilter.value);row.hidden=!visible;if(visible)count++;});document.getElementById('userFilterCount').textContent=count+' / '+rows.length+' tài khoản';};
        userSearch.addEventListener('input',filter);roleFilter.addEventListener('change',filter);filter();
    }
    const aiToggle=document.getElementById('toggleAiColumns');
    if(aiToggle){const table=aiToggle.parentElement.nextElementSibling?.querySelector('table');if(table){table.classList.add('ui-compact-ai');aiToggle.addEventListener('click',()=>{const compact=table.classList.toggle('ui-compact-ai');aiToggle.setAttribute('aria-expanded',String(!compact));aiToggle.innerHTML='<i class="bi bi-layout-three-columns me-1"></i>'+(compact?'Hiện cột phân tích AI':'Thu gọn cột phân tích AI');});}}
    // Refresh Chart.js sizing after switching sections; no values are changed.
    document.querySelectorAll('canvas').forEach(canvas=>{const chart=window.Chart?.getChart?.(canvas);if(chart){chart.options.font={family:'Segoe UI, Arial, sans-serif',size:11};if(chart.config.type==='bar'){chart.data.datasets.forEach(dataset=>{dataset.borderRadius=6;dataset.maxBarThickness=48;});}chart.resize();chart.update('none');}});
});
