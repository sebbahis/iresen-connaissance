import * as C from './core.js';

export function mountPortal({ snapshot, spreadsheetUrl, accountKey, readFolder }) {
const $ = selector => document.querySelector(selector);
const E = C.escape;
const T = C.text;
const MASTER = spreadsheetUrl;
const folderEntries = new Map();
const paths = {
 home:'M3 10 12 3l9 7v10a1 1 0 0 1-1 1h-5v-7H9v7H4a1 1 0 0 1-1-1Z',
 layers:'m12 3 10 5-10 5L2 8Zm-10 9 10 5 10-5M2 16l10 5 10-5',
 book:'M3 4h6a3 3 0 0 1 3 3v14a4 4 0 0 0-4-2H3Zm18 0h-6a3 3 0 0 0-3 3v14a4 4 0 0 1 4-2h5Z',
 flask:'M9 3h6m-5 0v6L4 19a1.3 1.3 0 0 0 1 2h14a1.3 1.3 0 0 0 1-2l-6-10V3M7 14h10',
 patent:'m12 3 8 4v6c0 5-8 8-8 8s-8-3-8-8V7Zm-4 9 3 3 5-6',
 branch:'M6 3v12a4 4 0 0 0 4 4h8M6 7h8a4 4 0 0 1 4 4v8M3 3h6M15 19h6',
 chart:'M4 3v18h17M8 17v-4m5 4V7m5 10v-7',
 folder:'M3 7V5a1 1 0 0 1 1-1h6l2 3h8a1 1 0 0 1 1 1v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1Z',
 link:'m10 13 4-4m-6 6-1 1a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 3 1-1a4 4 0 0 1 6 6l-4 4a4 4 0 0 1-6 0',
 alert:'m12 3 10 18H2ZM12 9v5m0 3v.1',
 building:'M4 21V5l8-3 8 3v16M2 21h20M8 7h1m6 0h1M8 11h1m6 0h1M8 15h1m6 0h1m-6 6v-3h4v3',
 star:'m12 3 2.8 5.7 6.2.9-4.5 4.4 1.1 6.2-5.6-3-5.6 3 1.1-6.2L3 9.6l6.2-.9Z',
 search:'M21 21l-5-5m2-6a7 7 0 1 1-14 0 7 7 0 0 1 14 0',
 menu:'M3 6h18M3 12h18M3 18h18',
 close:'m6 6 12 12M6 18 18 6',
 download:'M12 3v12m-5-5 5 5 5-5M4 16v5h16v-5',
 file:'M5 3h9l5 5v13H5Zm9 0v6h5M8 13h8M8 17h6',
 clock:'M12 7v5l3 2M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0',
 edit:'m15 4 5 5-11 11-6 1 1-6Zm-3 3 5 5',
 chevron:'m9 5 7 7-7 7',
 refresh:'M20 8a8 8 0 1 0 0 8M20 3v5h-5',
 external:'M14 3h7v7m0-7-12 12M10 3H3v18h18v-7'
};
const icon = (name, size=19) => '<svg width="'+size+'" height="'+size+'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="'+(paths[name]||paths.file)+'"/></svg>';
const sections = {
 accueil:{label:'Vue d’ensemble',icon:'home'},
 projets:{label:'Projets AAP',icon:'layers',title:'Le répertoire des projets',eyebrow:'Recherche & innovation',titleField:'Projet',sub:'Intitulé développé',cols:['Projet','Programme ou appel','État du projet'],filter:'État du projet'},
 appels:{label:'Appels & budgets',icon:'branch',title:'Lire les appels et leurs enveloppes',eyebrow:'Programmes de soutien',titleField:'Programme ou appel',cols:['Programme ou appel','Année — recensement','Projets recensés','Enveloppe numérique — source','Devise'],filter:'Devise'},
 editions:{label:'Éditions R&D',icon:'branch',title:'Les 20 éditions de programmes R&D',eyebrow:'Chronologie documentée',titleField:'Appel ou édition',cols:['Appel ou édition','Lancement — source','Conserver']},
 labos:{label:'Labos & équipements',icon:'flask',title:'Laboratoires & équipements',eyebrow:'Moyens expérimentaux',titleField:'Laboratoire ou plateforme — source',sub:'Projet rapproché ou proposé',cols:['Laboratoire ou plateforme — source','Projet rapproché ou proposé','Établissement(s) — source','Ville(s) — source'],filter:'Ville(s) — source'},
 organismes:{label:'Organismes & sites',icon:'building',title:'Les bénéficiaires et leurs sites',eyebrow:'Réseau de recherche',titleField:'Établissement — audit',sub:'Organisme de rattachement',cols:['Établissement — audit','Projet','Ville — audit','Affectation des équipements'],filter:'Ville — audit'},
 publications:{label:'Publications',icon:'book',title:'Les publications scientifiques',eyebrow:'Production scientifique',titleField:'Titre de la publication',sub:'Revue ou éditeur',cols:['Titre de la publication','Année','Projet lié','Agence de moyens — audit'],filter:'Prise en compte'},
 brevets:{label:'Brevets',icon:'patent',title:'Le répertoire des brevets',eyebrow:'Propriété intellectuelle',titleField:'Titre figurant dans la source',sub:'N° de dépôt',cols:['Titre figurant dans la source','N° de dépôt','Projet lié','Situation du dépôt'],filter:'Situation du dépôt'},
 gep:{label:'Green Energy Park',icon:'flask'},
 indicateurs:{label:'Chiffres & périmètres',icon:'chart',title:'Chaque chiffre, avec son contexte',eyebrow:'Indicateurs documentés',titleField:'Indicateur',sub:'Période indiquée',cols:['Indicateur','Périmètre','Valeur','Unité','Nature du chiffre'],filter:'Périmètre'},
 documents:{label:'Documents & archives',icon:'folder'},
 sources:{label:'Sources & preuves',icon:'link',title:'L’origine de chaque information',eyebrow:'Traçabilité',titleField:'Source',sub:'Ce que la source couvre',cols:['Source','Type de source','Date du document','Examen réalisé'],filter:'Type de source'},
 verification:{label:'À vérifier',icon:'alert',title:'Les points qui restent à confirmer',eyebrow:'Fiabilisation',titleField:'Sujet',sub:'Constat utile',cols:['Sujet','Prochaine action','Priorité'],filter:'Priorité'},
 favoris:{label:'Mes favoris',icon:'star'},
 contributions:{label:'Proposer une correction',icon:'edit'},
 aide:{label:'Aide & mises à jour',icon:'file'}
};
const data = C.validate(snapshot);
let route, lastRows = [], lastTable = '', sortField = '', sortDir = 1;
let favorites = [];
const favoritesKey = 'iresen:session:favorites:' + accountKey;
try {
  favorites = JSON.parse(sessionStorage.getItem(favoritesKey) || '[]');
  if (!Array.isArray(favorites)) favorites = [];
} catch {}
function out(url, label, cls='') { if (/^#\/[a-z]+\/[^?#]+$/.test(T(url))) return '<a class="'+E(cls)+'" href="'+E(url)+'">'+label+'</a>'; const safe=C.safeUrl(url); return safe?'<a class="'+E(cls)+'" href="'+E(safe)+'" target="_blank" rel="noopener noreferrer">'+label+'</a>':E(label); }
function href(section, params={}, id='') { const qs=new URLSearchParams(Object.entries(params).filter(([,v])=>v!==''&&v!=null)); return '#/'+section+(id?'/'+encodeURIComponent(id):'')+(qs.size?'?'+qs:''); }
function jump(section, params={}, id='') { location.hash=href(section,params,id); }
function date(v) { if(!v)return 'Non précisée'; const d=new Date(v); return Number.isNaN(d.getTime())?E(v):new Intl.DateTimeFormat('fr-FR',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'}).format(d); }
function number(v, field='') { if (/année/i.test(field) && typeof v==='number') return String(v); return typeof v==='number'?new Intl.NumberFormat('fr-FR',{maximumFractionDigits:2}).format(v):E(v??'—'); }
function title(table,r) { return T(r.values[sections[table]?.titleField]||r.id); }
function badge(v, tone='') { return '<span class="badge '+tone+'">'+E(v)+'</span>'; }
function toast(message) { $('#toast').textContent=message;$('#toast').hidden=false;clearTimeout(toast.timer);toast.timer=setTimeout(()=>$('#toast').hidden=true,3500); }
function workbookSheet(name, edit = false) {
  const link = new URL(MASTER);
  if (link.pathname.toLowerCase().endsWith('/doc.aspx')) {
    link.searchParams.set('action', edit ? 'edit' : 'view');
    link.searchParams.set('activeCell', "'" + name + "'!A1");
  }
  return link.href;
}
function heading(titleText, desc, eyebrow, actions='') { return '<div class="heading"><div><p class="eyebrow">'+E(eyebrow)+'</p><h1>'+E(titleText)+'</h1><p class="intro">'+E(desc)+'</p></div>'+(actions?'<div class="actions">'+actions+'</div>':'')+'</div>'; }
function stat(label,n,sub,section,ico) { return '<a class="stat" href="'+href(section)+'"><span class="stat-head">'+icon(ico)+E(label)+'</span><strong>'+number(n)+'</strong><small>'+E(sub)+'</small></a>'; }
function card(section,desc,tag) { const s=sections[section]; return '<a class="card" href="'+href(section)+'"><span class="card-icon">'+icon(s.icon,22)+'</span><h3>'+E(s.label)+'</h3><p>'+E(desc)+'</p>'+(tag?badge(tag):'')+'</a>'; }
function renderNav(section) {
 const st=C.stats(data);
 $('#navigation').innerHTML=Object.entries(sections).filter(([id])=>!['editions','organismes','contributions','aide'].includes(id)).map(([id,s])=>'<a href="'+href(id)+'" class="nav-link '+(id===section?'active ':'')+(id==='gep'?'gep':'')+'" '+(id===section?'aria-current="page"':'')+'>'+icon(s.icon,18)+'<span>'+E(s.label)+'</span>'+(id==='projets'?'<em>'+st.projects+'</em>':id==='verification'?'<em>'+st.pending+'</em>':'')+'</a>').join('');
 $('#search-icon').innerHTML=icon('search',18);$('#menu').innerHTML=icon('menu',20);
 $('#breadcrumb').textContent='IRESEN / '+(sections[section]?.label||'Recherche');
 $('#footer').innerHTML='<span>Audit documentaire du '+date(data.meta.auditDate)+' · Classeur lu le '+date(data.meta.readAt)+'</span>'+out(MASTER,'Classeur maître dans SharePoint');
}
function renderHome() {
 const s=C.stats(data), pending=data.tables.verification.rows.filter(r=>r.values.Priorité==='P1').slice(0,3);
 const y=C.years(C.filterRows(data,'publications')); const max=Math.max(1,...y.map(v=>v[1]));
 return '<section class="hero"><p class="eyebrow">IRESEN · Base de connaissances</p><h1>Retrouver l’information.<br>Remonter à la source.</h1><p>Projets, publications, brevets et moyens expérimentaux : un accès commun aux connaissances et à leurs justificatifs.</p><div class="actions"><a class="btn" href="#/projets">'+icon('layers',17)+'Explorer les projets</a>'+out(MASTER,'Ouvrir le classeur maître','btn ghost')+'</div></section>'+
 '<div class="freshness">'+icon('clock',16)+'<span>Fichier source modifié le '+date(data.meta.sourceModified)+' · Audit du '+date(data.meta.auditDate)+'</span><a href="#/aide">Comprendre les mises à jour</a><button class="btn small" id="reload-data">Actualiser les données</button></div>'+
 '<section class="stats" aria-label="Repères de la base">'+stat('Projets AAP soutenus',s.projects,'Portefeuille audité, PRAGMA-E inclus','projets','layers')+stat('Labos soutenus',s.labs,'Repère institutionnel ; inventaire par projet','labos','flask')+stat('Publications · Agence',s.agencyPubs,'Minimum documentaire audité','publications','book')+stat('Références brevets',s.patents,'Références utilisables, statuts datés','brevets','patent')+'</section>'+
 '<div class="section-title"><h2>Que souhaitez-vous retrouver ?</h2><a href="#/sources">'+s.sources+' références sources</a></div><section class="cards">'+card('projets','Consulter les appels, les projets et leurs productions.',s.projects+' projets AAP')+card('publications','Rechercher un titre, un auteur, un DOI ou un projet.','Filtres Agence · GEP · Corpus')+card('labos','Retrouver les équipements et les bénéficiaires recensés.',s.inventories+' inventaires par projet')+card('brevets','Lire les références, les titulaires et les statuts de dépôt.','Liens vers les registres')+card('gep','Explorer les publications et les indicateurs du GEP.','Identité Green Energy Park')+card('documents','Parcourir les sources, les fichiers courants et les archives.','Documents dans SharePoint')+'</section>'+
 '<div class="two-col"><section class="panel"><h2>Publications du périmètre Agence</h2><p>Répartition par année des '+s.agencyPubs+' références retenues par l’audit. Les autres périmètres se consultent séparément.</p><div class="chart" role="img" aria-label="'+E(y.map(v=>v[0]+': '+v[1]+' publications').join('; '))+'">'+y.map(([year,count])=>'<div class="bar" title="'+year+' : '+count+' publications"><b>'+count+'</b><i style="height:'+Math.max(2,count/max*105)+'px"></i><span>'+year+'</span></div>').join('')+'</div></section><section class="panel"><h2>Contribuer à une base fiable</h2><ul class="panel-list">'+pending.map(r=>'<li><a href="'+href('verification',{},r.id)+'">'+E(title('verification',r))+'</a>'+badge('À confirmer','amber')+'</li>').join('')+'</ul><div class="actions" style="margin-top:18px"><a class="btn small" href="#/contributions">Proposer une correction</a><a class="btn small" href="#/verification">Voir les '+s.pending+' points ouverts</a></div></section></div>'+
 '<details class="notice"><summary>Comment lire les 35 labos et les différents périmètres ?</summary><p>'+E(data.meta.institutionalLabsNote)+' Les '+number(s.agencyPubs)+' publications Agence, le corpus global et les publications GEP peuvent se recouper. Aucun total par addition n’est affiché.</p></details>';
}
function scopeOptions(table) {
 if(table==='publications')return [['agency','Agence · audit'],['gep','GEP · notices retenues'],['retained','Corpus global · avec réserves'],['confirmed','Corpus global · sans réserves'],['all','Toutes les notices']];
 if(table==='projets')return [['agency','70 projets AAP'],['all','Avec références hors AAP']];
 if(table==='brevets')return [['usable','Références utilisables'],['all','Avec anomalies exclues']];
 if(table==='indicateurs')return [['all','Tous les périmètres'],['gep','GEP']];
 return [];
}
function options(items, selected) { return items.map(([v,l])=>'<option value="'+E(v)+'" '+(v===selected?'selected':'')+'>'+E(l)+'</option>').join(''); }
function renderTable(table) {
 const config=sections[table], source=data.tables[table], p=route.params;
 const scopes=scopeOptions(table), scope=p.get('scope')||(scopes[0]?.[0]||'all');
 const opts={scope,query:p.get('q')||'',year:p.get('year')||'',project:p.get('project')||'',field:config.filter,value:p.get('filter')||''};
 let rows=C.filterRows(data,table,opts);
 const allForFilters=C.filterRows(data,table,{scope});
 if(sortField)rows.sort((a,b)=>{const va=a.values[sortField],vb=b.values[sortField];return sortDir*(typeof va==='number'&&typeof vb==='number'?va-vb:T(va).localeCompare(T(vb),'fr',{numeric:true}));});
 lastRows=rows;lastTable=table;
 const pageSize=20,pages=Math.max(1,Math.ceil(rows.length/pageSize)); let page=Math.min(pages,Math.max(1,Number(p.get('page'))||1));const slice=rows.slice((page-1)*pageSize,page*pageSize);
 const filterValues=config.filter?[...new Set(allForFilters.map(r=>T(r.values[config.filter])).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'fr')):[];
 const yearValues=table==='publications'?[...new Set(allForFilters.map(r=>T(r.values.Année)).filter(Boolean))].sort().reverse():[];
 const counts=C.stats(data);
 const intro=table==='publications'?(scope==='gep'?'Notices sans réserve avec une mention GEP ou Green Energy Park dans le champ combiné « Programme ou plateforme ». Ce filtre ne remplace pas le décompte institutionnel des affiliations.':scope==='agency'?counts.agencyPubs+' publications dans le périmètre documentaire Agence de moyens. Les rattachements et les décisions de l’audit sont conservés dans chaque fiche.':scope==='retained'?counts.retainedPubs+' références retenues dans le corpus global. Les réserves restent visibles dans les fiches.':scope==='confirmed'?'Références retenues sans réserve dans le corpus global.':data.tables.publications.rows.length+' notices, avec les doublons, exclusions et réserves pour conserver la traçabilité du recensement.'):source.description;
 const columns=table==='publications'&&scope!=='agency'?config.cols.map(c=>c==='Agence de moyens — audit'?'Prise en compte':c):config.cols;
 const tableRows=slice.map(r=>'<tr><td class="row-actions"><button class="favorite '+(favorites.includes(table+'/'+r.id)?'saved':'')+'" data-fav="'+E(table+'/'+r.id)+'" aria-label="'+(favorites.includes(table+'/'+r.id)?'Retirer des favoris':'Ajouter aux favoris')+'" aria-pressed="'+favorites.includes(table+'/'+r.id)+'">'+icon('star',18)+'</button></td>'+columns.map((col,i)=>'<td>'+(i===0?'<a class="record-title" href="'+href(table,Object.fromEntries(p),r.id)+'">'+E(r.values[col]||'Non renseigné')+'</a>'+(config.sub?'<div class="cell-sub">'+E(r.values[config.sub]||'')+'</div>':'')+'<span class="record-ref">'+E(r.id)+'</span>':cell(r,col))+'</td>').join('')+'<td><a class="btn small" href="'+href(table,Object.fromEntries(p),r.id)+'" aria-label="Ouvrir la fiche '+E(title(table,r))+'">Fiche</a></td></tr>').join('');
 let note='';
 if(table==='labos')note='<div class="notice amber">Le repère institutionnel est de <strong>35 labos soutenus</strong>. Cet inventaire documente <strong>35 projets avec acquisitions</strong> et ne permet pas de compter les laboratoires physiques distincts. <a href="#/organismes">Consulter les 62 relations projet–bénéficiaire</a>.</div>';
 if(table==='appels')note='<div class="notice amber">Les unités monétaires ne sont pas toutes précisées dans la source. Aucun total des enveloppes n’est calculé. <a href="#/editions">Consulter la liste distincte des 20 éditions R&D</a>.</div>';
 if(table==='publications')note='<div class="notice">Périmètre affiché : <strong>'+E(scopes.find(s=>s[0]===scope)?.[1]||scope)+'</strong>. Une notice n’est pas automatiquement une publication du périmètre Agence. Les doublons et réserves restent consultables dans « Toutes les notices ».</div>';
 const projectNotice=opts.project?'<div class="notice">Rattachement projet recherché : <strong>'+E(opts.project)+'</strong>. <button class="inline-link" data-clear-project>Afficher tout le périmètre</button></div>':'';
 return heading(config.title,intro,config.eyebrow,out(workbookSheet(source.sheet),'Voir dans le classeur','btn'))+note+projectNotice+
 '<div class="toolbar"><label class="table-search">Rechercher dans cette rubrique<input id="table-query" type="search" value="'+E(opts.query)+'" placeholder="Titre, nom, référence…"></label>'+
 (scopes.length?'<label>Périmètre<select data-param="scope">'+options(scopes,scope)+'</select></label>':'')+
 (yearValues.length?'<label>Année<select data-param="year">'+options([['','Toutes les années'],...yearValues.map(y=>[y,y])],opts.year)+'</select></label>':'')+
 (filterValues.length?'<label>'+E(config.filter)+'<select data-param="filter">'+options([['','Tous'],...filterValues.map(v=>[v,v])],opts.value)+'</select></label>':'')+
 '<button class="btn export" data-export>'+icon('download',16)+'Exporter les résultats</button></div>'+
 (rows.length?'<div class="table-wrap"><table><thead><tr><th aria-label="Favori"></th>'+columns.map(col=>'<th scope="col" aria-sort="'+(sortField===col?(sortDir===1?'ascending':'descending'):'none')+'"><button data-sort="'+E(col)+'">'+E(col)+(sortField===col?(sortDir===1?' ⌃':' ⌄'):'')+'</button></th>').join('')+'<th scope="col">Détail</th></tr></thead><tbody>'+tableRows+'</tbody></table></div><div class="pagination"><span>'+number(rows.length)+' résultat'+(rows.length>1?'s':'')+' · '+((page-1)*pageSize+1)+'–'+Math.min(page*pageSize,rows.length)+'</span><div class="pages"><button class="btn small" data-page="'+(page-1)+'" '+(page===1?'disabled':'')+'>Précédent</button><span>'+page+' / '+pages+'</span><button class="btn small" data-page="'+(page+1)+'" '+(page===pages?'disabled':'')+'>Suivant</button></div></div>':empty('Aucun résultat','Modifiez les filtres ou le périmètre pour élargir la recherche.','Réinitialiser',href(table)));
}
function cell(r,col) { const v=r.values[col];if(v==null||v==='')return '<span class="muted">Non renseigné</span>';if(r.links[col])return out(r.links[col],E(v));if(['Situation du dépôt','État du projet','Agence de moyens — audit','Priorité'].includes(col))return badge(v,/confirmer|exclu|rejet|réserve|P1/i.test(T(v))?'amber':'gray');return number(v,col); }
function empty(head,desc,label,link) { return '<div class="empty"><h2>'+E(head)+'</h2><p>'+E(desc)+'</p>'+(link?'<a class="btn" href="'+E(link)+'">'+E(label)+'</a>':'')+'</div>'; }
function renderGep() {
 const s=C.stats(data), rows=data.tables.indicateurs.rows.filter(r=>r.values.Périmètre==='GEP'), find=id=>rows.find(r=>r.id===id);
 const communicated=rows.filter(r=>['OBS-0007','OBS-0008','OBS-0009','OBS-0011'].includes(r.id));
 const pubReference=find('OBS-0008'); const officialPublications=pubReference?number(pubReference.values.Valeur)+(pubReference.values.Précision==='Au moins'?'+':''):'Non renseigné';
 return '<div class="gep-heading"><img src="assets/gep.svg" alt="Green Energy Park"><p>Les connaissances du Green Energy Park, dans leur périmètre propre. Les références restent reliées aux sources de la base IRESEN.</p></div>'+heading('Green Energy Park','Consulter les productions et les repères institutionnels du GEP.','Plateforme technologique')+
 '<div class="stats">'+communicated.map(r=>'<a class="stat" href="'+href('indicateurs',{scope:'gep'},r.id)+'"><span class="stat-head">'+E(r.values.Indicateur)+'</span><strong>'+number(r.values.Valeur)+(r.values.Précision==='Au moins'?'+':'')+'</strong><small>Chiffre communiqué · '+E(r.values['Période indiquée'])+'</small></a>').join('')+'</div>'+
 '<div class="notice green">Les <strong>'+officialPublications+' publications communiquées</strong> et les <strong>'+s.gepPubs+' notices sans réserve rattachées au GEP dans le champ « Programme ou plateforme »</strong> répondent à des méthodes et dates de recensement différentes. Cette sélection repose sur le champ combiné du classeur et ne remplace pas un décompte fondé exclusivement sur les affiliations. Ces valeurs ne sont ni remplacées ni additionnées automatiquement.</div>'+
 '<div class="cards"><a class="card" href="'+href('publications',{scope:'gep'})+'"><span class="card-icon">'+icon('book')+'</span><h3>Publications GEP</h3><p>Titres, auteurs, années et notices avec un rattachement explicite au GEP.</p>'+badge(s.gepPubs+' notices retenues')+'</a><a class="card" href="'+href('indicateurs',{scope:'gep'})+'"><span class="card-icon">'+icon('chart')+'</span><h3>Indicateurs GEP</h3><p>Valeurs, unités, périodes et limites du recensement.</p>'+badge(rows.length+' repères')+'</a><a class="card" href="#/sources?q=GEP"><span class="card-icon">'+icon('link')+'</span><h3>Sources GEP</h3><p>Retrouver les pièces et les recensements qui documentent ces indicateurs.</p></a></div>'+
 '<div class="notice">Le portefeuille nominatif des 45 projets GEP n’est pas fourni dans le classeur courant. Le site conserve ce chiffre comme repère communiqué, sans créer de fiches projet non documentées.</div>';
}
function readFolderOnDemand(path) {
  if (folderEntries.has(path)) return;
  folderEntries.set(path, { loading: true });
  readFolder(path).then(folder => {
    folderEntries.set(path, { folder });
  }).catch(error => {
    const message = [401, 403, 404].includes(error.status) ? error.message : 'Le dossier SharePoint est temporairement indisponible.';
    folderEntries.set(path, { error: message });
  }).finally(() => {
    if (route.section === 'documents' && (route.params.get('path') || '') === path) render();
  });
}

function renderFiles() {
  const path = route.params.get('path') || '';
  readFolderOnDemand(path);
  const entry = folderEntries.get(path);
  const labels = {
    '02_Documents_sources': 'Sources et justificatifs',
    '90_Pour_equipe_informatique': 'Données et éléments techniques',
    '99_Archives': 'Archives historiques'
  };
  let current = '';
  const breadcrumbs = path.split('/').filter(Boolean).map(part => {
    current += (current ? '/' : '') + part;
    return icon('chevron', 13) + '<a href="' + href('documents', { path: current }) + '">' +
      E(labels[part] || part) + '</a>';
  }).join('');
  const bread = '<nav class="files-path" aria-label="Dossiers SharePoint"><a href="#/documents">Documents de référence</a>' +
    breadcrumbs + '</nav>';
  const top = heading('Documents et archives', 'Parcourir les dossiers accessibles avec votre compte Microsoft 365.',
    'SharePoint', entry.folder ? out(entry.folder.url, 'Ouvrir dans SharePoint', 'btn') : '');
  if (entry.loading) return top + bread + '<div class="notice" role="status">Lecture du dossier SharePoint…</div>';
  if (entry.error) return top + bread + '<div class="notice amber" role="status">' +
    E(entry.error) + ' <button class="inline-link" id="reload-folder">Réessayer</button></div>';
  const { items, partial, readAt } = entry.folder;
  const search = route.params.get('q') || '';
  const matching = items.filter(file => C.normalize(file.name).includes(C.normalize(search)))
    .sort((a, b) => Number(b.folder) - Number(a.folder) || a.name.localeCompare(b.name, 'fr'));
  const rows = matching.map(file => {
    const child = path ? path + '/' + file.name : file.name;
    const archived = child.startsWith('99_Archives');
    const label = E(labels[file.name] || file.name);
    const nameLink = file.folder
      ? '<a class="record-title" href="' + href('documents', { path: child }) + '">' + label + '</a>'
      : out(file.url, label, 'record-title');
    const status = archived ? 'Archive' : file.folder ? 'Dossier' : file.name.endsWith('.zip') ? 'Copie figée' : 'Référence';
    return '<tr><td><div class="file-name">' + icon(file.folder ? 'folder' : 'file', 20) + nameLink +
      '</div></td><td class="nowrap">' + date(file.modified) + '</td><td>' + badge(status, archived ? 'amber' : 'gray') +
      '</td><td>' + out(file.url, 'SharePoint', 'btn small') + '</td></tr>';
  }).join('');
  return top + bread + '<div class="source-status">' + icon('clock', 17) +
    '<span>Dossier chargé le ' + date(readAt) + '.</span><button class="inline-link" id="reload-folder">Actualiser le dossier</button></div>' +
    (partial ? '<div class="notice">Ouvrez le dossier dans SharePoint pour voir les autres éléments.</div>' : '') +
    '<div class="toolbar"><label class="table-search">Rechercher dans ce dossier<input id="table-query" type="search" value="' + E(search) + '"' +
    ' placeholder="Nom du fichier ou du dossier…"></label></div>' +
    (matching.length ? '<div class="table-wrap"><table><thead><tr><th scope="col">Fichier ou dossier</th>' +
      '<th scope="col">Modifié le</th><th scope="col">Repère</th><th scope="col">Accès</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div>' :
      empty('Aucun fichier affiché', 'Vérifiez la recherche ou ouvrez un autre dossier.', '', ''));
}

function renderResults(query, favoritesOnly=false) {
 const results=[];
 Object.entries(data.tables).forEach(([table,t])=>{if(table==='contributions')return;const rows=favoritesOnly?t.rows.filter(r=>favorites.includes(table+'/'+r.id)):C.searchRows(t.rows,query);rows.forEach(r=>results.push({table,r}));});
 const page=Math.max(1,Number(route.params.get('page'))||1),pageSize=30,pages=Math.max(1,Math.ceil(results.length/pageSize)),currentPage=Math.min(page,pages);
 return heading(favoritesOnly?'Mes références favorites':'Une recherche, toutes les rubriques',favoritesOnly?'Une sélection conservée pour cette session.':'Recherche dans les titres, les auteurs, les projets, les équipements et les sources.','Consultation transversale')+
 '<p class="result-count">'+number(results.length)+' résultat'+(results.length>1?'s':'')+(query?' pour « '+E(query)+' »':'')+'</p>'+
 (results.length?results.slice((currentPage-1)*pageSize,currentPage*pageSize).map(({table,r})=>'<a class="result-card" href="'+href(table,{},r.id)+'">'+badge(sections[table]?.label||table)+'<strong>'+E(title(table,r))+'</strong><p>'+E(T(r.values[sections[table]?.sub]||Object.values(r.values)[1]||'').slice(0,240))+'</p><span class="record-ref">'+E(r.id)+'</span></a>').join('')+'<div class="pagination"><span>'+currentPage+' / '+pages+'</span><div class="pages"><button class="btn small" data-page="'+(currentPage-1)+'" '+(currentPage===1?'disabled':'')+'>Précédent</button><button class="btn small" data-page="'+(currentPage+1)+'" '+(currentPage===pages?'disabled':'')+'>Suivant</button></div></div>':empty(favoritesOnly?'Aucun favori pour le moment':'Aucune référence trouvée',favoritesOnly?'Utilisez l’étoile à côté d’une référence pour la retrouver ici.':'Essayez un nom de projet, un mot du titre ou une année.','Explorer les projets','#/projets'));
}
function renderHelp() {
  return heading('Utiliser la base de connaissances', 'Lire, rechercher et proposer une correction dans les sources SharePoint.', 'Mode d’emploi') +
    '<div class="help-grid"><section class="panel"><h2>Consulter une information</h2>' +
    '<p>Recherchez dans toutes les rubriques ou choisissez un périmètre. Ouvrez la fiche pour consulter les références et les justificatifs.</p></section>' +
    '<section class="panel"><h2>Classeur et documents SharePoint</h2>' +
    '<p>Le portail lit le classeur maître à l’ouverture et chaque dossier lors de sa consultation. L’accès dépend des autorisations de votre compte Microsoft 365.</p>' +
    '<p>Modification du classeur le ' + date(data.meta.sourceModified) + '. Audit documentaire du ' + date(data.meta.auditDate) + '.</p>' +
    '<button class="btn" id="reload-data">Relire le classeur SharePoint</button></section>' +
    '<section class="panel"><h2>Proposer une correction</h2>' +
    '<p>Renseignez la référence, la correction et le lien vers sa source dans le classeur maître.</p>' +
    out(workbookSheet('Contributions', true), 'Ouvrir Contributions dans SharePoint', 'btn primary') + '</section></div>';
}
function renderContributions() {
 return heading('Proposer une correction documentée','Un circuit commun pour éviter les corrections dispersées et les doublons.','Amélioration de la base')+
 '<section class="panel"><h2>Renseigner une ligne dans Contributions</h2><ol><li>La rubrique et l’identifiant de la référence concernée.</li><li>L’information à ajouter ou corriger.</li><li>Un lien vers le justificatif et sa date.</li><li>Votre nom, la date et l’avancement de la proposition.</li></ol>'+out(workbookSheet('Contributions',true),'Ouvrir l’onglet Contributions','btn primary')+'</section><div class="notice">Le site n’envoie pas de message et ne modifie pas les données source. La validation et l’intégration des propositions se font dans le classeur maître.</div>';
}
function renderDetail(table,id) {
 const r=data.tables[table]?.rows.find(x=>x.id===id);const dialog=$('#detail');
 if(!r){toast('Cette référence n’existe pas dans le classeur actuel.');return;}
 const green=C.isGep(table,r);dialog.classList.toggle('green-detail',green);
 const full=/précaution|observation|constat|équipement|auteurs|limites|contrôle|décision|intitulé|titre|preuve|prochaine action/i;
 const links=Object.entries(r.links).filter(([,url])=>C.safeUrl(url));
 const uniqueLinks=[...new Map(links.map(([label,url])=>[url,{label,url}])).values()];
 const relatedTargets=['projets','publications','brevets','labos','organismes'].filter(t=>t!==table);
 const relatedHtml=relatedTargets.map(t=>{const rows=C.related(data,table,r,t);if(!rows.length)return '';const project=C.projectNames(table,r)[0];return '<section class="related"><h3>'+E(sections[t].label)+' · '+rows.length+' référence'+(rows.length>1?'s':'')+'</h3>'+rows.slice(0,5).map(row=>'<a href="'+href(t,{},row.id)+'">'+E(title(t,row))+'</a>').join('')+(rows.length>5?'<a class="btn small" href="'+href(t,{project,scope:'all'})+'">Voir les références liées</a>':'')+'</section>';}).join('');
 $('#detail-content').innerHTML='<div class="dialog-head"'+(green?' style="border-top:4px solid #4d792b;background:#f4f8ee"':'')+'><div>'+(green?'<img src="assets/gep.svg" alt="Green Energy Park" style="width:110px;height:60px;object-fit:contain;display:block;margin-bottom:12px">':'')+'<span class="badge">'+E(sections[table]?.label||table)+'</span><h2 id="detail-title">'+E(title(table,r))+'</h2><span class="record-ref">'+E(r.id)+'</span></div><button class="icon-button" id="close-detail" aria-label="Fermer la fiche">'+icon('close')+'</button></div><div class="dialog-body"><dl class="detail-grid">'+data.tables[table].columns.map(col=>'<div class="detail-field '+(full.test(col)?'full':'')+'"><dt>'+E(col)+'</dt><dd>'+(r.links[col]?out(r.links[col],E(r.values[col]||'Ouvrir')):r.values[col]==null||r.values[col]===''?'<span class="unknown">Non renseigné</span>':number(r.values[col],col))+'</dd></div>').join('')+'</dl>'+(uniqueLinks.length?'<section class="related"><h3>Pièces et registres</h3><div class="actions">'+uniqueLinks.map(l=>out(l.url,E(l.label)+' '+icon('external',14),'btn small')).join('')+'</div></section>':'')+relatedHtml+'<div class="dialog-footer"><button class="btn" id="copy-link">Copier le lien de la fiche</button><button class="btn" data-fav="'+E(table+'/'+r.id)+'">'+icon('star',16)+(favorites.includes(table+'/'+r.id)?'Retirer des favoris':'Ajouter aux favoris')+'</button><button class="btn" id="print-detail">Imprimer la fiche</button></div><p class="cell-sub" style="margin-top:15px">Rapprochements fondés sur les libellés projet présents dans le classeur ; aucun lien non documenté n’est ajouté.</p></div>';
 if(!dialog.open)dialog.showModal();dialog.scrollTop=0;
 $('#close-detail').onclick=closeDetail;
 $('#copy-link').onclick=async()=>{try{if(!navigator.clipboard)throw new Error();await navigator.clipboard.writeText(location.href);toast('Lien de la fiche copié.');}catch(_){toast('Copiez l’adresse de la fiche dans la barre du navigateur.');}};
 $('#print-detail').onclick=()=>window.print();
}
function closeDetail() {const p=Object.fromEntries(route.params);jump(route.section,p);}
function render() {
 const active=document.activeElement;const restore=active?.id==='table-query'?{id:active.id,start:active.selectionStart}:null;
 const match=location.hash.replace(/^#\/?/,'')||'accueil';const [path,qs='']=match.split('?');const [section,id]=path.split('/');
 let reference = '';
 try { if (id) reference = decodeURIComponent(id); } catch {}
 route = {
   section: Object.hasOwn(sections, section) || section === 'recherche' ? section : 'accueil',
   id: reference,
   params: new URLSearchParams(qs)
 };
 const globalQuery=route.section==='recherche'?route.params.get('q')||'':'';
 if(document.activeElement!==$('#global-search'))$('#global-search').value=globalQuery;
 document.body.classList.toggle('green',route.section==='gep'||(['publications','indicateurs'].includes(route.section)&&route.params.get('scope')==='gep'));
 renderNav(route.section);
 let body;
 if(route.section==='accueil')body=renderHome();else if(route.section==='gep')body=renderGep();else if(route.section==='documents')body=renderFiles();else if(route.section==='aide')body=renderHelp();else if(route.section==='contributions')body=renderContributions();else if(route.section==='favoris')body=renderResults('',true);else if(route.section==='recherche')body=globalQuery.trim().length<2?empty('Que recherchez-vous ?','Saisissez au moins deux caractères dans la recherche générale.','',''):renderResults(globalQuery);else body=renderTable(route.section);
 if(['publications','indicateurs'].includes(route.section)&&route.params.get('scope')==='gep') body='<div class="gep-heading"><img src="assets/gep.svg" alt="Green Energy Park"><p>Vue du périmètre Green Energy Park. Les données et leurs justificatifs restent distincts des totaux Agence de moyens.</p></div>'+body;
 $('#content').innerHTML=body;document.title=(sections[route.section]?.label||'Recherche')+' · IRESEN';
 bind();
 if(restore){const input=$('#'+restore.id);if(input){input.focus();try{input.setSelectionRange(restore.start,restore.start);}catch(_){}}}
 if(route.id&&data.tables[route.section])renderDetail(route.section,route.id);else if($('#detail').open)$('#detail').close();
}
function updateParam(key,value) {const p=Object.fromEntries(route.params);if(value)p[key]=value;else delete p[key];if(key!=='page')delete p.page;jump(route.section,p);}
function download(content,name,type) {const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),2000);}
function bind() {
 document.querySelectorAll('[data-param]').forEach(el=>el.onchange=()=>updateParam(el.dataset.param,el.value));
 const input=$('#table-query');if(input)input.oninput=()=>{clearTimeout(bind.searchTimer);const v=input.value;bind.searchTimer=setTimeout(()=>updateParam('q',v),130);};
 document.querySelectorAll('[data-sort]').forEach(b=>b.onclick=()=>{sortDir=sortField===b.dataset.sort?-sortDir:1;sortField=b.dataset.sort;render();});
 document.querySelectorAll('[data-page]').forEach(b=>b.onclick=()=>{updateParam('page',b.dataset.page);$('#content').scrollIntoView({block:'start'});});
 const clear=$('[data-clear-project]');if(clear)clear.onclick=()=>updateParam('project','');
 const exp=$('[data-export]');if(exp)exp.onclick=()=>{if(!lastRows.length){toast('Aucun résultat à exporter.');return;}download(C.csv(lastRows,data.tables[lastTable].columns),'IRESEN_'+lastTable+'_'+new Date().toISOString().slice(0,10)+'.csv','text/csv;charset=utf-8');};
 const refresh = $('#reload-data');
 if (refresh) refresh.onclick = () => location.reload();
 const reloadFolder = $('#reload-folder');
 if (reloadFolder) reloadFolder.onclick = () => {
   const path = route.params.get('path') || '';
   folderEntries.delete(path);
   render();
 };
}
function mobile(open) {$('#sidebar').classList.toggle('open',open);$('#scrim').hidden=!open;$('#menu').setAttribute('aria-expanded',String(open));}
$('#menu').onclick=()=>mobile(!$('#sidebar').classList.contains('open'));$('#scrim').onclick=()=>mobile(false);
$('#global-search').oninput=e=>{clearTimeout(render.globalTimer);const q=e.target.value;render.globalTimer=setTimeout(()=>{const hash=href(q?'recherche':'accueil',q?{q}:{});location.replace(hash);},130);};
$('#detail').addEventListener('cancel',e=>{e.preventDefault();closeDetail();});
$('#detail').addEventListener('click',e=>{if(e.target===$('#detail')){const r=e.target.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)closeDetail();}});
window.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();$('#global-search').focus();}if(e.key==='Escape')mobile(false);});
document.addEventListener('click',e=>{const fav=e.target.closest('[data-fav]');if(fav){const k=fav.dataset.fav;favorites=favorites.includes(k)?favorites.filter(x=>x!==k):[...favorites,k];try{sessionStorage.setItem(favoritesKey, JSON.stringify(favorites));}catch(_){}render();toast(favorites.includes(k)?'Référence ajoutée aux favoris.':'Référence retirée des favoris.');}});
window.addEventListener('hashchange',()=>{mobile(false);render();});
render();
}
