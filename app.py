# -*- coding: utf-8 -*-
"""TraficLab : circulation alternée, simulation pédagogique sans données persistantes."""
from __future__ import annotations

import copy
import hashlib
import io
import json
import math
import time
from pathlib import Path
from html import escape
import zipfile

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import simpy
from scipy import stats
import streamlit as st
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

VERSION='1.0'
RUBY,CYAN,INK,GREEN='#B92046','#007D8A','#242B30','#25734F'
MODES=['Poisson','Régulières','Comptage fourni']
DEFAULT=dict(name='Mon alternat routier',minutes=15,rate_a=240.,rate_b=180.,green_a=40,
    green_b=40,clearance=20,headway=3.,length=120.,speed=30.,initial_a=0,initial_b=0,seed=42,mode='Poisson')
VEHICLE_COLUMNS=['véhicule','sens','origine','arrivée_s','entrée_s','sortie_théorique_s','attente_observée_s','attente_avant_entrée_s','état']


def number(value,label,low,high,integer=False):
    if isinstance(value,(bool,np.bool_)):raise ValueError(label+' : nombre attendu.')
    try:x=float(value)
    except (ValueError,TypeError,OverflowError):raise ValueError(label+' : renseigner une valeur numérique.') from None
    if not math.isfinite(x) or not low<=x<=high or integer and not x.is_integer():
        raise ValueError(label+' : valeur hors limites'+(' ou non entière.' if integer else '.'))
    return int(x) if integer else x


def check_config(raw):
    if not isinstance(raw,dict) or not set(DEFAULT)<=set(raw):raise ValueError('Paramètres du scénario incomplets.')
    c={k:raw[k] for k in DEFAULT}
    if not isinstance(c['name'],str) or not 1<=len(c['name'].strip())<=80 or any(ord(s)<32 for s in c['name']):
        raise ValueError('Nom du scénario : de 1 à 80 caractères, sur une seule ligne.')
    c['name']=c['name'].strip()
    for key,label,lo,hi,integer in [
        ('minutes','Durée (min)',1,60,True),('rate_a','Débit A',0,1500,False),('rate_b','Débit B',0,1500,False),
        ('green_a','Vert A',5,300,True),('green_b','Vert B',5,300,True),('clearance','Dégagement',1,600,True),
        ('headway','Espacement de passage',1,15,False),('length','Longueur',10,1000,False),('speed','Vitesse',5,60,False),
        ('initial_a','File initiale A',0,200,True),('initial_b','File initiale B',0,200,True),('seed','Graine',0,2147483647,True)]:
        c[key]=number(c[key],label,lo,hi,integer)
    if c['mode'] not in MODES:raise ValueError('Mode d’arrivée inconnu.')
    if min(c['green_a'],c['green_b'])<c['headway']:raise ValueError('Chaque vert doit permettre au moins un créneau de passage.')
    if c['clearance']+1e-9<travel_time(c):
        raise ValueError(f'Dégagement trop court pour ce modèle : au moins {math.ceil(travel_time(c))} s à vitesse constante. Ce minimum ne vaut pas validation de sécurité.')
    return c


def travel_time(c):return c['length']/(c['speed']/3.6)
def cycle_time(c):return c['green_a']+c['green_b']+2*c['clearance']
def capacity(c,sense):return math.floor(c['green_'+sense.lower()]/c['headway']+1e-9)*3600/cycle_time(c)
def fingerprint(c,events):return hashlib.sha256(json.dumps([c,events],sort_keys=True,ensure_ascii=False).encode()).hexdigest()[:16]
def fmt(v,d=0):return '—' if v is None else f'{v:,.{d}f}'.replace(',',' ').replace('.',',')


def check_events(records,c):
    if not isinstance(records,list) or len(records)>5000:raise ValueError('Maximum : 5 000 arrivées par scénario.')
    cleaned=[]
    for row in records:
        if not isinstance(row,dict) or not {'arrivee_s','sens'}<=row.keys():raise ValueError('Colonnes attendues : arrivee_s et sens.')
        t=number(row['arrivee_s'],'Seconde d’arrivée',0,c['minutes']*60)
        if t>=c['minutes']*60:raise ValueError('Chaque arrivée doit précéder strictement la fin de l’observation.')
        if not isinstance(row['sens'],str) or row['sens'].strip().upper() not in ('A','B'):raise ValueError('Sens attendu : A ou B.')
        cleaned.append(dict(arrivee_s=t,sens=row['sens'].strip().upper()))
    return sorted(cleaned,key=lambda r:(r['arrivee_s'],r['sens']))


def generate_events(c,seed=None):
    rng=np.random.default_rng(c['seed'] if seed is None else seed);horizon=c['minutes']*60;rows=[]
    for sense,key in [('A','rate_a'),('B','rate_b')]:
        rate=c[key]
        if rate==0:continue
        if c['mode']=='Régulières':times=np.arange(1800/rate,horizon,3600/rate)
        else:
            count=int(rng.poisson(rate*horizon/3600))
            if count>5000:raise ValueError('Trop d’arrivées : réduire le débit ou la durée.')
            times=np.sort(rng.uniform(0,horizon,count))
        rows.extend(dict(arrivee_s=float(t),sens=sense) for t in times)
    return check_events(rows,c)


def simulate(config,events=None):
    """Créneaux FIFO SimPy ; les files initiales précèdent les arrivées à t=0."""
    c=check_config(config)
    rows=check_events(events,c) if events is not None else generate_events(c)
    if c['mode']=='Comptage fourni' and events is None:raise ValueError('Renseigner ou importer les arrivées du comptage.')
    horizon=c['minutes']*60;cycle=cycle_time(c);transit=travel_time(c)
    vehicles=[]
    for sense in ('A','B'):
        for _ in range(c['initial_'+sense.lower()]):
            vehicles.append(dict(sens=sense,origin='File initiale',arrival=0.))
    vehicles.extend(dict(sens=r['sens'],origin='Arrivée fournie' if c['mode']=='Comptage fourni' else 'Arrivée simulée',arrival=r['arrivee_s']) for r in rows)
    vehicles.sort(key=lambda v:v['arrival'])
    for i,v in enumerate(vehicles):v.update(id=i+1,entry=float('nan'))
    env=simpy.Environment();queues={s:simpy.Store(env) for s in ('A','B')}

    def arrive(v):
        yield env.timeout(v['arrival'])
        yield queues[v['sens']].put(v)

    def serve(sense):
        green=c['green_'+sense.lower()];offset=0 if sense=='A' else c['green_a']+c['clearance']
        count=math.floor(green/c['headway']+1e-9)
        start=offset
        while start<horizon:
            for k in range(1,count+1):
                t=start+k*c['headway']
                if t>=horizon-1e-9:return
                yield env.timeout(max(0,t-env.now))
                if queues[sense].items:
                    v=yield queues[sense].get();v['entry']=float(env.now)
            start+=cycle

    for v in vehicles:env.process(arrive(v))
    for sense in ('A','B'):env.process(serve(sense))
    env.run(until=horizon)
    records=[]
    for v in vehicles:
        entered=math.isfinite(v['entry']);end=v['entry']+transit if entered else float('nan')
        records.append([v['id'],v['sens'],v['origin'],v['arrival'],v['entry'],end,
            (v['entry'] if entered else horizon)-v['arrival'],v['entry']-v['arrival'] if entered else float('nan'),
            'Sorti' if entered and end<=horizon+1e-9 else 'En traversée' if entered else 'En attente'])
    frame=pd.DataFrame(records,columns=VEHICLE_COLUMNS)
    times={0.,float(horizon)}
    for v in vehicles:
        times.add(v['arrival'])
        if math.isfinite(v['entry']):times.add(v['entry'])
    ts=np.array(sorted(times));timeline=pd.DataFrame({'seconde':ts,'minute':ts/60})
    for sense in ('A','B'):
        subset=frame[frame['sens']==sense]
        arrivals=np.sort(subset['arrivée_s'].to_numpy(dtype=float))
        entries=np.sort(subset['entrée_s'].dropna().to_numpy(dtype=float))
        timeline['file_'+sense]=np.searchsorted(arrivals,ts,side='right')-np.searchsorted(entries,ts,side='right')
    timeline['file_totale']=timeline['file_A']+timeline['file_B']
    admitted=frame['entrée_s'].notna();exited=frame['état']=='Sorti'
    summary=dict(total=len(frame),arrivals=len(rows),initial=c['initial_a']+c['initial_b'],admitted=int(admitted.sum()),
        exited=int(exited.sum()),crossing=int((frame['état']=='En traversée').sum()),remaining=int((~admitted).sum()),
        max_queue=int(timeline.file_totale.max()),mean_queue=float(np.dot(timeline.file_totale.to_numpy()[:-1],np.diff(ts))/horizon),
        mean_wait=float(frame['attente_observée_s'].mean()) if len(frame) else None,
        served_wait=float(frame.loc[admitted,'attente_avant_entrée_s'].mean()) if admitted.any() else None)
    return dict(config=c,events=rows,vehicles=frame,timeline=timeline,summary=summary)


@st.cache_data(max_entries=16,ttl=3600,show_spinner=False)
def cached_run(c,events):return simulate(c,events)


def compare(c,events,alternative):
    a=check_config(c);b=check_config(alternative)
    if a['minutes']!=b['minutes']:raise ValueError('Les scénarios doivent avoir la même durée d’observation.')
    if (a['initial_a'],a['initial_b'])!=(b['initial_a'],b['initial_b']):raise ValueError('Conserver les mêmes files initiales pour comparer les réglages.')
    arrivals=check_events(events,a)
    return simulate(a,arrivals),simulate(b,arrivals)


@st.cache_data(max_entries=8,ttl=3600,show_spinner=False)
def monte_carlo(config,repetitions,seed,threshold,sensitivity=False):
    c=check_config(config);n=number(repetitions,'Répétitions',20,200,True)
    seed=number(seed,'Graine Monte-Carlo',0,2147483647,True);threshold=number(threshold,'Seuil de file',0,5000,True)
    expected=(c['rate_a']+c['rate_b'])*c['minutes']/60+c['initial_a']+c['initial_b']
    if expected*n*(3 if sensitivity else 1)>600000:raise ValueError('Calcul trop volumineux : réduire la durée, les débits ou les répétitions.')
    c['mode']='Poisson';seeds=np.random.SeedSequence(seed).spawn(n);rows=[]
    factors=[.8,1.,1.2] if sensitivity else [1.]
    for factor in factors:
        q=copy.deepcopy(c);q['rate_a']*=factor;q['rate_b']*=factor
        if max(q['rate_a'],q['rate_b'])>1500:raise ValueError('Sensibilité +20 % : un débit dépasserait la limite de 1 500 véh/h. Réduire le débit de référence.')
        for i,s in enumerate(seeds):
            local=int(s.generate_state(1)[0]);r=simulate(q,generate_events(q,local))['summary']
            rows.append(dict(essai=i+1,facteur=factor,file_max=r['max_queue'],file_finale=r['remaining'],attente_observee_s=r['mean_wait'],sortis=r['exited']))
    return pd.DataFrame(rows)


def risk_stats(values,threshold):
    x=np.asarray(values,dtype=float)
    if not len(x) or not np.isfinite(x).all():raise ValueError('Échantillon probabiliste vide ou invalide.')
    n=len(x);k=int((x>threshold).sum());ci=stats.binomtest(k,n).proportion_ci(confidence_level=.95,method='wilson')
    return dict(probability=k/n,low=float(ci.low),high=float(ci.high),q50=float(np.quantile(x,.5,method='inverted_cdf')),q95=float(np.quantile(x,.95,method='inverted_cdf')))


def read_arrivals(raw,filename,c):
    if len(raw)>2*1024*1024:raise ValueError('Fichier limité à 2 Mo.')
    try:
        if filename.lower().endswith('.xlsx'):
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if sum(x.file_size for x in archive.infolist())>20*1024*1024:raise ValueError('Classeur décompressé trop volumineux.')
            f=pd.read_excel(io.BytesIO(raw),nrows=5001)
        else:f=pd.read_csv(io.BytesIO(raw),sep=None,engine='python',encoding='utf-8-sig',nrows=5001)
    except ValueError:raise
    except Exception:raise ValueError('Fichier CSV UTF-8 ou Excel XLSX illisible.') from None
    if not {'arrivee_s','sens'}<=set(f.columns):raise ValueError('Colonnes attendues : arrivee_s et sens.')
    return check_events(f[['arrivee_s','sens']].to_dict('records'),c)


def backup_bytes(study):
    return json.dumps(dict(version=1,config=study['config'],events=study['events']),ensure_ascii=False,indent=2,allow_nan=False).encode('utf-8')


def restore(raw):
    if len(raw)>1024*1024:raise ValueError('Sauvegarde limitée à 1 Mo.')
    try:data=json.loads(raw)
    except (ValueError,UnicodeError):raise ValueError('Sauvegarde JSON illisible.') from None
    if not isinstance(data,dict) or type(data.get('version')) is not int or data['version']!=1:raise ValueError('Version de sauvegarde non reconnue.')
    c=check_config(data.get('config'));events=check_events(data.get('events'),c)
    return dict(config=c,events=events,result=None)


def safe_frame(f):
    return f.map(lambda x:"'"+x if isinstance(x,str) and x.startswith(('=','+','-','@','\t','\r')) else x)


def csv_bytes(f):return safe_frame(f).to_csv(index=False).encode('utf-8-sig')


def excel_bytes(result):
    w=Workbook();w.remove(w.active)
    tables=[('Parametres',pd.DataFrame(result['config'].items(),columns=['Paramètre','Valeur'])),
        ('Indicateurs',pd.DataFrame(result['summary'].items(),columns=['Indicateur','Valeur'])),
        ('Arrivees',pd.DataFrame(result['events'],columns=['arrivee_s','sens'])),('Vehicules',result['vehicles']),('Files',result['timeline'])]
    for name,f in tables:
        ws=w.create_sheet(name);ws.append(list(f.columns))
        for row in safe_frame(f).itertuples(index=False,name=None):ws.append([None if pd.isna(v) else v for v in row])
        ws.freeze_panes='A2'
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=min(38,max(16,max(len(str(v.value or '')) for v in col)+2))
    out=io.BytesIO();w.save(out);return out.getvalue()


def pdf_bytes(result,workspace):
    font='Helvetica'
    for candidate in [Path('C:/Windows/Fonts/arial.ttf'),Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')]:
        if candidate.exists():
            if 'TraficFont' not in pdfmetrics.getRegisteredFontNames():pdfmetrics.registerFont(TTFont('TraficFont',str(candidate)))
            font='TraficFont';break
    styles=getSampleStyleSheet()
    for name in ['BodyText','Title','Heading2']:styles[name].fontName=font
    styles['BodyText'].fontSize=9;styles['BodyText'].leading=14
    out=io.BytesIO();c=result['config'];r=result['summary'];P=lambda x:Paragraph(escape(str(x)),styles['BodyText'])
    story=[Paragraph('TraficLab',styles['Title']),P(c['name']),P(workspace+' | '+c['mode']+' | '+str(c['minutes'])+' minutes'),Spacer(1,14)]
    def table(headers,rows,widths):
        t=Table([[P(x) for x in headers]]+[[P(x) for x in row] for row in rows],colWidths=widths,repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#F7E6EC')),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#C6CFD2')),
            ('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]))
        story.extend([t,Spacer(1,14)])
    story.append(Paragraph('Bilan des véhicules',styles['Heading2']))
    table(['Indicateur','Valeur'],[['Véhicules suivis',r['total']],['File initiale',r['initial']],['Nouvelles arrivées',r['arrivals']],['Admis dans la section',r['admitted']],
        ['Sortis avant la fin',r['exited']],['Encore en traversée',r['crossing']],['Encore en attente',r['remaining']],['File totale maximale',r['max_queue']],
        ['Attente moyenne observée (s)',fmt(r['mean_wait'],1)],['Attente moyenne des admis (s)',fmt(r['served_wait'],1)]],[310,190])
    story.append(Paragraph('Réglages et capacité du modèle',styles['Heading2']))
    table(['Paramètre','Valeur'],[['Débits A / B (véh/h)',f"{c['rate_a']} / {c['rate_b']}"],['Verts A / B (s)',f"{c['green_a']} / {c['green_b']}"],
        ['Dégagement par transition (s)',c['clearance']],['Espacement des entrées (s)',c['headway']],['Longueur (m) / vitesse (km/h)',f"{c['length']} / {c['speed']}"],
        ['Cycle (s)',cycle_time(c)],['Capacités à file saturée A / B (véh/h)',f"{capacity(c,'A'):.1f} / {capacity(c,'B'):.1f}"],['Graine',c['seed']]],[310,190])
    story.append(Paragraph('Hypothèses et limites',styles['Heading2']))
    for line in [
        'Conservation : file initiale + nouvelles arrivées = encore en file + en traversée + sortis.',
        'Le scénario commence par le vert A. Les files sont FIFO. Un créneau est disponible toutes les h secondes de vert ; le premier est à h secondes. Les créneaux vides sont perdus.',
        'Capacité par sens : partie entière(vert / h) x 3600 / cycle. Ce débit suppose une file toujours alimentée, pas une garantie de service.',
        'La traversée dure longueur / (vitesse / 3,6). Le dégagement vérifié couvre seulement ce temps idéal à vitesse constante.',
        'L’attente observée des véhicules non admis s’arrête à la fin de la simulation : leur attente finale est inconnue. La moyenne observée est une borne inférieure de la moyenne finale.',
        'Poisson suppose des arrivées indépendantes à débit constant. Les comptages fournis ne valident pas cette hypothèse. Pas de piétons, de pannes, de carrefour voisin ou de diversité des véhicules.',
        'Outil pédagogique uniquement. Aucun réglage ne doit être appliqué sur une route réelle sans étude et autorisation des personnes compétentes.']:
        story.extend([P(line),Spacer(1,8)])
    def footer(canvas,doc):
        canvas.setFont(font,8);canvas.drawString(46,25,'TraficLab | Rapport pédagogique');canvas.drawRightString(A4[0]-46,25,str(doc.page))
    SimpleDocTemplate(out,pagesize=A4,rightMargin=46,leftMargin=46,topMargin=40,bottomMargin=45).build(story,onFirstPage=footer,onLaterPages=footer)
    return out.getvalue()


def css():
    st.html('''<style>
:root{color-scheme:light} .stApp,[data-testid="stAppViewContainer"]{background:#f4f7f8;color:#242b30}
[data-testid="stMainBlockContainer"]{max-width:1510px;padding:4.5rem 2.1rem 2rem}
[data-testid="stHeader"]{background:#f4f7f8}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #d8e0e2}
.tl-logo{font:800 27px/1.2 Arial;color:#242b30;padding:14px 0 4px}.tl-logo span{color:#b92046}
.tl-sub{font-size:10px;letter-spacing:0;color:#52636b;margin-bottom:26px}
[data-testid="stSidebar"] .stButton button{background:#fff;border:0;border-left:3px solid transparent;border-radius:0;min-height:44px;justify-content:flex-start;color:#33444b}
[data-testid="stSidebar"] .stButton button[kind="primary"]{background:#f9e8ed!important;border-color:#b92046!important;color:#931a38!important}
[data-testid="stSidebar"] button p{color:inherit!important}
.tl-head{display:flex;justify-content:space-between;align-items:center;gap:18px;border-bottom:1px solid #cdd9dd;padding:0 0 20px;margin-bottom:24px}
.tl-head h1{font:800 34px/1.3 Arial;color:#242b30;margin:0;padding:0;letter-spacing:0}.tl-head h1 span{color:#b92046}
.tl-head p{font-size:12px;color:#4d626b;margin:9px 0 0}.tl-head small{font-size:11px;line-height:1.7;color:#485a62;max-width:230px;text-align:right}
.tl-status{display:flex;align-items:center;gap:10px;margin:6px 0 19px;color:#42565e;font-size:12px;overflow-wrap:anywhere}
.tl-status b{background:#dceff0;color:#006772;padding:6px 9px;border-radius:4px;font-size:11px}
.tl-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:16px 0 22px}
.tl-kpi{border-radius:6px;min-width:0;background:#fff;border-bottom:3px solid #007d8a;padding:16px;color:#242b30}
.tl-kpi:nth-child(2){background:#242b30;color:#fff;border-color:#63cbd2}.tl-kpi:nth-child(3){background:#f7e5ec;border-color:#b92046}.tl-kpi:nth-child(4){background:#e1efe7;border-color:#25734f}
.tl-kpi span,.tl-kpi small{display:block;font-size:12px;line-height:1.55}.tl-kpi strong{display:block;font-size:28px;line-height:1.25;margin:8px 0;overflow-wrap:anywhere}
h2{font-size:24px!important;letter-spacing:0!important;color:#242b30}h3{font-size:19px!important;letter-spacing:0!important;color:#242b30}
[data-testid="stMain"] [data-testid="stMarkdownContainer"],[data-testid="stMain"] [data-testid="stWidgetLabel"] p{color:#242b30}
[data-testid="stCaptionContainer"] p{color:#4f626a!important}
.stButton button,.stDownloadButton button,[data-testid="stFormSubmitButton"] button{border-radius:5px;min-height:42px;white-space:normal}
.stButton button p,.stDownloadButton button p,[data-testid="stFormSubmitButton"] button p{white-space:normal;overflow-wrap:anywhere}
[data-testid="stMain"] button[kind="primary"]{background:#b92046;border-color:#b92046;color:#fff}
[data-testid="stMain"] button[kind="primary"] p{color:#fff!important}
[data-testid="stForm"]{background:#fff;border-radius:6px;border-color:#cbd6d9}
.tl-foot{border-top:1px solid #cbd6d9;margin-top:30px;padding-top:16px;color:#53666e;font-size:11px;line-height:1.7}
@media(max-width:1050px){[data-testid="stMainBlockContainer"]{padding:4.4rem 1.2rem 2rem}.tl-kpi{padding:12px}.tl-kpi strong{font-size:24px}}
@media(max-width:740px){.tl-head{align-items:flex-start;flex-direction:column;gap:10px}.tl-head small{text-align:left;max-width:100%}.tl-head h1{font-size:30px}.tl-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.tl-kpi strong{font-size:23px}.tl-status{flex-wrap:wrap}[data-testid="stHorizontalBlock"]{flex-wrap:wrap}[data-testid="stHorizontalBlock"]>[data-testid="stColumn"]{width:100%!important;min-width:100%!important}}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>''')


def kpis(items):
    st.html('<div class="tl-kpis">'+''.join('<div class="tl-kpi"><span>'+escape(label)+'</span><strong>'+escape(str(value))+'</strong><small>'+escape(sub)+'</small></div>' for label,value,sub in items)+'</div>')


def chart(fig,key,alt,height=340):
    fig.update_layout(template='plotly_white',height=height,paper_bgcolor='#fff',plot_bgcolor='#fff',
        font=dict(family='Arial',color=INK,size=12),margin=dict(l=45,r=20,t=75,b=60),
        legend=dict(orientation='h',y=1.08,yanchor='bottom',x=0,font=dict(size=11)))
    fig.update_xaxes(automargin=True);fig.update_yaxes(automargin=True)
    st.plotly_chart(fig,width='stretch',theme=None,key=key,alt=alt,config={'displaylogo':False,'scrollZoom':False,'responsive':True})


def phase(c,t):
    local=t%cycle_time(c)
    if local<c['green_a']:return 'A'
    if c['green_a']+c['clearance']<=local<c['green_a']+c['clearance']+c['green_b']:return 'B'
    return 'Dégagement'


def scene_figure(c,result,t=0):
    f=result['vehicles'] if result else pd.DataFrame(columns=VEHICLE_COLUMNS)
    def snapshot(second):
        data=[];counts=[]
        for sense,color in [('A',CYAN),('B',RUBY)]:
            subset=f[f['sens']==sense]
            queued=subset[(subset['arrivée_s']<=second)&(subset['entrée_s'].isna()|(subset['entrée_s']>second))]
            counts.append(len(queued));n=min(10,len(queued))
            x=[24-i*2.2 for i in range(n)] if sense=='A' else [76+i*2.2 for i in range(n)]
            data.append(go.Scatter(x=x,y=[0]*n,mode='markers',marker=dict(symbol='square',size=13,color=color,line=dict(color='#fff',width=2)),hoverinfo='skip',showlegend=False))
        for sense,color in [('A',CYAN),('B',RUBY)]:
            moving=f[(f['sens']==sense)&f['entrée_s'].notna()&(f['entrée_s']<=second)&(f['sortie_théorique_s']>second)]
            fractions=(second-moving['entrée_s'])/travel_time(c)
            x=25+50*fractions if sense=='A' else 75-50*fractions
            data.append(go.Scatter(x=x,y=[0]*len(moving),mode='markers',marker=dict(symbol='square',size=13,color=color,line=dict(color='#fff',width=2)),hoverinfo='skip',showlegend=False))
        active=phase(c,second)
        data.append(go.Scatter(x=[24,76],y=[3.2,3.2],mode='markers',marker=dict(size=19,color=[GREEN if active=='A' else RUBY,GREEN if active=='B' else RUBY],line=dict(width=3,color=INK)),hoverinfo='skip',showlegend=False))
        data.append(go.Scatter(x=[11,89],y=[-2.2,-2.2],mode='text',text=[f'A : {counts[0]}',f'B : {counts[1]}'],textfont=dict(color=INK,size=13),hoverinfo='skip',showlegend=False))
        annotations=[dict(x=.5,y=1.03,xref='paper',yref='paper',text=f'{int(second//60):02d}:{int(second%60):02d} · '+('Vert '+active if active in ('A','B') else 'Dégagement'),showarrow=False,font=dict(size=13,color=INK)),
            dict(x=50,y=2.1,text='TRAVAUX',showarrow=False,font=dict(size=11,color=INK))]
        return data,annotations
    traces,annotations=snapshot(t);fig=go.Figure(traces)
    fig.add_shape(type='rect',x0=0,x1=100,y0=-1,y1=1.2,fillcolor='#3b454c',line_width=0,layer='below')
    fig.add_shape(type='rect',x0=27,x1=73,y0=1.35,y1=2.8,fillcolor='#d4dde0',line_width=0,layer='below')
    for x in range(28,73,5):fig.add_shape(type='line',x0=x,x1=x+2,y0=1.4,y1=2.7,line=dict(color='#82959d',width=3),layer='below')
    for x in [25,75]:fig.add_shape(type='line',x0=x,x1=x,y0=-1,y1=1.2,line=dict(color='#fff',width=3),layer='below')
    fig.add_shape(type='line',x0=0,x1=100,y0=-1.1,y1=-1.1,line=dict(color='#a9bcc3',width=2),layer='below')
    fig.update_layout(height=280,paper_bgcolor='#fff',plot_bgcolor='#fff',margin=dict(l=12,r=12,t=40,b=25),annotations=annotations,
        xaxis=dict(range=[-6,106],visible=False,fixedrange=True),yaxis=dict(range=[-3.4,4.3],visible=False,fixedrange=True),showlegend=False)
    return fig


def stop_replay():
    replay=st.session_state.get('road_replay')
    if replay:replay['playing']=False


def road_player(c,result,t):
    key=st.session_state.workspace+'_'+fingerprint(c,result['events'])+'_'+str(t)
    if st.session_state.get('road_replay',{}).get('key')!=key:
        st.session_state.road_replay=dict(key=key,playing=False,position=float(t),end=float(min(c['minutes']*60,t+cycle_time(c))))
    replay=st.session_state.road_replay

    @st.fragment(run_every=.25 if replay['playing'] else None)
    def panel():
        if replay['playing']:
            replay['position']=min(replay['end'],replay['from_time']+(time.monotonic()-replay['wall_start'])*cycle_time(c)/10)
            if replay['position']>=replay['end']:
                replay['playing']=False;st.rerun()
        with st.container(horizontal=True):
            if st.button('Animer',icon=':material/play_arrow:',disabled=replay['playing'] or replay['position']>=replay['end'],key='road_play'):
                replay.update(playing=True,from_time=replay['position'],wall_start=time.monotonic());st.rerun()
            if st.button('Pause',icon=':material/pause:',disabled=not replay['playing'],key='road_pause'):
                replay['playing']=False;st.rerun()
            if st.button('Recommencer',icon=':material/replay:',key='road_restart'):
                replay.update(playing=False,position=float(t));st.rerun()
        st.plotly_chart(scene_figure(c,result,replay['position']),width='stretch',theme=None,key='road_scene',
            alt='Files et véhicules traversant la section à l’instant affiché',config={'displayModeBar':False,'responsive':True})
    panel()


def queue_chart(result,key):
    f=result['timeline'];fig=go.Figure()
    for sense,color in [('A',CYAN),('B',RUBY)]:
        fig.add_trace(go.Scatter(x=f.minute,y=f['file_'+sense],name='Sens '+sense,mode='lines',line=dict(color=color,width=2,shape='hv'),fill='tozeroy',opacity=.8))
    fig.update_layout(xaxis_title='Minutes',yaxis_title='Véhicules en file')
    chart(fig,key,'Files d’attente A et B à chaque événement de la simulation')


def new_study():return dict(config=copy.deepcopy(DEFAULT),events=[],result=None)


def example_study():
    s=new_study();s['config'].update(name='Alternat Campus · exemple fictif',rate_a=360.,rate_b=240.,initial_a=4,initial_b=2,seed=42)
    s['result']=cached_run(s['config'],None)
    return s


def load_example():
    stop_replay()
    st.session_state.studies['Exemple fictif']=example_study()
    st.session_state.workspace='Exemple fictif';st.session_state.page='Poste de circulation'
    st.session_state.flash='Exemple fictif chargé dans un espace séparé.'


def selected_events(study):return study['events'] if study['config']['mode']=='Comptage fourni' else None


def run_study(study):
    study['result']=cached_run(study['config'],selected_events(study))


def go_page(name):stop_replay();st.session_state.page=name;st.rerun()


def dashboard(study):
    c=study['config'];result=study['result']
    st.header('Poste de circulation')
    with st.container(horizontal=True):
        if st.button('Lancer la simulation',type='primary',icon=':material/play_arrow:'):
            run_study(study);st.rerun()
        if st.button('Régler le chantier',icon=':material/tune:'):go_page('Paramètres et comptages')
        st.button('Charger l’exemple fictif',icon=':material/science:',on_click=load_example)
    if result:
        r=result['summary']
        kpis([('Véhicules suivis',r['total'],f"{r['initial']} initiaux + {r['arrivals']} arrivées"),('Sortis de la section',r['exited'],f"{r['crossing']} encore en traversée"),
            ('Encore en attente',r['remaining'],'À la fin de l’observation'),('File totale maximale',r['max_queue'],'Deux sens cumulés, même instant')])
        sig=fingerprint(c,study['events']);peak=float(result['timeline'].loc[result['timeline'].file_totale.idxmax(),'seconde'])
        t=st.slider('Instant de départ (secondes)',0,c['minutes']*60,min(int(peak),c['minutes']*60-1),key='instant_'+st.session_state.workspace+'_'+sig)
    else:
        kpis([('Capacité du modèle A',fmt(capacity(c,'A')),'véh/h, file toujours alimentée'),('Capacité du modèle B',fmt(capacity(c,'B')),'véh/h, file toujours alimentée'),
            ('Cycle complet',str(cycle_time(c))+' s','A / dégagement / B / dégagement'),('Durée observée',str(c['minutes'])+' min','Simulation non lancée')]);t=0
    if result:road_player(c,result,t)
    else:st.plotly_chart(scene_figure(c,result,t),width='stretch',theme=None,key='road_preview',
        alt='Section routière en travaux avec deux approches et une voie partagée',config={'displayModeBar':False,'responsive':True})
    st.caption(f"Section : {fmt(c['length'])} m · Traversée idéale : {fmt(travel_time(c),1)} s · Dégagement : {c['clearance']} s par transition. Schéma non à l’échelle, au plus 10 véhicules dessinés par file ; compteurs exhaustifs.")
    if not result:return
    a,b=st.columns(2)
    with a:queue_chart(result,'files_dashboard')
    with b:
        f=result['vehicles'];fig=go.Figure()
        for sense,color in [('A',CYAN),('B',RUBY)]:
            vals=f.loc[(f.sens==sense)&f['entrée_s'].notna(),'attente_avant_entrée_s']
            fig.add_trace(go.Histogram(x=vals,name='Sens '+sense,marker_color=color,opacity=.7,nbinsx=14))
        fig.update_layout(barmode='overlay',xaxis_title='Attente des admis (s)',yaxis_title='Véhicules')
        chart(fig,'attentes_dashboard','Distribution des attentes avant entrée, véhicules admis uniquement')
    st.caption(f"Attente moyenne observée : {fmt(r['mean_wait'],1)} s. Moyenne des seuls véhicules admis : {fmt(r['served_wait'],1)} s. Pour les {r['remaining']} véhicules encore en file, l’attente finale reste inconnue.")
    if c['minutes']*60<3*cycle_time(c):st.warning('Horizon inférieur à trois cycles : le démarrage par le sens A pèse fortement sur le résultat.')
    if c['mode']!='Comptage fourni' and any(c['rate_'+s.lower()]>capacity(c,s) for s in ('A','B')):
        st.warning('Au moins un débit demandé dépasse la capacité à file saturée du modèle. La file peut croître avec la durée d’observation.')


def settings_page(study):
    st.header('Paramètres et comptages');c=study['config'];sig=fingerprint(c,study['events'])
    with st.form('settings_'+st.session_state.workspace+'_'+sig):
        name=st.text_input('Nom du scénario',c['name'],max_chars=80)
        a,b=st.columns(2);minutes=a.number_input('Durée d’observation (min)',1,60,c['minutes']);mode=b.selectbox('Mode des arrivées',MODES,index=MODES.index(c['mode']))
        a,b=st.columns(2);ra=a.number_input('Débit du sens A (véh/h)',0.,1500.,c['rate_a'],10.);rb=b.number_input('Débit du sens B (véh/h)',0.,1500.,c['rate_b'],10.)
        a,b=st.columns(2);ga=a.number_input('Vert A (s)',5,300,c['green_a']);gb=b.number_input('Vert B (s)',5,300,c['green_b'])
        a,b=st.columns(2);clearance=a.number_input('Dégagement à chaque transition (s)',1,600,c['clearance']);headway=b.number_input('Espacement entre entrées (s)',1.,15.,c['headway'],.5)
        a,b=st.columns(2);length=a.number_input('Longueur de la section (m)',10.,1000.,c['length'],10.);speed=b.number_input('Vitesse constante de traversée (km/h)',5.,60.,c['speed'],5.)
        a,b=st.columns(2);ia=a.number_input('File initiale A (véhicules)',0,200,c['initial_a']);ib=b.number_input('File initiale B (véhicules)',0,200,c['initial_b'])
        seed=st.number_input('Graine aléatoire',0,2147483647,c['seed'])
        run=st.form_submit_button('Appliquer les paramètres',icon=':material/save:',type='primary')
    if run:
        q=check_config(dict(name=name,minutes=minutes,mode=mode,rate_a=ra,rate_b=rb,green_a=ga,green_b=gb,clearance=clearance,headway=headway,length=length,speed=speed,initial_a=ia,initial_b=ib,seed=seed))
        arrivals=check_events(study['events'],q)
        study.update(config=q,events=arrivals,result=None);st.session_state.flash='Paramètres enregistrés. La simulation précédente a été retirée.';st.rerun()
    st.caption('Les débits servent aux arrivées simulées et aux scénarios probabilistes. En mode Comptage fourni, les arrivées individuelles ci-dessous les remplacent pour la simulation principale.')
    st.subheader('Arrivées individuelles')
    with st.form('arrivals_'+st.session_state.workspace+'_'+sig):
        table=pd.DataFrame(study['events'],columns=['arrivee_s','sens']).astype({'arrivee_s':'float64','sens':'str'})
        edited=st.data_editor(table,num_rows='dynamic',hide_index=True,width='stretch',key='events_'+st.session_state.workspace+'_'+sig,
            column_config={'arrivee_s':st.column_config.NumberColumn('Arrivée (s)',min_value=0,max_value=c['minutes']*60,required=True,format='%.2f'),
                'sens':st.column_config.SelectboxColumn('Sens',options=['A','B'],required=True)},alt='Une ligne par arrivée, seconde depuis le début et sens A ou B')
        save=st.form_submit_button('Utiliser ce comptage',icon=':material/save:')
    if save:
        events=check_events(edited.to_dict('records'),c)
        study.update(events=events,config={**c,'mode':'Comptage fourni'},result=None)
        st.session_state.flash='Comptage remplacé et mode Comptage fourni activé.';st.rerun()
    st.caption('Une ligne représente un véhicule. Deux lignes identiques représentent deux véhicules arrivés ensemble. Les files initiales ne doivent pas être ressaisies dans ce tableau.')
    template=pd.DataFrame({'arrivee_s':[0.,0.,5.,12.],'sens':['A','B','A','B']})
    st.download_button('Modèle de comptage CSV',csv_bytes(template),'traficlab-modele.csv','text/csv',icon=':material/download:')
    upload=st.file_uploader('Importer un comptage CSV ou Excel',type=['csv','xlsx'],max_upload_size=2,key='import_'+st.session_state.workspace)
    if upload:
        events=read_arrivals(upload.getvalue(),upload.name,c)
        st.dataframe(pd.DataFrame(events),hide_index=True,width='stretch',alt='Prévisualisation des arrivées à importer')
        if st.button('Remplacer par ce fichier',icon=':material/upload:'):
            study.update(events=events,config={**c,'mode':'Comptage fourni'},result=None);st.session_state.flash='Comptage importé, anciennes arrivées remplacées.';st.rerun()


def comparison_page(study):
    st.header('Comparer deux réglages');c=study['config'];sig=fingerprint(c,study['events'])
    with st.form('compare_'+st.session_state.workspace+'_'+sig):
        a,b=st.columns(2);ga=a.number_input('Vert A alternatif (s)',5,300,c['green_a']);gb=b.number_input('Vert B alternatif (s)',5,300,c['green_b'])
        a,b=st.columns(2);clearance=a.number_input('Dégagement alternatif (s)',1,600,c['clearance']);h=b.number_input('Espacement alternatif (s)',1.,15.,c['headway'],.5)
        run=st.form_submit_button('Comparer à arrivées identiques',type='primary',icon=':material/compare_arrows:')
    if run:
        alternative={**c,'green_a':ga,'green_b':gb,'clearance':clearance,'headway':h}
        arrivals=selected_events(study)
        if arrivals is None:arrivals=generate_events(c)
        results=compare(c,arrivals,alternative)
        st.session_state['comparison_'+st.session_state.workspace]=(sig,results)
    saved=st.session_state.get('comparison_'+st.session_state.workspace)
    if not saved or saved[0]!=sig:return
    results=saved[1];rows=[];fig=go.Figure()
    for label,result,color in zip(['Référence','Alternative'],results,[CYAN,RUBY]):
        r=result['summary'];q=result['config']
        rows.append({'Scénario':label,'Cycle (s)':cycle_time(q),'Capacité A (véh/h)':capacity(q,'A'),'Capacité B (véh/h)':capacity(q,'B'),
            'File maximale':r['max_queue'],'File finale':r['remaining'],'Sortis':r['exited'],'Attente observée (s)':r['mean_wait']})
        f=result['timeline'];fig.add_trace(go.Scatter(x=f.minute,y=f.file_totale,name=label,line=dict(color=color,shape='hv',width=2)))
    st.dataframe(pd.DataFrame(rows).round(2),hide_index=True,width='stretch',alt='Comparaison des capacités, files, sorties et attentes observées des deux réglages')
    fig.update_layout(xaxis_title='Minutes',yaxis_title='File totale (véhicules)');chart(fig,'comparison','Comparaison temporelle des deux files totales avec les mêmes arrivées')
    st.download_button('Exporter la comparaison CSV',csv_bytes(pd.DataFrame(rows)),'traficlab-comparaison.csv','text/csv',icon=':material/download:')
    st.caption('Mêmes arrivées individuelles, files initiales et horizon. Les scénarios démarrent par A. Une meilleure attente parmi les admis peut masquer une file finale plus grande : comparer plusieurs indicateurs.')


def probability_page(study):
    st.header('Probabilités et incertitudes');c=study['config'];sig=fingerprint(c,study['events'])
    st.subheader('Combien de véhicules peuvent arriver ?')
    with st.form('poisson_'+st.session_state.workspace+'_'+sig):
        a,b=st.columns(2);sense=a.selectbox('Sens étudié',['A','B']);interval=b.number_input('Intervalle étudié (min)',.1,60.,1.,.5)
        limit=st.number_input('Plus de combien de véhicules ?',0,5000,5)
        run=st.form_submit_button('Calculer la probabilité',icon=':material/calculate:')
    if run:
        if sense not in ('A','B'):raise ValueError('Sens invalide.')
        interval=number(interval,'Intervalle',.1,60);limit=number(limit,'Seuil',0,5000,True)
        mu=c['rate_'+sense.lower()]*interval/60;p=float(stats.poisson.sf(limit,mu))
        st.session_state['poisson_'+st.session_state.workspace]=(sig,mu,p,limit,sense)
    saved=st.session_state.get('poisson_'+st.session_state.workspace)
    if saved and saved[0]==sig:
        _,mu,p,limit,sense=saved
        st.metric(f'P(N > {limit}) · sens {sense}',fmt(100*p,2)+' %')
        hi=int(stats.poisson.ppf(.999,mu)) if mu else 5
        x=np.arange(max(0,int(mu-4*math.sqrt(mu+1))),hi+2)
        fig=go.Figure(go.Bar(x=x,y=stats.poisson.pmf(x,mu),marker_color=[RUBY if k>limit else CYAN for k in x]))
        fig.update_layout(xaxis_title='Véhicules arrivant',yaxis_title='Probabilité');chart(fig,'poisson_counts','Loi de Poisson du nombre d’arrivées, seuil de dépassement en rouge')
        st.caption(f'Moyenne théorique μ = débit × intervalle / 60 = {fmt(mu,2)} véhicules. Arrivées indépendantes, débit constant. Ce calcul porte sur les arrivées, pas sur la longueur de file.')
    st.subheader('Risque d’une longue file')
    with st.form('mc_'+st.session_state.workspace+'_'+sig):
        a,b=st.columns(2);n=a.select_slider('Nombre de répétitions',[20,50,100,200],value=100);seed=b.number_input('Graine Monte-Carlo',0,2147483647,123)
        threshold=st.number_input('Seuil de file totale (véhicules)',0,5000,15)
        sensitivity=st.checkbox('Inclure des débits à −20 % et +20 %',False)
        run=st.form_submit_button('Simuler des arrivées Poisson',icon=':material/play_arrow:',type='primary')
    if run:
        with st.spinner('Simulation des alternats…'):frame=monte_carlo(c,n,seed,threshold,sensitivity)
        st.session_state['mc_result_'+st.session_state.workspace]=(sig,frame,int(threshold))
    saved=st.session_state.get('mc_result_'+st.session_state.workspace)
    if not saved or saved[0]!=sig:return
    _,frame,threshold=saved;base=frame[frame.facteur==1.];r=risk_stats(base.file_max,threshold)
    kpis([('P(file max > seuil)',fmt(r['probability']*100,1)+' %',f'Seuil : {threshold} véhicules'),('Intervalle à 95 %',fmt(100*r['low'],1)+' à '+fmt(100*r['high'],1)+' %','Incertitude Monte-Carlo, Wilson'),
        ('Médiane des maxima',fmt(r['q50']),'Véhicules dans les deux files'),('Quantile 95 %',fmt(r['q95']),'Quantile empirique des maxima')])
    fig=go.Figure(go.Histogram(x=base.file_max,marker_color=CYAN,xbins=dict(size=1)))
    fig.add_vline(x=threshold,line_color=RUBY,line_dash='dash');fig.update_layout(xaxis_title='File totale maximale',yaxis_title='Simulations')
    chart(fig,'mc_distribution','Distribution des files maximales parmi les simulations Monte-Carlo')
    if len(frame.facteur.unique())>1:
        rows=[]
        for factor,group in frame.groupby('facteur'):
            estimate=risk_stats(group.file_max,threshold)
            rows.append({'Débit relatif (%)':round(100*factor),'Risque estimé (%)':100*estimate['probability'],'Quantile 95 % (véhicules)':estimate['q95']})
        st.dataframe(pd.DataFrame(rows).round(2),hide_index=True,width='stretch',alt='Sensibilité du risque et du quantile aux débits augmentés ou réduits de vingt pour cent')
    st.download_button('Exporter les simulations CSV',csv_bytes(frame),'traficlab-monte-carlo.csv','text/csv',icon=':material/download:')
    st.caption('Chaque répétition remplace les arrivées par un nouveau processus de Poisson ; géométrie, débits de référence et files initiales restent fixés. Les comptages importés ne sont pas rééchantillonnés. L’intervalle de Wilson ne couvre pas l’incertitude sur le modèle ou les débits.')


def report_page(study):
    st.header('Rapports et sauvegardes');result=study['result'];sig=fingerprint(study['config'],study['events'])
    if result:
        r=result['summary'];st.success(f"Bilan : {r['initial']} initiaux + {r['arrivals']} arrivées = {r['remaining']} en attente + {r['crossing']} en traversée + {r['exited']} sortis.")
        st.dataframe(result['vehicles'].round(2),hide_index=True,width='stretch',alt='Arrivées, entrées, sorties théoriques, attentes et état de chaque véhicule')
        st.download_button('Véhicules CSV',csv_bytes(result['vehicles']),'traficlab-vehicules.csv','text/csv',icon=':material/download:')
        st.download_button('Files CSV',csv_bytes(result['timeline']),'traficlab-files.csv','text/csv',icon=':material/download:')
        if st.button('Préparer le PDF et le classeur',icon=':material/description:',type='primary'):
            st.session_state['reports_'+st.session_state.workspace]=(sig,pdf_bytes(result,st.session_state.workspace),excel_bytes(result))
        reports=st.session_state.get('reports_'+st.session_state.workspace)
        if reports and reports[0]==sig:
            st.download_button('Rapport PDF',reports[1],'traficlab-rapport.pdf','application/pdf',icon=':material/download:')
            st.download_button('Classeur Excel',reports[2],'traficlab-resultats.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',icon=':material/download:')
    else:st.info('Aucun résultat simulé dans cet espace.')
    st.subheader('Sauvegarde du scénario')
    st.download_button('Sauvegarde JSON',backup_bytes(study),'traficlab-scenario.json','application/json',icon=':material/save:')
    upload=st.file_uploader('Restaurer un scénario JSON',type=['json'],max_upload_size=1,key='backup_'+st.session_state.workspace)
    if upload:
        restored=restore(upload.getvalue());st.text(restored['config']['name'])
        confirm=st.checkbox('Remplacer le scénario de cet espace',key='confirm_'+st.session_state.workspace)
        if st.button('Restaurer le scénario',disabled=not confirm,icon=':material/restore:'):
            st.session_state.studies[st.session_state.workspace]=restored
            for prefix in ['reports_','comparison_','mc_result_','poisson_']:st.session_state.pop(prefix+st.session_state.workspace,None)
            st.session_state.flash='Scénario restauré. Relancer la simulation pour produire les résultats.';st.rerun()
    st.caption('Données en mémoire de session, sans base en ligne. La sauvegarde JSON conserve les paramètres et comptages ; les fichiers CSV/Excel exportent les résultats, pas une sauvegarde réimportable complète.')


def teaching_page():
    st.header('Comprendre et tester')
    st.subheader('Deux files, une seule voie')
    st.markdown('Le sens A passe pendant son vert. Les deux sens sont ensuite bloqués pendant le dégagement. Le sens B passe à son tour, puis un nouveau dégagement précède le cycle suivant.')
    st.latex(r'C=g_A+g_B+2r,\qquad t_{tr}=\frac{L}{v/3.6},\qquad K_A=\left\lfloor\frac{g_A}{h}\right\rfloor\frac{3600}{C}')
    st.markdown('**C** : cycle (s), **g** : vert (s), **r** : dégagement par transition (s), **L** : longueur (m), **v** : vitesse (km/h), **h** : espacement des entrées (s), **K** : capacité à file toujours alimentée (véh/h).')
    st.subheader('Un test vérifiable à la calculatrice')
    st.markdown('Durée **1 min**, verts **10 s / 10 s**, dégagement **6 s**, espacement **5 s**, longueur **10 m**, vitesse **36 km/h**, files initiales **0 / 0**. En mode comptage, saisir trois véhicules A et deux B, tous à la seconde 0.')
    st.table(pd.DataFrame({'Véhicule':['A1','A2','A3','B1','B2'],'Entrée (s)':[5,10,37,21,26],'Sortie (s)':[6,11,38,22,27],'Attente (s)':[5,10,37,21,26]}),alt='Test exact de cinq véhicules avec entrées et sorties calculables manuellement')
    st.markdown('Résultats attendus : **5 sortis**, **0 en attente**, **file maximale 5**, attente moyenne **19,8 s**. Cycle **32 s** ; capacités théoriques **225 véh/h par sens**. Le troisième A attend le cycle suivant, même si aucun véhicule B ne passe pendant le vert A.')
    st.subheader('Vérifications complémentaires')
    st.markdown('**Aucune arrivée :** débits 0 / 0, files initiales 0 / 0 → aucun véhicule et une file maximale nulle.\n\n**Poisson :** débit A 60 véh/h, intervalle 1 min, seuil 0 → P(N > 0) = **63,21 %**.\n\n**Comparaison :** mêmes réglages des deux côtés → résultats identiques.\n\n**Monte-Carlo :** débits 0 / 0 et files initiales 0 / 0 → file maximale 0 à chaque répétition, risque de dépasser 0 égal à 0 dans les essais.\n\n**Dégagement :** longueur 120 m, vitesse 30 km/h → traversée 14,4 s ; dégagement 10 s refusé par le modèle.')
    st.subheader('Ce que les résultats ne disent pas')
    st.markdown('L’horizon démarre à vide ou avec la file initiale saisie, toujours par le sens A. L’attente finale des véhicules encore bloqués est inconnue. Un débit inférieur à la capacité ne supprime pas les files dues au rouge. Les arrivées groupées par un feu voisin, les poids lourds, piétons et incidents ne sont pas modélisés.')
    st.warning('Application pédagogique. La vérification longueur / vitesse n’est ni un calcul réglementaire du dégagement ni une autorisation de régler une signalisation réelle.')
    st.markdown('[SimPy : files FIFO](https://simpy.readthedocs.io/en/stable/api_reference/simpy.resources.html) · [SciPy : loi de Poisson](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.poisson.html) · [FHWA : circulation alternée, document américain](https://mutcd.fhwa.dot.gov/htm/2009/part4/part4h.htm). Ces références ne constituent pas une norme camerounaise.')


def main():
    st.set_page_config(page_title='TraficLab · Alternat routier',page_icon=':material/traffic:',layout='wide')
    st.session_state.setdefault('studies',{'Mes données':new_study(),'Exemple fictif':new_study()})
    st.session_state.setdefault('workspace','Mes données');st.session_state.setdefault('page','Poste de circulation')
    css()
    pages=[('Poste de circulation','traffic'),('Paramètres et comptages','tune'),('Comparer','compare_arrows'),('Probabilités','query_stats'),('Rapports','description'),('Comprendre et tester','school')]
    with st.sidebar:
        st.html('<div class="tl-logo">Trafic<span>Lab</span></div><div class="tl-sub">CIRCULATION / TRAVAUX ROUTIERS</div>')
        for name,icon in pages:
            if st.button(name,key='nav_'+icon,icon=f':material/{icon}:',width='stretch',wrap=True,type='primary' if st.session_state.page==name else 'secondary'):go_page(name)
        st.space(20)
        st.segmented_control('Espace de travail',['Mes données','Exemple fictif'],key='workspace',required=True,wrap=True,on_change=stop_replay)
        st.button('Charger l’exemple',on_click=load_example,icon=':material/science:',width='stretch',wrap=True)
        st.caption('v'+VERSION+' · usage pédagogique')
    study=st.session_state.studies[st.session_state.workspace];c=study['config']
    st.html('<header class="tl-head"><div><h1>Trafic<span>Lab</span></h1><p>CIRCULATION ALTERNÉE · GÉNIE CIVIL</p></div><small>'+escape(c['name'])+'<br>'+str(c['minutes'])+' min · '+fmt(c['length'])+' m · '+escape(c['mode'])+'</small></header>')
    st.html('<div class="tl-status"><b>'+escape(st.session_state.workspace)+'</b><span>'+('Arrivées fournies, non vérifiées' if c['mode']=='Comptage fourni' else 'Arrivées simulées · '+escape(c['mode']))+'</span></div>')
    if st.session_state.get('flash'):st.success(st.session_state.pop('flash'))
    try:
        page=st.session_state.page
        if page=='Poste de circulation':dashboard(study)
        elif page=='Paramètres et comptages':settings_page(study)
        elif page=='Comparer':comparison_page(study)
        elif page=='Probabilités':probability_page(study)
        elif page=='Rapports':report_page(study)
        else:teaching_page()
    except (ValueError,TypeError,KeyError,OverflowError) as exc:st.error(str(exc))
    st.html('<footer class="tl-foot">TraficLab · Files FIFO, arrivées et traversées distinctes · Modèle pédagogique, sans validation réglementaire</footer>')


if __name__=='__main__':main()
