import streamlit as st
import pandas as pd, numpy as np, subprocess, math, re, os, glob
from pathlib import Path
import plotly.graph_objects as go

st.set_page_config(layout="wide")

# ═══════════════════ Hilfs-Funktionen ═══════════════════════════════════════
def draw_target_plotly(disc: str) -> go.Figure:
    specs_map = {
        '10m Air Rifle': [(45.5,'white'),(40.5,'white'),(35.5,'white'),(30.5,'black'),
                          (25.5,'black'),(20.5,'black'),(15.5,'black'),(10.5,'black'),
                          (5.5,'black'),(0.5,'white')],
        '50m Rifle': [(154.4,'white'),(138.4,'white'),(122.4,'white'),(106.4,'white'),
                      (90.4,'white'),(74.4,'black'),(58.4,'black'),(42.4,'black'),
                      (26.4,'black'),(10.4,'black')],
        '300m Rifle': [(1000,'white'),(900,'white'),(800,'white'),(700,'white'),
                       (600,'black'),(500,'black'),(400,'black'),(300,'black'),
                       (200,'black'),(100,'black')],
        '10m Air Pistol': [(155.5,'white'),(139.5,'white'),(123.5,'white'),(107.5,'white'),
                           (91.5,'white'),(75.5,'white'),(59.5,'black'),(43.5,'black'),
                           (27.5,'black'),(11.5,'white')],
        '25m Rapid Fire Pistol': [(500,'black'),(420,'black'),(340,'black'),(260,'black'),
                                  (180,'black'),(100,'black')],
        '25m Precision Pistol': [(500,'white'),(450,'white'),(400,'white'),(350,'white'),
                                 (300,'white'),(250,'black'),(200,'black'),(150,'black'),
                                 (100,'black'),(50,'black')]
    }
    specs = specs_map.get(disc)
    fig = go.Figure()
    if specs is None:
        fig.add_annotation(text="Unknown discipline",x=0.5,y=0.5,showarrow=False)
        fig.update_layout(xaxis_visible=False,yaxis_visible=False)
        return fig
    for dia,col in specs:
        r = dia/2; edge='black' if col=='white' else 'white'
        fig.add_shape(type='circle',x0=-r,y0=-r,x1=r,y1=r,
                      fillcolor=col,line_color=edge,layer='below')
    fig.update_xaxes(autorange=True,scaleanchor='y',scaleratio=1,
                     showgrid=False,zeroline=False,title_text="mm")
    fig.update_yaxes(autorange=True,showgrid=False,zeroline=False,title_text="mm")
    fig.update_layout(width=600,height=600,dragmode='zoom',
                      margin=dict(l=20,r=20,t=20,b=20),showlegend=False)
    return fig

def safe_rm(p): 
    try: os.remove(p)
    except FileNotFoundError: pass

# ═══════════════════ Datei laden ════════════════════════════════════════════
@st.cache_data(show_spinner="Exportiere SCATT …")
def load_content(buf: bytes, name: str) -> list[str]:
    if name.lower().endswith(".txt"):
        try:  return buf.decode("utf-8").splitlines()
        except UnicodeDecodeError: return buf.decode("latin-1").splitlines()

    pipe = Path(__file__).with_name("export_pipeline"); pipe.mkdir(exist_ok=True)
    f_scatt = pipe/name; f_scatt.write_bytes(buf)
    res = subprocess.run(["export.bat", f_scatt.name],
                         cwd=pipe,shell=True,text=True,capture_output=True)
    if res.returncode:
        safe_rm(f_scatt)
        raise RuntimeError(f"Export-Fehler:\n{res.stderr}")
    txt = next(pipe.glob(f_scatt.stem+"*.txt"), None)
    if not txt:
        safe_rm(f_scatt); raise RuntimeError("Kein TXT erstellt.")
    lines = txt.read_text(encoding="utf-8",errors="ignore").splitlines()
    safe_rm(f_scatt); safe_rm(txt)
    return lines

up = st.file_uploader("SCATT-Datei / TXT", type=["scatt","txt"])
if not up: st.stop()
try: lines = load_content(up.getvalue(), up.name)
except Exception as e: st.error(str(e)); st.stop()
if not lines: st.error("Datei leer."); st.stop()

# ═══════════════════ Disziplin erkennen ════════════════════════════════════
disc_list = ['10m Air Rifle','50m Rifle','300m Rifle',
             '10m Air Pistol','25m Rapid Fire Pistol','25m Precision Pistol']
m = re.match(r'^([^\(]+)', lines[0]); discipline = m.group(1).strip() \
          if m and m.group(1).strip() in disc_list else disc_list[0]
st.sidebar.success(f"Disziplin: **{discipline}**")
st.text_area("Erste Zeile", lines[0], height=70)

# ═══════════════════ Shots parsen ══════════════════════════════════════════
def parse(ls):
    shots, cur = [], []
    for l in (ln.strip() for ln in ls):
        if l.startswith("Shot #"): 
            if cur: shots.append(pd.DataFrame(cur,columns=["t","x","y"])); cur=[]
        elif l:
            p=l.split()
            if len(p)>=3:
                try:
                    t=float(p[0])
                    x=float(p[1].split('=')[1] if '=' in p[1] else p[1])
                    y=float(p[2].split('=')[1] if '=' in p[2] else p[2])
                    cur.append([t,x,y])
                except: pass
    if cur: shots.append(pd.DataFrame(cur,columns=["t","x","y"]))
    return shots

@st.cache_data(show_spinner=False)
def get_shots(ls): return parse(ls)

shots = get_shots(lines)
if not shots: st.warning("Keine Schüsse."); st.stop()

# ═══════════════════ Konstanten & Score ════════════════════════════════════
PROJECTILE_DIAM = {
    '10m Air Rifle':4.5,'50m Rifle':5.6,'25m Rapid Fire Pistol':5.6,
    '25m Precision Pistol':5.6,'300m Rifle':7.62,'10m Air Pistol':4.5}
TEN_RING_DIAM = {
    '10m Air Rifle':0.5,'50m Rifle':10.4,'300m Rifle':100.0,
    '10m Air Pistol':11.5,'25m Rapid Fire Pistol':100.0,'25m Precision Pistol':50.0}

def issf_score(d, disc):
    if not math.isfinite(d):       # <-- NaN abfangen
        return float("nan")
    r10 = TEN_RING_DIAM[disc]/2; rproj = PROJECTILE_DIAM[disc]/2
    pen = (r10+rproj)/10
    raw = 11 - 0.1*(math.floor(d*100)/100)/pen
    return max(0,min(10.9,raw))

t0 = np.arange(0.5,-15.001,-0.01); idx0=np.argmin(np.abs(t0))

# ═══════════════════ Bias (Fadenkreuz-Verschiebung) ════════════════════════
@st.cache_data(show_spinner=False)
def bias(sh):
    arr=[]
    for s in sh:
        ts=s.sort_values('t')
        xi=np.interp(t0,ts['t'],ts['x'])
        yi=np.interp(t0,ts['t'],-ts['y'])
        arr.append((xi[idx0],yi[idx0]))
    return np.mean(arr,axis=0)
xbias,ybias=bias(shots)

# ═══════════════════ Sidebar-Einstellungen ═════════════════════════════════
min_t=min(s['t'].min() for s in shots)
start_hold = st.sidebar.slider("Start Hold-Phase (s)", 
                               float(round(max(-15,min_t),2)),
                               -0.25, value=-0.5, step=0.01)
mask_hold = (t0>=start_hold)&(t0<=-0.2)

st.sidebar.markdown("### Phasen anzeigen")
show_phase={k:st.sidebar.checkbox(k.title(),True) for k in
            ['approach','hold','release','recoil']}
show_avg = st.sidebar.checkbox("Ø Aiming-Punkt",True)
show_virtual = st.sidebar.checkbox("Virtueller Schuss",True)

# ═══════════════════ Metriken berechnen ════════════════════════════════════
@st.cache_data(show_spinner="Berechne Metriken…")
def metrics(sh, st_hold):
    mask=(t0>=st_hold)&(t0<=-0.2)
    xi_all,yi_all=[],[]
    for s in sh:
        ts=s.sort_values('t')
        xi_all.append(np.interp(t0,ts['t'],ts['x'])-xbias)
        yi_all.append(np.interp(t0,ts['t'],-ts['y'])-ybias)
    xi_all,yi_all=np.array(xi_all),np.array(yi_all)

    finite=np.isfinite(xi_all[:,mask])&np.isfinite(yi_all[:,mask])
    xm=[xi_all[i,mask][finite[i]].mean() if finite[i].any() else np.nan
        for i in range(len(sh))]
    ym=[yi_all[i,mask][finite[i]].mean() if finite[i].any() else np.nan
        for i in range(len(sh))]

    rows=[]
    for i in range(len(sh)):
        aiming=math.hypot(xm[i],ym[i])
        x0,y0=xi_all[i,idx0],yi_all[i,idx0]
        trigger=math.hypot(xm[i]-x0,ym[i]-y0)
        mx,my=xi_all[i,mask],yi_all[i,mask]
        stab=np.nan
        if np.isfinite(mx).sum()>1 and np.isfinite(my).sum()>1:
            cov=np.cov(mx,my)
            if np.all(np.isfinite(cov)):
                ev=np.linalg.eigvals(cov)
                stab=math.pi*5.991*math.sqrt(ev.max()*ev.min())
        dist=math.hypot(x0,y0)
        score=round(issf_score(dist,discipline),1)
        rows.append(dict(Shot=f"Shot {i+1}",Aiming_Error=aiming,
            Trigger_Error=trigger,Stability=stab,
            Center_Distance=dist,Score=score))
    df=pd.DataFrame(rows)
    # Serien & Overall
    series=[]; n=len(df)
    for s in range((n+9)//10):
        m=df.iloc[s*10:(s+1)*10].mean(numeric_only=True)
        m['Shot']=f"Series {s+1}"; series.append(m)
    overall=df.mean(numeric_only=True); overall['Shot']="Overall"
    return (pd.concat([df]+series+[overall],ignore_index=True)
              .set_index("Shot"))

all_metrics = metrics(shots, start_hold)

# ═══════════════════ Checkbox-Filter Tabelle ═══════════════════════════════
st.sidebar.markdown("### Tabelle filtern")
valid_labels=[str(x) for x in all_metrics.index if pd.notna(x)]
chosen=[]
for i,l in enumerate(valid_labels):
    if st.sidebar.checkbox(l,False,key=f"f_{i}_{l}"): chosen.append(l)
display_rows=chosen or valid_labels

# ═══════════════════ CSV-Download ══════════════════════════════════════════
st.sidebar.download_button("Metriken CSV",
    all_metrics.loc[display_rows].to_csv().encode("utf-8"),
    "scatt_metrics.csv","text/csv")

# ═══════════════════ Tabelle anzeigen ══════════════════════════════════════
st.subheader("Metriken")
st.dataframe(all_metrics.loc[display_rows],use_container_width=True)

# ═══════════════════ Auswahl-Umschalter (statt Tabs) ═══════════════════════
tab_choice = st.radio("Ansicht wählen",
                      ("🎯 Ziel","📈 Geschwindigkeit","📏 Ringwert"),
                      key="main_tabs", horizontal=True)

# ═══════════════════ Gemeinsame Helfer für Speed & Ring  ═══════════════════
@st.cache_data(show_spinner=False)
def vel_dist_arrays(sh, xb, yb):
    vel, dist = [], []
    for s in sh:
        t=s.sort_values('t')['t'].to_numpy()
        x=s.sort_values('t')['x'].to_numpy()-xb
        y=-s.sort_values('t')['y'].to_numpy()-yb
        xi=np.interp(t0,t,x,left=np.nan,right=np.nan)
        yi=np.interp(t0,t,y,left=np.nan,right=np.nan)
        vel.append(np.sqrt(np.gradient(xi,t0)**2+np.gradient(yi,t0)**2))
        dist.append(np.sqrt(xi**2+yi**2))
    return np.array(vel),np.array(dist)

vel_arr, dist_arr = vel_dist_arrays(shots,xbias,ybias)

def label_to_indices(labels):
    idx=[]
    for lb in labels:
        if lb.startswith("Shot"): idx.append(int(lb.split()[1])-1)
        elif lb.startswith("Series"):
            s=int(lb.split()[1])-1
            idx.extend(range(s*10,min((s+1)*10,len(shots))))
        else: idx.extend(range(len(shots)))
    return sorted(set(idx))

sel_idx = label_to_indices(display_rows)

# ═══════════════════ TAB: ZIEL ═════════════════════════════════════════════
if tab_choice == "🎯 Ziel":
    fig = draw_target_plotly(discipline)
    phase_col={'approach':'green','hold':'yellow','release':'blue','recoil':'red'}
    for idx in sel_idx:
        xi=np.interp(t0,shots[idx]['t'],shots[idx]['x'])-xbias
        yi=np.interp(t0,shots[idx]['t'],-shots[idx]['y'])-ybias
        if show_phase['approach']:
            m=t0<start_hold
            fig.add_trace(go.Scatter(x=xi[m],y=yi[m],mode='lines',
                line_width=2,line_color=phase_col['approach'],showlegend=False))
        if show_phase['hold']:
            m=(t0>=start_hold)&(t0<-0.2)
            fig.add_trace(go.Scatter(x=xi[m],y=yi[m],mode='lines',
                line_width=2,line_color=phase_col['hold'],showlegend=False))
        if show_phase['release']:
            m=(t0>=-0.2)&(t0<0)
            fig.add_trace(go.Scatter(x=xi[m],y=yi[m],mode='lines',
                line_width=2,line_color=phase_col['release'],showlegend=False))
        if show_phase['recoil']:
            m=(t0>=0)&(t0<=0.5)
            fig.add_trace(go.Scatter(x=xi[m],y=yi[m],mode='lines',
                line_width=2,line_color=phase_col['recoil'],showlegend=False))
        if show_avg:
            xm,ym=np.nanmean(xi[mask_hold]),np.nanmean(yi[mask_hold])
            fig.add_trace(go.Scatter(x=[xm],y=[ym],mode='markers',
                marker=dict(symbol='x',size=16,color='yellow'),showlegend=False))
        if show_virtual:
            x0_,y0_=xi[idx0],yi[idx0]; d=PROJECTILE_DIAM[discipline]
            fig.add_shape(type='circle',x0=x0_-d/2,y0=y0_-d/2,
                          x1=x0_+d/2,y1=y0_+d/2,
                          fillcolor='rgba(255,255,255,0.45)',
                          line_color='white',layer='above')
    st.plotly_chart(fig,use_container_width=True,
        config={'scrollZoom':True,'displaylogo':False})

# ═══════════════════ TAB: GESCHWINDIGKEIT ══════════════════════════════════
elif tab_choice == "📈 Geschwindigkeit":
    st.subheader("Geschwindigkeit")
    c1,c2=st.columns(2)
    with c1:
        x_min,x_max=st.slider("Zeit-Bereich (s)",
                              float(t0.min()),float(t0.max()),
                              (float(t0.min()),0.0),0.1)
    with c2:
        y_min,y_max=st.slider("v-Achse (mm/s)",0.0,1500.0,
                              (0.0,200.0),10.0)
    if not sel_idx: st.info("Keine Schüsse ausgewählt.")
    else:
        v_sel=vel_arr[sel_idx]
        # Einzel-Schüsse
        fig1=go.Figure()
        for i,idx in enumerate(sel_idx):
            fig1.add_trace(go.Scatter(x=t0,y=v_sel[i],mode='lines',
                name=f"Shot {idx+1}",
                hovertemplate="t=%{x:.2f}s<br>v=%{y:.1f} mm/s<extra></extra>"))
        fig1.add_vline(x=0,line_dash='dash',line_color='grey')
        fig1.update_layout(height=300,margin=dict(l=10,r=10,t=30,b=30),
            xaxis=dict(range=[x_min,x_max],title="Zeit (s)"),
            yaxis=dict(range=[y_min,y_max],title="Geschwindigkeit (mm/s)"))
        st.plotly_chart(fig1,use_container_width=True)
        # Mittelwert ± σ
        m=np.nanmean(v_sel,axis=0); s=np.nanstd(v_sel,axis=0)
        ok=np.isfinite(m); t_ok,m_ok,s_ok=t0[ok],m[ok],s[ok]
        fig2=go.Figure()
        fig2.add_trace(go.Scatter(x=t_ok,y=m_ok,mode='lines',
            name="Mittelwert",line=dict(color='black'),connectgaps=True))
        fig2.add_trace(go.Scatter(
            x=np.concatenate([t_ok,t_ok[::-1]]),
            y=np.concatenate([m_ok+s_ok,(m_ok-s_ok)[::-1]]),
            fill='toself',fillcolor='rgba(0,0,255,0.15)',
            line=dict(color='rgba(255,255,255,0)'),hoverinfo='skip',name='±1 σ'))
        fig2.add_vline(x=0,line_dash='dash',line_color='grey')
        fig2.update_layout(height=300,margin=dict(l=10,r=10,t=30,b=30),
            xaxis=dict(range=[x_min,x_max],title="Zeit (s)"),
            yaxis=dict(range=[y_min,y_max],title="Geschwindigkeit (mm/s)"))
        st.plotly_chart(fig2,use_container_width=True)

# ═══════════════════ TAB: RINGWERT ═════════════════════════════════════════
else:  # "📏 Ringwert"
    st.subheader("Ringwert")
    c1,c2=st.columns(2)
    with c1:
        x_min_d,x_max_d=st.slider("Zeit-Bereich (s)",
                                  float(t0.min()),float(t0.max()),
                                  (float(t0.min()),0.0),0.1,key="x_ring")
    with c2:
        y_min_d,y_max_d=st.slider("Ring-Achse",0.0,10.9,
                                  (0.0,10.9),0.1,key="y_ring")
    if not sel_idx: st.info("Keine Schüsse ausgewählt.")
    else:
        rings = np.array([[issf_score(d,discipline) for d in dist_arr[i]]
                          for i in sel_idx])
        # Einzel
        figR1=go.Figure()
        for i,idx in enumerate(sel_idx):
            figR1.add_trace(go.Scatter(x=t0,y=rings[i],mode='lines',
                name=f"Shot {idx+1}",
                hovertemplate="t=%{x:.2f}s<br>Ring=%{y:.1f}<extra></extra>"))
        figR1.add_vline(x=0,line_dash='dash',line_color='grey')
        figR1.update_layout(height=300,margin=dict(l=10,r=10,t=30,b=30),
            xaxis=dict(range=[x_min_d,x_max_d],title="Zeit (s)"),
            yaxis=dict(range=[y_min_d,y_max_d],title="Ringwert"))
        st.plotly_chart(figR1,use_container_width=True)
        # Mittel ± σ
        r_mean=np.nanmean(rings,axis=0); r_std=np.nanstd(rings,axis=0)
        ok=np.isfinite(r_mean); t_ok,r_ok,s_ok=t0[ok],r_mean[ok],r_std[ok]
        figR2=go.Figure()
        figR2.add_trace(go.Scatter(x=t_ok,y=r_ok,mode='lines',
            name="Mittelwert",line=dict(color='black'),connectgaps=True))
        figR2.add_trace(go.Scatter(
            x=np.concatenate([t_ok,t_ok[::-1]]),
            y=np.concatenate([r_ok+s_ok,(r_ok-s_ok)[::-1]]),
            fill='toself',fillcolor='rgba(255,0,0,0.15)',
            line=dict(color='rgba(255,255,255,0)'),hoverinfo='skip',name='±1 σ'))
        figR2.add_vline(x=0,line_dash='dash',line_color='grey')
        figR2.update_layout(height=300,margin=dict(l=10,r=10,t=30,b=30),
            xaxis=dict(range=[x_min_d,x_max_d],title="Zeit (s)"),
            yaxis=dict(range=[y_min_d,y_max_d],title="Ringwert"))
        st.plotly_chart(figR2,use_container_width=True)
