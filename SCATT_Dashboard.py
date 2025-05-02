import streamlit as st
import pandas as pd, numpy as np, subprocess, math, re, os, glob
from pathlib import Path
import plotly.graph_objects as go

st.set_page_config(layout="wide")
if "zoom_range" not in st.session_state:
    st.session_state.zoom_range = {"x": None, "y": None}


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

@st.cache_data(show_spinner="Lade Excel …")
def load_excel(file: bytes) -> list[str]:
    # Größtes Tabellenblatt automatisch wählen
    xls = pd.ExcelFile(file)
    sheet_lengths = {sheet: pd.read_excel(file, sheet_name=sheet).shape[0] for sheet in xls.sheet_names}
    main_sheet = max(sheet_lengths, key=sheet_lengths.get)
    df = pd.read_excel(file, sheet_name=main_sheet)

    # Zeitachse extrahieren
    if "Zeit (s)" not in df.columns:
        raise ValueError("Spalte 'Zeit (s)' fehlt in der Excel-Datei.")
    t_vals = df["Zeit (s)"].to_numpy()

    # Alle Spaltenpaare 'Sx x', 'Sx y' erkennen
    schussnummern = sorted(set(
        col.split()[0][1:] for col in df.columns if col.startswith("S") and ("x" in col or "y" in col)
    ), key=int)

    lines = []
    for nr in schussnummern:
        col_x = f"S{nr} x"
        col_y = f"S{nr} y"
        if col_x in df.columns and col_y in df.columns:
            x_vals = df[col_x].to_numpy()
            y_vals = df[col_y].to_numpy()
            lines.append(f"Shot #{nr}")
            for t, x, y in zip(t_vals, x_vals, y_vals):
                lines.append(f"{t:.3f} {x:.2f} {y:.2f}")
    return lines



up = st.file_uploader("SCATT- oder Excel-Datei", type=["scatt", "txt", "xlsx"])

if not up:
    st.stop()

try:
    if up.name.lower().endswith(".xlsx"):
        shots = load_excel(up)
        lines = ["Excel-Datei geladen"]  # Dummy-Zeile für Kompatibilität
    else:
        lines = load_content(up.getvalue(), up.name)
        shots = get_shots(lines)
except Exception as e:
    st.error(str(e))
    st.stop()


if not shots:
    st.error("Keine Schüsse erkannt.")
    st.stop()


# ═══════════════════ Disziplin erkennen ════════════════════════════════════
disc_list = ['10m Air Rifle','50m Rifle','300m Rifle',
             '10m Air Pistol','25m Rapid Fire Pistol','25m Precision Pistol']

if up.name.lower().endswith(".xlsx"):
    discipline = st.sidebar.selectbox("Disziplin wählen", disc_list)
else:
    m = re.match(r'^([^\(]+)', lines[0])
    discipline = m.group(1).strip() if m and m.group(1).strip() in disc_list else disc_list[0]
    st.text_area("Erste Zeile", lines[0], height=70)

st.sidebar.success(f"Disziplin: **{discipline}**")


# ═══════════════════ Auswahl-Umschalter (statt Tabs) ═══════════════════════
tab_choice = st.radio("Ansicht wählen",
    ("🎯 Ziel","📈 Geschwindigkeit","📏 Ringwert","📊 Gruppenvergleich"),
    key="main_tabs", horizontal=True)

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
show_timing_vecs = st.sidebar.checkbox("Timing-Vektoren anzeigen", True)


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

    rows = []
    for i in range(len(sh)):
        aiming = math.hypot(xm[i], ym[i])
        x0, y0 = xi_all[i, idx0], yi_all[i, idx0]
        trigger = math.hypot(xm[i] - x0, ym[i] - y0)

        mx, my = xi_all[i, mask], yi_all[i, mask]
        stab = np.nan
        if np.isfinite(mx).sum() > 1 and np.isfinite(my).sum() > 1:
            cov = np.cov(mx, my)
            if np.all(np.isfinite(cov)):
                ev = np.linalg.eigvals(cov)
                stab = math.pi * 5.991 * math.sqrt(ev.max() * ev.min())

        dist = math.hypot(x0, y0)
        score = round(issf_score(dist, discipline), 1)

        # Geschwindigkeit: v = sqrt(vx^2 + vy^2)
        vx = np.gradient(xi_all[i], t0)
        vy = np.gradient(yi_all[i], t0)
        speed = np.sqrt(vx**2 + vy**2)
        speed_hold = speed[mask]
        hold_speed = np.nanmean(speed_hold) if np.isfinite(speed_hold).any() else np.nan

        # Timing Winkel
        timing_idx_fixed = np.argmin(np.abs(t0 - (-0.05)))
        xtime = xi_all[i][timing_idx_fixed]
        ytime = yi_all[i][timing_idx_fixed]

        # Vektoren: Timingpunkt → Schuss & Zentrum
        v1 = np.array([x0 - xtime, y0 - ytime])        # zum Schuss
        v2 = np.array([-xtime, -ytime])               # zum Scheibenzentrum

        dot = np.dot(v1, v2)
        norm_product = np.linalg.norm(v1) * np.linalg.norm(v2)
        if norm_product == 0:
            timing_angle = np.nan
        else:
            angle_rad = np.arccos(np.clip(dot / norm_product, -1.0, 1.0))
            timing_angle = np.degrees(angle_rad)



        rows.append(dict(
            Shot=f"Shot {i+1}",
            Aiming_Error=aiming,
            Trigger_Error=trigger,
            Stability=stab,
            Hold_Speed=hold_speed,
            Timing_Angle=timing_angle,
            X_Timing=xtime,
            Y_Timing=ytime,
            Center_Distance=dist,
            Score=score
        ))


    df = pd.DataFrame(rows)
    # Serien & Overall
    series = []
    n = len(df)
    for s in range((n + 9) // 10):
        m = df.iloc[s * 10:(s + 1) * 10].mean(numeric_only=True)
        m['Shot'] = f"Series {s + 1}"
        series.append(m)
    overall = df.mean(numeric_only=True)
    overall['Shot'] = "Overall"
    return (pd.concat([df] + series + [overall], ignore_index=True)
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

if tab_choice == "🎯 Ziel":
    fig = draw_target_plotly(discipline)
    phase_col = {'approach': 'green', 'hold': 'yellow', 'release': 'blue', 'recoil': 'red'}

    for idx in sel_idx:
        xi = np.interp(t0, shots[idx]['t'], shots[idx]['x']) - xbias
        yi = np.interp(t0, shots[idx]['t'], -shots[idx]['y']) - ybias

        phase_order = ['approach', 'hold', 'release', 'recoil']
        phase_masks = {
            'approach': t0 < start_hold,
            'hold': (t0 >= start_hold) & (t0 < -0.2),
            'release': (t0 >= -0.2) & (t0 < 0),
            'recoil': (t0 >= 0) & (t0 <= 0.5),
        }

        phase_segments = {}
        for phase in phase_order:
            if not show_phase[phase]: continue
            mask = phase_masks[phase]
            x_seg = xi[mask]
            y_seg = yi[mask]
            if len(x_seg) == 0: continue
            fig.add_trace(go.Scatter(x=x_seg, y=y_seg, mode='lines',
                                     line=dict(width=2, color=phase_col[phase]),
                                     showlegend=False))
            phase_segments[phase] = (x_seg, y_seg)

        for i in range(len(phase_order) - 1):
            curr = phase_order[i]
            next_ = phase_order[i + 1]
            if curr not in phase_segments or next_ not in phase_segments:
                continue
            x1 = phase_segments[curr][0][0]
            y1 = phase_segments[curr][1][0]
            x2 = phase_segments[next_][0][-1]
            y2 = phase_segments[next_][1][-1]
            fig.add_trace(go.Scatter(
                x=[x1, x2], y=[y1, y2],
                mode='lines',
                line=dict(width=1, color=phase_col[curr]),
                showlegend=False
            ))

        if show_avg:
            xm, ym = np.nanmean(xi[mask_hold]), np.nanmean(yi[mask_hold])
            fig.add_trace(go.Scatter(x=[xm], y=[ym], mode='markers',
                                     marker=dict(symbol='x', size=16, color='yellow'),
                                     showlegend=False))

        if show_virtual:
            x0_, y0_ = xi[idx0], yi[idx0]
            d = PROJECTILE_DIAM[discipline]
            fig.add_shape(type='circle',
                          x0=x0_ - d/2, y0=y0_ - d/2,
                          x1=x0_ + d/2, y1=y0_ + d/2,
                          fillcolor='rgba(255,255,255,0.45)',
                          line_color='white', layer='above')

        # Timing-Vektoren
        if show_timing_vecs:
            shot_label = f"Shot {idx+1}"
            if shot_label in all_metrics.index:
                row = all_metrics.loc[shot_label]
                xtime, ytime = row.get("X_Timing", np.nan), row.get("Y_Timing", np.nan)
                if pd.notna(xtime) and pd.notna(ytime):
                    # Richtung Schuss
                    fig.add_trace(go.Scatter(
                        x=[xtime, x0_], y=[ytime, y0_],
                        mode='lines+markers',
                        line=dict(color='lime', dash='dot'),
                        marker=dict(size=6),
                        name=f"{shot_label} ➝ Schuss",
                        showlegend=False
                    ))
                    # Richtung Zentrum
                    fig.add_trace(go.Scatter(
                        x=[xtime, 0], y=[ytime, 0],
                        mode='lines+markers',
                        line=dict(color='orange', dash='dot'),
                        marker=dict(size=6),
                        name=f"{shot_label} ➝ Zentrum",
                        showlegend=False
                    ))

    fig.update_layout(
        xaxis=dict(autorange=True),
        yaxis=dict(autorange=True)
    )

    st.plotly_chart(fig, use_container_width=True,
                    config={'scrollZoom': True, 'displaylogo': False})




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
elif tab_choice == "📏 Ringwert":  # "📏 Ringwert"
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
elif tab_choice == "📊 Gruppenvergleich":
    st.subheader("Feature-Vergleich nach Gruppen")

    import scipy.stats as stats

    n_groups = st.slider("Anzahl Gruppen", 2, 6, 3)
    all_shot_labels = [f"Shot {i+1}" for i in range(len(shots))]

    # Gruppennamen links, Zuweisung rechts
    col_names, col_selects = st.columns(2)
    with col_names:
        group_labels = [
            st.text_input(f"Name für Gruppe {i+1}", value=f"Gruppe {i+1}", key=f"gname_{i}")
            for i in range(n_groups)
        ]

    assigned = set()
    group_assignments = {}

    with col_selects:
        for i, label in enumerate(group_labels):
            available = [s for s in all_shot_labels if s not in assigned]
            sel = st.multiselect(f"{label} – Schüsse wählen", options=available, key=f"group_{i}")
            group_assignments[label] = sel
            assigned.update(sel)

    # Feature-Auswahl
    feature_options = all_metrics.columns.tolist()
    feature_choice = st.selectbox("Feature wählen", feature_options)

    # Gruppenzugehörigkeit speichern
    shot_group_map = {}
    for grp_label, shots_ in group_assignments.items():
        for shot in shots_:
            shot_group_map[shot] = grp_label

    all_metrics["Gruppe"] = [shot_group_map.get(idx, "") for idx in all_metrics.index]

    # Plot-Daten vorbereiten
    plot_df = pd.DataFrame([
        {"Gruppe": grp, "Wert": all_metrics.loc[shot, feature_choice]}
        for grp, shots_ in group_assignments.items()
        for shot in shots_ if shot in all_metrics.index
    ])

    if plot_df.empty:
        st.info("Bitte mindestens einen Schuss einer Gruppe zuweisen.")
    else:
        # Funktion für Mittelwert + CI95
        def mean_ci95(vals):
            vals = np.array(vals, dtype=float)
            if len(vals) < 2:
                return np.nan, 0.0
            mean = np.mean(vals)
            se = stats.sem(vals, nan_policy='omit')
            ci = se * stats.t.ppf((1 + 0.95) / 2., len(vals)-1)
            return mean, ci

        summary_list = []
        for grp, values in plot_df.groupby("Gruppe")["Wert"]:
            mean, ci = mean_ci95(values)
            summary_list.append(dict(Gruppe=grp, Mittelwert=mean, CI=ci))

        summary_df = pd.DataFrame(summary_list)

        fig = go.Figure()
        for _, row in summary_df.iterrows():
            fig.add_trace(go.Bar(
                x=[row["Gruppe"]],
                y=[row["Mittelwert"]],
                error_y=dict(type="data", array=[row["CI"]], visible=True),
                name=row["Gruppe"]
            ))

        fig.update_layout(title=f"{feature_choice} – Gruppenvergleich mit 95%-Konfidenzintervallen",
                          xaxis_title="Gruppe", yaxis_title=feature_choice,
                          height=400, margin=dict(l=20, r=20, t=40, b=40))
        st.plotly_chart(fig, use_container_width=True)

# ═══════════════════ Metrikenanzeige mit Bewertungs-Option ═══════════════════════════
st.divider()
color_coded = st.toggle("🔁 Farbcodierte Darstellung", value=False)

ampel_cols = ["Aiming_Error", "Trigger_Error", "Stability", "Center_Distance", "Hold_Speed", "Timing_Angle"]

# Terzil-basierte Farbzuweisung
def ampelformat_colors(df, cols):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in cols:
        if col not in df.columns: continue
        values = df[col].dropna()
        if len(values) < 3: continue

        q1, q2 = values.quantile([1/3, 2/3])

        for idx in df.index:
            val = df.loc[idx, col]
            if pd.isna(val): continue
            if val <= q1:
                styles.loc[idx, col] = 'background-color: #d4edda'     # grün
            elif val <= q2:
                styles.loc[idx, col] = 'background-color: #fff3cd'     # gelb
            else:
                styles.loc[idx, col] = 'background-color: #f8d7da'     # rot
    return styles

df_disp = all_metrics.loc[display_rows].copy()
df_disp = df_disp.drop(columns=["X_Timing", "Y_Timing"], errors="ignore")
df_disp[ampel_cols] = df_disp[ampel_cols].applymap(lambda x: round(x, 2) if pd.notna(x) else "")

if color_coded:
    st.subheader("📋 Metriken-Tabelle (Ampelsystem)")
    styles = ampelformat_colors(df_disp, ampel_cols)
    st.dataframe(
        df_disp.style
            .apply(lambda _: styles, axis=None)
            .format(precision=2),
        use_container_width=True
    )
else:
    st.subheader("📋 Metriken-Tabelle (Rohwerte)")
    st.dataframe(df_disp, use_container_width=True)
