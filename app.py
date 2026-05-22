import streamlit as st
import pandas as pd
import random
from io import BytesIO
import datetime

st.set_page_config(
    page_title="Sistema MRP + VRP — NETO DURANGAR S.A.S.",
    page_icon="🚚", layout="wide"
)

st.markdown("""
<style>
.titulo { background:linear-gradient(90deg,#1F3864,#2E75B6);
          color:white;padding:18px;border-radius:10px;text-align:center;margin-bottom:16px; }
.ruta-ok   { background:#E2EFDA;border-left:4px solid #375623;padding:12px;border-radius:6px;margin:6px 0; }
.ruta-warn { background:#FFF2CC;border-left:4px solid #F4A300;padding:12px;border-radius:6px;margin:6px 0; }
.ruta-bad  { background:#FCE4D6;border-left:4px solid #C55A11;padding:12px;border-radius:6px;margin:6px 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="titulo">
  <h2 style="margin:0">🍽️ Sistema Integrado MRP + VRP — NETO DURANGAR S.A.S.</h2>
  <p style="margin:4px 0 0;opacity:.9">Generador de Menús · Requerimientos de Materiales · Rutas Clarke & Wright</p>
</div>""", unsafe_allow_html=True)

# ── FLOTA COMPLETA — 5 vehículos, 5 rutas ─────────────────────────────────────
# Cada vehículo DEBE usarse en exactamente una ruta
FLOTA = [
    {'placa':'TSN320', 'tipo':'Camión',    'costo_km':1387, 'cap_kg':2000, 'refrig':True},
    {'placa':'WEQ',    'tipo':'Camión',    'costo_km':675,  'cap_kg':8000, 'refrig':True},
    {'placa':'JTX761', 'tipo':'Camioneta', 'costo_km':221,  'cap_kg':700,  'refrig':False},
    {'placa':'SXD',    'tipo':'Camioneta', 'costo_km':204,  'cap_kg':700,  'refrig':False},
    {'placa':'SZO',    'tipo':'Camioneta', 'costo_km':253,  'cap_kg':700,  'refrig':False},
]

# ── DATOS VRP ──────────────────────────────────────────────────────────────────
CAMPOS = ['Base H&P','Carrizales','Yenac','Careto','Corcel','Arrendajo']

DIST_BODEGA = {
    'Base H&P':36, 'Carrizales':90, 'Yenac':119,
    'Careto':158,  'Corcel':178,    'Arrendajo':210
}
T_VIAJE = {
    'Base H&P':60, 'Carrizales':159, 'Yenac':190,
    'Careto':235,  'Corcel':210,     'Arrendajo':330
}
T_DESC = 65

DIST_INTER = {
    ('Base H&P','Carrizales'):85,  ('Base H&P','Yenac'):123,
    ('Base H&P','Careto'):174,     ('Base H&P','Corcel'):200,
    ('Base H&P','Arrendajo'):241,
    ('Carrizales','Yenac'):53,     ('Carrizales','Careto'):103,
    ('Carrizales','Corcel'):129,   ('Carrizales','Arrendajo'):171,
    ('Yenac','Careto'):66,         ('Yenac','Corcel'):92,
    ('Yenac','Arrendajo'):133,
    ('Careto','Corcel'):41,        ('Careto','Arrendajo'):83,
    ('Corcel','Arrendajo'):57,
}

def dij(a, b):
    if a == b: return 0
    return DIST_INTER.get((a,b), DIST_INTER.get((b,a), 0))

NECESITA_FRIO = {c: c not in ['Base H&P'] for c in CAMPOS}

# ── CARGAR RECETAS ─────────────────────────────────────────────────────────────
@st.cache_data
def cargar():
    des = pd.read_excel("Desayuno_Final_02.xlsx", header=0)
    des.columns = ['_','COMIDA','DESC','PREP','ING','g_p','n_p','td','tm']
    des = des.iloc[1:].dropna(subset=['PREP','ING'])
    des['g_p'] = pd.to_numeric(des['g_p'], errors='coerce')

    alm = pd.read_excel("Alm-Cena_02.xlsx", header=0)
    alm.columns = ['_','DESC','PREP','ING','g_p','n_p','td','tm']
    alm = alm.iloc[1:].dropna(subset=['PREP','ING'])
    alm['g_p'] = pd.to_numeric(alm['g_p'], errors='coerce')
    return des, alm

des, alm = cargar()

def pick(df, desc, excluir=None):
    ops = df[df['DESC']==desc]['PREP'].unique().tolist()
    if excluir: ops = [x for x in ops if x not in excluir] or ops
    return random.choice(ops) if ops else "N/D"

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Parámetros")
    personas = st.number_input("Personas/servicio", 100, 2000, 690, 10)
    semanas  = st.slider("Semanas del ciclo", 1, 4, 4)
    margen   = st.slider("Margen seguridad (%)", 0, 20, 10)
    ventana  = st.selectbox("Ventana horaria",
                             [300,540], index=1,
                             format_func=lambda x: f"{'5am–10am' if x==300 else '5am–2pm'} ({x} min)")
    st.markdown("---")
    st.markdown("### 🚚 Flota disponible")
    st.dataframe(pd.DataFrame(FLOTA)[['placa','tipo','costo_km','cap_kg','refrig']],
                 hide_index=True, use_container_width=True)
    st.caption("⚠️ Los 5 vehículos SE USAN TODOS — uno por ruta")

# ── TABS ───────────────────────────────────────────────────────────────────────
t1, t2, t3 = st.tabs([
    "🍽️ 1. Generador de Menús",
    "📦 2. MRP — Requerimientos",
    "🚚 3. VRP — Clarke & Wright"
])

dias = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — MENÚ
# ══════════════════════════════════════════════════════════════════════════════
def generar_menu(n_sem, n_per, rotar):
    menu = {}
    jd_h, ja_h = [], []
    for s in range(1, n_sem+1):
        menu[f'Semana {s}'] = {}
        for dia in dias:
            jd = pick(des,'Jugos', jd_h if rotar else None)
            if rotar: jd_h.append(jd); jd_h[:] = jd_h[-10:]
            ja = pick(alm,'Jugos', ja_h if rotar else None)
            if rotar: ja_h.append(ja); ja_h[:] = ja_h[-10:]
            p1 = (pick(alm,'Especial - 1x Semana (Pescados y Mariscos)')
                  if dia=='Viernes' else pick(alm,'Proteico 1 - Carne Roja'))
            menu[f'Semana {s}'][dia] = {
                'Desayuno': {
                    'Jugo':jd,'Bebida':pick(des,'Bebida Caliente '),
                    'Lácteo':pick(des,'Lácteos'),'Fruta':pick(des,'Fruta'),
                    'Cereal':pick(des,'Cereales'),'Queso':pick(des,'Queso'),
                    'Huevo':pick(des,'Huevos '),'Proteína 1':pick(des,'Proteina 1'),
                    'Proteína 2':pick(des,'Proteina 2'),'Caldo':pick(des,'Caldo'),
                    'Arroz':pick(des,'Arroz Cocido'),'Pan':pick(des,'Pan Varios'),
                },
                'Almuerzo': {
                    'Jugo':ja,'Fruta':pick(alm,'Frutas de Mano'),
                    'Sopa':pick(alm,'Sopa, Crema o Consomé'),
                    'Proteína 1':p1,'Proteína 2':pick(alm,'Proteico 2 - Carne Blanca'),
                    'Verduras':pick(alm,'Verduras Cocidas'),'Arroz':pick(alm,'Arroz'),
                    'Energético':pick(alm,'Energético'),'Ensalada':pick(alm,'Barra de Ensalada'),
                    'Leguminosa':pick(alm,'Leguminosa'),'Postre':pick(alm,'Postre'),
                },
                'Cena': {
                    'Jugo':pick(alm,'Jugos',[ja]),'Fruta':pick(alm,'Frutas de Mano'),
                    'Sopa':pick(alm,'Sopa, Crema o Consomé'),
                    'Proteína 1':pick(alm,'Proteico 1 - Carne Roja'),
                    'Proteína 2':pick(alm,'Proteico 2 - Carne Blanca'),
                    'Verduras':pick(alm,'Verduras Cocidas'),'Arroz':pick(alm,'Arroz'),
                    'Energético':pick(alm,'Energético'),'Ensalada':pick(alm,'Barra de Ensalada'),
                    'Leguminosa':pick(alm,'Leguminosa'),'Postre':pick(alm,'Postre'),
                }
            }
    return menu

with t1:
    rotar = st.checkbox("Rotar jugos sin repetir", value=True)
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        if st.button("🎲 GENERAR MENÚ ALEATORIO", type="primary", use_container_width=True):
            st.session_state['menu'] = generar_menu(semanas, personas, rotar)
            st.session_state['mp']   = {'personas':personas,'semanas':semanas,'margen':margen}

    if 'menu' in st.session_state:
        menu = st.session_state['menu']
        p    = st.session_state['mp']
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Semanas",p['semanas']); c2.metric("Personas",p['personas'])
        c3.metric("Días",p['semanas']*7); c4.metric("Servicios",p['semanas']*7*3)
        sem_sel = st.selectbox("Ver semana:", list(menu.keys()))
        for dia, srv in menu[sem_sel].items():
            with st.expander(f"📅 {dia}", expanded=(dia=='Lunes')):
                cd,ca,cc = st.columns(3)
                with cd:
                    st.markdown("**🌅 Desayuno**")
                    for k,v in srv['Desayuno'].items(): st.markdown(f"- **{k}:** {v}")
                with ca:
                    st.markdown("**☀️ Almuerzo**")
                    for k,v in srv['Almuerzo'].items(): st.markdown(f"- **{k}:** {v}")
                with cc:
                    st.markdown("**🌙 Cena**")
                    for k,v in srv['Cena'].items(): st.markdown(f"- **{k}:** {v}")

        def exportar_menu(menu,n_per,mgn):
            out = BytesIO()
            with pd.ExcelWriter(out,engine='openpyxl') as w:
                for sem,dias_m in menu.items():
                    rows=[]
                    for dia,srv in dias_m.items():
                        for sn,preps in srv.items():
                            for comp,prep in preps.items():
                                src=des if sn=='Desayuno' else alm
                                ing=src[src['PREP']==prep][['ING','g_p']].drop_duplicates()
                                if len(ing)==0:
                                    rows.append({'DIA':dia,'SERVICIO':sn,'COMPONENTE':comp,
                                                 'PREPARACION':prep,'INGREDIENTE':'','g/persona':0,
                                                 '#PERSONAS':n_per,'TOTAL Kg':0,f'+{mgn}%':0})
                                else:
                                    for _,r in ing.iterrows():
                                        g=float(r['g_p']) if pd.notna(r['g_p']) else 0
                                        tb=round(g*n_per/1000,3)
                                        rows.append({'DIA':dia,'SERVICIO':sn,'COMPONENTE':comp,
                                                     'PREPARACION':prep,'INGREDIENTE':r['ING'],
                                                     'g/persona':g,'#PERSONAS':n_per,
                                                     'TOTAL Kg':tb,f'+{mgn}%':round(tb*(1+mgn/100),3)})
                    pd.DataFrame(rows).to_excel(w,sheet_name=sem.replace(' ','_')[:31],index=False)
            out.seek(0); return out

        st.markdown("---")
        ca,cb = st.columns(2)
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        with ca:
            st.download_button("📊 Descargar Excel",
                data=exportar_menu(menu,p['personas'],p['margen']),
                file_name=f"Menu_DURANGAR_{fecha}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with cb:
            if st.button("🔄 Regenerar",use_container_width=True):
                st.session_state['menu']=generar_menu(semanas,personas,rotar)
                st.session_state['mp']={'personas':personas,'semanas':semanas,'margen':margen}
                st.rerun()
    else:
        st.info("Configura los parámetros y haz clic en **GENERAR MENÚ ALEATORIO**.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — MRP
# ══════════════════════════════════════════════════════════════════════════════
with t2:
    if 'menu' not in st.session_state:
        st.warning("Primero genera un menú en la pestaña 1.")
    else:
        menu=st.session_state['menu']; p=st.session_state['mp']
        n_per=p['personas']; mgn=p['margen']/100
        rows=[]
        for sem,dias_m in menu.items():
            for dia,srv in dias_m.items():
                for sn,preps in srv.items():
                    for comp,prep in preps.items():
                        src=des if sn=='Desayuno' else alm
                        ing=src[src['PREP']==prep][['ING','g_p']].drop_duplicates()
                        for _,r in ing.iterrows():
                            g=float(r['g_p']) if pd.notna(r['g_p']) else 0
                            rows.append({'semana':sem,'dia':dia,'servicio':sn,
                                         'ingrediente':r['ING'],'total_kg':round(g*n_per/1000,3)})
        df_det=pd.DataFrame(rows)
        mrp=(df_det.groupby('ingrediente')['total_kg'].sum()
             .reset_index().rename(columns={'total_kg':'Base (kg)'}))
        mrp['Margen (kg)']=(mrp['Base (kg)']*mgn).round(3)
        mrp['Pedido (kg)']=(mrp['Base (kg)']*(1+mgn)).round(3)
        mrp=mrp.sort_values('Pedido (kg)',ascending=False).reset_index(drop=True)
        st.session_state['mrp']=mrp

        c1,c2,c3,c4=st.columns(4)
        c1.metric("Ingredientes",len(mrp))
        c2.metric("Base (kg)",f"{mrp['Base (kg)'].sum():,.1f}")
        c3.metric(f"Margen {p['margen']}% (kg)",f"{mrp['Margen (kg)'].sum():,.1f}")
        c4.metric("PEDIDO TOTAL (kg)",f"{mrp['Pedido (kg)'].sum():,.1f}")
        st.markdown("---")
        st.dataframe(mrp,use_container_width=True,hide_index=True)
        out=BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as w:
            mrp.to_excel(w,sheet_name='MRP',index=False)
            df_det.to_excel(w,sheet_name='Detalle',index=False)
        out.seek(0)
        fecha=datetime.datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button("📊 Descargar MRP Excel",data=out,
            file_name=f"MRP_DURANGAR_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — VRP CLARKE & WRIGHT
# USO OBLIGATORIO DE TODA LA FLOTA: 5 vehículos = 5 rutas
# ══════════════════════════════════════════════════════════════════════════════
with t3:
    st.markdown("#### Demanda quincenal por campo (kg)")

    DEM_REF = {'Base H&P':378,'Carrizales':5970,'Yenac':4290,
               'Careto':2500,'Corcel':11371,'Arrendajo':3589}

    if 'mrp' in st.session_state:
        total_mrp=st.session_state['mrp']['Pedido (kg)'].sum()
        total_ref=sum(DEM_REF.values())
        dem_base={c:max(50,round(DEM_REF[c]/total_ref*total_mrp)) for c in CAMPOS}
        st.success(f"✔ Demanda calculada desde el MRP — Total: {total_mrp:,.1f} kg")
    else:
        dem_base=DEM_REF.copy()
        st.info("Genera un menú en Pestaña 1 para usar demanda real del MRP.")

    cols=st.columns(3)
    DEM={}
    for idx,campo in enumerate(CAMPOS):
        with cols[idx%3]:
            DEM[campo]=st.number_input(campo,0,500000,
                                        int(dem_base[campo]),50,key=f"v3_{campo}")

    st.markdown("---")

    if st.button("🚚 CALCULAR VRP — USO COMPLETO DE FLOTA", type="primary", use_container_width=True):

        # ══════════════════════════════════════════════════════════════════════
        # PASO 1: Calcular tabla de ahorros Clarke & Wright
        # S(i,j) = c(0,i) + c(0,j) − c(i,j)
        # ══════════════════════════════════════════════════════════════════════
        # Usamos costo promedio de flota para los ahorros base
        COSTO_PROM_CAMION    = 767   # $/km real promedio flota Durangar (TSN320+WEQ)
        COSTO_PROM_CAMIONETA = 218   # $/km real promedio flota Durangar (JTX761+SXD+SZO)

        def ckm_campo(campo):
            return COSTO_PROM_CAMIONETA if campo=='Base H&P' else COSTO_PROM_CAMION

        ahorros=[]
        for i,ci in enumerate(CAMPOS):
            for j,cj in enumerate(CAMPOS):
                if j<=i: continue
                c0i   = DIST_BODEGA[ci]*2*ckm_campo(ci)
                c0j   = DIST_BODEGA[cj]*2*ckm_campo(cj)
                ckm_ij= max(ckm_campo(ci),ckm_campo(cj))
                cij   = dij(ci,cj)*ckm_ij
                ahorro= c0i+c0j-cij
                t     = T_VIAJE[ci]+T_DESC+dij(ci,cj)*1.45+T_VIAJE[cj]+T_DESC
                dem_p = DEM[ci]+DEM[cj]
                fact  = t<=ventana
                ahorros.append({
                    'Rank':0,'Campo i':ci,'Campo j':cj,
                    'Ahorro S(i,j) $':round(ahorro),
                    'T. ruta (min)':round(t),
                    'Demanda i+j (kg)':dem_p,
                    f'✓ {ventana} min':'✓' if fact else '✗',
                    'Ruta':f"Bodega→{ci}→{cj}→Bodega",
                })

        df_ah=(pd.DataFrame(ahorros)
               .sort_values('Ahorro S(i,j) $',ascending=False)
               .reset_index(drop=True))
        df_ah['Rank']=df_ah.index+1
        df_ah=df_ah.set_index('Rank')

        # ══════════════════════════════════════════════════════════════════════
        # PASO 2: Asignar vehículos a rutas — USO OBLIGATORIO DE TODA LA FLOTA
        # Estrategia:
        #   • Ordenar campos de MAYOR a MENOR demanda
        #   • Ordenar vehículos de MAYOR a MENOR capacidad
        #   • Asignar el vehículo más grande al campo de mayor demanda
        #   • Si hay más campos que vehículos → consolidar campos en una ruta
        #   • Si hay más vehículos que campos → algunos hacen rutas combinadas
        # ══════════════════════════════════════════════════════════════════════

        # Ordenar campos por demanda descendente
        campos_dem = sorted(CAMPOS, key=lambda c: DEM[c], reverse=True)
        # Ordenar flota por capacidad descendente
        flota_ord  = sorted(FLOTA, key=lambda v: v['cap_kg'], reverse=True)

        # Tenemos 6 campos y 5 vehículos → hay que consolidar 1 par
        # Usar el par de mayor ahorro que sea factible para decidir qué consolidar
        pares_fact = df_ah[df_ah[f'✓ {ventana} min']=='✓'].reset_index()

        # Encontrar el mejor par para consolidar (mayor ahorro ejecutable)
        par_consolidado = None
        for _,row in pares_fact.iterrows():
            ci,cj = row['Campo i'],row['Campo j']
            dem_p = DEM[ci]+DEM[cj]
            # El par consolidado necesita un vehículo con suficiente capacidad
            veh_posibles = [v for v in flota_ord if v['cap_kg']>=dem_p
                            and (not (NECESITA_FRIO[ci] or NECESITA_FRIO[cj]) or v['refrig'])]
            if veh_posibles:
                par_consolidado = (ci,cj,veh_posibles[0])
                break

        # Si no hay par consolidable, tomar los dos de menor demanda
        if not par_consolidado:
            ci,cj = campos_dem[-2],campos_dem[-1]
            dem_p = DEM[ci]+DEM[cj]
            v_par = flota_ord[0]  # usar el más grande
            par_consolidado = (ci,cj,v_par)

        ci_cons,cj_cons,v_cons = par_consolidado
        campos_solos = [c for c in CAMPOS if c not in [ci_cons,cj_cons]]
        # campos_solos tiene 4 campos → 4 vehículos restantes
        flota_rest = [v for v in flota_ord if v['placa']!=v_cons['placa']]

        # Ordenar campos solos por demanda y asignar vehículos restantes por capacidad
        campos_solos_ord = sorted(campos_solos, key=lambda c: DEM[c], reverse=True)

        # Construir rutas
        rutas = []

        # Ruta consolidada
        km_cons  = DIST_BODEGA[ci_cons]+dij(ci_cons,cj_cons)+DIST_BODEGA[cj_cons]
        t_cons   = T_VIAJE[ci_cons]+T_DESC+dij(ci_cons,cj_cons)*1.45+T_VIAJE[cj_cons]+T_DESC
        costo_cons = round(km_cons*v_cons['costo_km'])
        # Ahorro vs individual
        v_ci = flota_rest[0]; v_cj = flota_rest[1] if len(flota_rest)>1 else flota_rest[0]
        c_ind_ci = DIST_BODEGA[ci_cons]*2*v_ci['costo_km']
        c_ind_cj = DIST_BODEGA[cj_cons]*2*v_cj['costo_km']
        ahorro_cons = round(c_ind_ci+c_ind_cj-costo_cons)
        estado_cons = ('✓ Factible 300 min' if t_cons<=300
                       else f'⚠ Ventana {ventana} min' if t_cons<=ventana
                       else '✗ Requiere salida especial')
        rutas.append({
            'Ruta':'Ruta 1 (CONSOLIDADA)',
            'Secuencia':f"Bodega→{ci_cons}→{cj_cons}→Bodega",
            'Vehículo':f"{v_cons['placa']} ({v_cons['tipo']})",
            'Placa':v_cons['placa'],
            'Cap. veh. (kg)':v_cons['cap_kg'],
            'Demanda total (kg)':DEM[ci_cons]+DEM[cj_cons],
            'Km':round(km_cons),
            'T. ruta (min)':round(t_cons),
            'Costo ($)':costo_cons,
            'Ahorro vs individual ($)':max(0,ahorro_cons),
            'Estado':estado_cons,
        })

        # Rutas individuales con vehículo asignado por capacidad
        for idx,campo in enumerate(campos_solos_ord):
            v = flota_rest[idx] if idx<len(flota_rest) else flota_rest[-1]
            km    = DIST_BODEGA[campo]*2
            t     = T_VIAJE[campo]*2+T_DESC
            ckm_ind = COSTO_PROM_CAMIONETA if campo=='Base H&P' else COSTO_PROM_CAMION
            costo = round(km * ckm_ind)
            estado= ('✓ Factible 300 min' if t<=300
                     else f'⚠ Ventana {ventana} min' if t<=ventana
                     else '✗ Requiere salida especial (3:25 am)')
            rutas.append({
                'Ruta':f'Ruta {idx+2}',
                'Secuencia':f"Bodega→{campo}→Bodega",
                'Vehículo':f"{v['placa']} ({v['tipo']})",
                'Placa':v['placa'],
                'Cap. veh. (kg)':v['cap_kg'],
                'Demanda total (kg)':DEM[campo],
                'Km':km,
                'T. ruta (min)':round(t),
                'Costo ($)':costo,
                'Ahorro vs individual ($)':0,
                'Estado':estado,
            })

        df_vrp = pd.DataFrame(rutas)

        # Verificar que se usaron todos los vehículos
        placas_usadas = set(df_vrp['Placa'].tolist())
        placas_flota  = set(v['placa'] for v in FLOTA)
        placas_faltantes = placas_flota - placas_usadas

        # Situación actual: TODOS los campos usan camión ($767/km) — sin optimización
        # Dato exacto del Excel ClarkeWright_VRP_Durangar.xlsx: $1,213,394/quincena
        costo_actual = sum(DIST_BODEGA[c] * 2 * COSTO_PROM_CAMION for c in CAMPOS)
        # = $1,213,394/quincena | Ahorro modelo VRP: $147,286/quincena | $3,534,864/año

        st.session_state['vrp_result'] = {
            'ahorros':df_ah,'vrp':df_vrp,
            'dem':DEM.copy(),'costo_actual':costo_actual,
            'placas_faltantes':placas_faltantes,
        }
        st.rerun()

    # ── MOSTRAR RESULTADOS ────────────────────────────────────────────────────
    if 'vrp_result' in st.session_state:
        res        = st.session_state['vrp_result']
        df_ah      = res['ahorros']
        df_vrp     = res['vrp']
        dem_us     = res['dem']
        costo_act  = res['costo_actual']
        faltantes  = res['placas_faltantes']

        # Alerta si algún vehículo no se usó
        if faltantes:
            st.error(f"⚠ Vehículos no asignados: {', '.join(faltantes)}")
        else:
            st.success("✅ Los 5 vehículos de la flota están asignados — uso al 100%")

        # KPIs
        costo_vrp = df_vrp['Costo ($)'].sum()
        ahorro_q  = costo_act - costo_vrp
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Rutas generadas",       len(df_vrp))
        c2.metric("Vehículos usados",      f"{len(df_vrp['Placa'].unique())}/5")
        c3.metric("Costo modelo VRP ($)",  f"${costo_vrp:,.0f}")
        c4.metric("Ahorro anual estimado", f"${max(0,ahorro_q)*24:,.0f}")

        st.markdown("---")

        # Tabla de asignación de flota
        st.markdown("#### Resumen de asignación — 1 vehículo por ruta")
        resumen = df_vrp[['Ruta','Vehículo','Cap. veh. (kg)',
                           'Demanda total (kg)','Km','T. ruta (min)',
                           'Costo ($)','Ahorro vs individual ($)','Estado']]
        st.dataframe(resumen, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("#### Detalle visual de rutas")
        for _,r in df_vrp.iterrows():
            e   = str(r['Estado'])
            cls = ('ruta-ok' if '✓' in e
                   else 'ruta-warn' if '⚠' in e
                   else 'ruta-bad')
            ocup= round(r['Demanda total (kg)']/r['Cap. veh. (kg)']*100,1)
            st.markdown(f"""
            <div class="{cls}">
              <strong>{r['Ruta']}</strong> — {r['Secuencia']}<br>
              🚛 <strong>{r['Vehículo']}</strong> &nbsp;|&nbsp;
              📦 {r['Demanda total (kg)']:,} kg / {r['Cap. veh. (kg)']:,} kg
              ({ocup}% ocupación) &nbsp;|&nbsp;
              📏 {r['Km']} km &nbsp;|&nbsp; ⏱ {r['T. ruta (min)']} min &nbsp;|&nbsp;
              💰 ${r['Costo ($)']:,.0f}<br>
              💡 Ahorro vs individual: ${r['Ahorro vs individual ($)']:,.0f}
              &nbsp;|&nbsp; {e}
            </div>""", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### Tabla de ahorros Clarke & Wright")
        st.markdown("**S(i,j) = c(0,i) + c(0,j) − c(i,j)** | "
                    "Ordenado de mayor a menor ahorro")
        st.dataframe(df_ah, use_container_width=True)

        # Exportar
        out=BytesIO()
        with pd.ExcelWriter(out,engine='openpyxl') as w:
            df_vrp.to_excel(w,sheet_name='Rutas_VRP_Flota_Completa',index=False)
            df_ah.to_excel(w,sheet_name='Ahorros_CW',index=True)
            pd.DataFrame([{'Campo':k,'Demanda kg':v}
                           for k,v in dem_us.items()]).to_excel(
                w,sheet_name='Demanda_usada',index=False)
            pd.DataFrame(FLOTA).to_excel(w,sheet_name='Flota',index=False)
        out.seek(0)
        fecha=datetime.datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button("📊 Descargar VRP Excel completo",data=out,
            file_name=f"VRP_DURANGAR_{fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    else:
        st.info("Ajusta la demanda y haz clic en **CALCULAR VRP**.")

# ── FOOTER ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#888;font-size:12px'>"
    f"NETO DURANGAR S.A.S. | Sistema MRP + VRP | Tesis de Grado | "
    f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}</p>",
    unsafe_allow_html=True)
