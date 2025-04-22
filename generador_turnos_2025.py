# -*- coding: utf-8 -*-
"""Generador de Turnos 2025 con Streamlit"""

import streamlit as st
import pandas as pd, numpy as np, math, itertools, time, random, os
from pyworkforce.scheduling import MinAbsDifference

# --- UI: Título y carga de archivo -------------------------------------
st.title("Generador de Turnos 2025")
uploaded = st.file_uploader("Sube tu archivo Excel (Requerido.xlsx)", type=["xlsx"])
if not uploaded:
    st.warning("Por favor, sube el archivo Requerido.xlsx para continuar.")
    st.stop()

# --- Leer datos ---------------------------------------------------------
df = pd.read_excel(uploaded)

# --- Parámetros de configuración ---------------------------------------
MAX_ITER      = 20      # None = infinito hasta Ctrl+C
TIME_SOLVER   = 120.0   # segundos por ejecución del solver
SEED_START    = 0       
PERTURB_NOISE = 0.20     # 20% de ruido a la distribución base
MIN_REST_PCT  = 0.05     # cada día libre al menos 5%
IN_OFF        = 0.00    # % extra modalidad in-office
OUT_OFF       = 0.00    # % extra modalidad out-office

# --- Inicializar RNG ---------------------------------------------------
rng = np.random.default_rng(SEED_START)

# --- 1. DEMANDA ---------------------------------------------------------
required_resources = [[] for _ in range(7)]
for _, r in df.iterrows():
    required_resources[int(r['Día'])-1].append(r['Suma de Agentes Requeridos Erlang'])
assert all(len(day) == 24 for day in required_resources), "Cada día debe tener 24 periodos"

daily_demand   = [sum(day) for day in required_resources]
inv_weights    = [1 / max(1, x) for x in daily_demand]
base_rest_dist = np.array(inv_weights) / sum(inv_weights)

dias_semana = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

# --------------------------------------------------------------
# 2. DEFINICIÓN DE TURNOS  (diccionario completo)
# --------------------------------------------------------------
shifts_coverage = {
    # ----------------------------------------------------------
    # TURNOS FULL‑TIME 8H
    # ----------------------------------------------------------
    "FT_00:00_1":[1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_00:00_2":[1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_00:00_3":[1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_01:00_1":[0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_01:00_2":[0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_01:00_3":[0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_02:00_1":[0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_02:00_2":[0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_02:00_3":[0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_03:00_1":[0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_03:00_2":[0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_03:00_3":[0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0],
    "FT_04:00_1":[0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0],
    "FT_04:00_2":[0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0],
    "FT_04:00_3":[0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0],
    "FT_05:00_1":[0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0],
    "FT_05:00_2":[0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0],
    "FT_05:00_3":[0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0],
    "FT_06:00_1":[0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0],
    "FT_06:00_2":[0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0],
    "FT_06:00_3":[0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0],
    "FT_07:00_1":[0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0],
    "FT_07:00_2":[0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0],
    "FT_07:00_3":[0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0],
    "FT_08:00_1":[0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0,0],
    "FT_08:00_2":[0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0],
    "FT_08:00_3":[0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0],
    "FT_09:00_1":[0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0,0],
    "FT_09:00_2":[0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0],
    "FT_09:00_3":[0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0],
    "FT_10:00_1":[0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0,0],
    "FT_10:00_2":[0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0],
    "FT_10:00_3":[0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0,0],
    "FT_11:00_1":[0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0,0],
    "FT_11:00_2":[0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0],
    "FT_11:00_3":[0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0,0],
    "FT_12:00_1":[0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0,0],
    "FT_12:00_2":[0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0],
    "FT_12:00_3":[0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0,0],
    "FT_13:00_1":[0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0,0],
    "FT_13:00_2":[0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0],
    "FT_13:00_3":[0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0,0],
    "FT_14:00_1":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1,0],
    "FT_14:00_2":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0],
    "FT_14:00_3":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1,0],
    "FT_15:00_1":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1,1],
    "FT_15:00_2":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1],
    "FT_15:00_3":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1,1],
    "FT_16:00_1":[1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1,1],
    "FT_16:00_2":[1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1],
    "FT_16:00_3":[1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1,1],
    "FT_17:00_1":[1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1,1],
    "FT_17:00_2":[1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1,1],
    "FT_17:00_3":[1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,0,1],
    "FT_18:00_1":[1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1,1],
    "FT_18:00_2":[1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,1],
    "FT_18:00_3":[1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,1],
    "FT_19:00_1":[1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,0,1],
    "FT_19:00_2":[1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1],
    "FT_19:00_3":[0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1],
    "FT_20:00_1":[1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1],
    "FT_20:00_2":[0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1],
    "FT_20:00_3":[1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1],
    "FT_21:00_1":[0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1],
    "FT_21:00_2":[1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1],
    "FT_21:00_3":[1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1],
    "FT_22:00_1":[1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1],
    "FT_22:00_2":[1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1],
    "FT_22:00_3":[1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1],
    "FT_23:00_1":[1,1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
    "FT_23:00_2":[1,1,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
    "FT_23:00_3":[1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],

    # ----------------------------------------------------------
    # TURNOS PART‑TIME 4H
    # ----------------------------------------------------------
    "00_4":[1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "01_4":[0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "02_4":[0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "03_4":[0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "04_4":[0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "05_4":[0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "06_4":[0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "07_4":[0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "08_4":[0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0],
    "09_4":[0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0],
    "10_4":[0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0,0],
    "11_4":[0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,0],
    "12_4":[0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0],
    "13_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0,0],
    "14_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0],
    "15_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0],
    "16_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0],
    "17_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0],
    "18_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0],
    "19_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,0],
    "20_4":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1],
    "21_4":[1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1],
    "22_4":[1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1],
    "23_4":[1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1],
}


# --- FUNCIONES AUXILIARES ------------------------------------------------
def adjust_required(dist):
    adj = []
    for i, day in enumerate(required_resources):
        factor   = 1 / (1 - dist[i])
        base_adj = [math.ceil(r * factor) for r in day]
        in_off   = [math.ceil(r * (1 + IN_OFF)) for r in base_adj]
        out_off  = [math.ceil(r * (1 + OUT_OFF)) for r in base_adj]
        adj.append((in_off, out_off))
    return adj


def build_scheduler(adj, seed):
    return MinAbsDifference(
        num_days=7,
        periods=24,
        shifts_coverage=shifts_coverage,
        required_resources=[d[0] for d in adj],
        max_period_concurrency=5000,
        max_shift_concurrency=300,
        max_search_time=TIME_SOLVER,
        num_search_workers=8,
        random_seed=seed
    )


def greedy_day_off_assignment(n_shifts, dist):
    result = []
    counts = np.zeros(7, int)
    quota  = (n_shifts * dist).round().astype(int)
    for _ in range(n_shifts):
        idx = np.argmax(quota - counts)
        result.append(dias_semana[idx])
        counts[idx] += 1
    return result


def coverage_pct(sol, dist):
    if sol['status'] not in ('OPTIMAL', 'FEASIBLE'):
        return 0.0
    shift_order = list(shifts_coverage.keys())
    dayoff_map  = dict(zip(shift_order, greedy_day_off_assignment(len(shift_order), dist)))
    diff_tot, req_tot = 0, sum(map(sum, required_resources))
    for d, day_name in enumerate(dias_semana):
        for h in range(24):
            req = required_resources[d][h]
            work = sum(
                row['resources']
                for row in sol['resources_shifts']
                if row['day']==d and shifts_coverage[row['shift']][h] and dayoff_map[row['shift']]!=day_name
            )
            diff_tot += abs(work - req)
    return (1 - diff_tot / req_tot) * 100


def mutate_dist(base):
    noise = rng.normal(0, PERTURB_NOISE, 7)
    cand  = np.clip(base * (1 + noise), 1e-9, None)
    cand /= cand.sum()
    mask    = cand < MIN_REST_PCT
    deficit = (MIN_REST_PCT - cand[mask]).sum()
    cand[mask] = MIN_REST_PCT
    if deficit > 0:
        surplus = ~mask
        cand[surplus] -= deficit / surplus.sum()
        if (cand < 0).any():
            return mutate_dist(base)
    return cand / cand.sum()


def export_reports(sol, dist, tag):
    df_res = pd.DataFrame(sol['resources_shifts'])
    df_res.to_excel(f"Result_{tag}.xlsx", index=False)
    df_res.to_csv(f"Result_{tag}.csv", sep=';', index=False)
    summary = (df_res.groupby('shift')
               .agg(resources=('resources','sum'))
               .reset_index())
    summary['Personal a Contratar'] = (summary['resources']/7).round().astype(int)
    summary['Tipo de Contrato']     = summary['shift'].apply(lambda s: 'Full Time (8h)' if s.startswith('FT_') else 'Part Time (4h)')
    summary['Día de Descanso']      = greedy_day_off_assignment(len(summary), dist)
    summary['Refrigerio']            = summary.apply(
        lambda r: f"Refrigerio {r['shift'].split('_')[-1]}" if r['Tipo de Contrato'].startswith('Full') else '-', axis=1)
    summary.rename(columns={'shift':'Horario'}, inplace=True)
    summary[['Horario','Tipo de Contrato','Personal a Contratar','Día de Descanso','Refrigerio']]
    summary.to_excel(f"Plan_Contratacion_{tag}.xlsx", index=False)

# --- 5. BÚSQUEDA META-HEURÍSTICA --------------------------------------
best_cov, best_sol, best_dist = -1, None, None
iterator = range(MAX_ITER) if MAX_ITER is not None else itertools.count()
for it in iterator:
    dist = base_rest_dist.copy() if it==0 else mutate_dist(base_rest_dist)
    adj  = adjust_required(dist)
    sol  = build_scheduler(adj, SEED_START+it).solve()
    cov  = coverage_pct(sol, dist)
    if cov > best_cov:
        best_cov, best_sol, best_dist = cov, sol, dist.copy()
        print(f"[{it:03}] 🔹 NUEVO BEST {cov:5.2f}%")
    else:
        print(f"[{it:03}] cobertura {cov:5.2f}% (best {best_cov:5.2f}%)")

# --- 6. EXPORTA MEJOR RESULTADO Y BOTONES DE DESCARGA -------------------
if best_sol:
    tag = time.strftime("%Y%m%d_%H%M%S")
    export_reports(best_sol, best_dist, tag)
    st.success(f"Reportes generados. Mejor cobertura: {best_cov:5.2f}%")
    # Botones para descargar archivos
    files = [f for f in os.listdir('.') if tag in f]
    st.write("### Descarga tus reportes:")
    for fname in files:
        with open(fname, 'rb') as file:
            st.download_button(
                label=fname,
                data=file,
                file_name=fname,
                mime='application/octet-stream'
            )
else:
    st.error("No se encontró solución factible.")


