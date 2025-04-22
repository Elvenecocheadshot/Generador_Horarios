

# ===============================================================
# PLANIFICACIÓN – OPTIMIZA TURNOS Y DESCANSOS EN CONJUNTO
# ===============================================================

import pandas as pd, numpy as np, math, itertools, time, random
from pyworkforce.scheduling import MinAbsDifference

# 🔧 PARAMETROS -------------------------------------------------
MAX_ITER     = 300          # None = infinito hasta Ctrl+C
TIME_SOLVER  = 120.0         # seg. por ejecución del solver
SEED_START   = 0
ARCH_EXCEL   = "/content/sample_data/Requerido.xlsx"   # archivo de entrada
PERTURB_NOISE= 0.20          # 20 % de ruido a la distribución base
MIN_REST_PCT = 0.05          # cada día libre al menos 5 %
# --------------------------------------------------------------

rng = np.random.default_rng(SEED_START)

# 1. ------------- DEMANDA -------------------------------------
df = pd.read_excel(ARCH_EXCEL)
required_resources = [[] for _ in range(7)]
for _, r in df.iterrows():
    required_resources[int(r['Día'])-1].append(r['Suma de Agentes Requeridos Erlang'])
assert all(len(d)==24 for d in required_resources), "Cada día debe tener 24 periodos"


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


# 3. ------------- PORCENTAJES IN / OUT ------------------------
IN_OFF  = 0.00          # % extra si hay modalidad in‑office
OUT_OFF = 0.00          # % extra si hay modalidad out‑office

# 4. ------------- BASE: INVERSO DE DEMANDA --------------------
daily_demand   = [sum(d) for d in required_resources]
inv_weights    = [1 / max(1, x) for x in daily_demand]
base_rest_dist = np.array(inv_weights) / sum(inv_weights)

dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves',
               'Viernes', 'Sábado', 'Domingo']

# ----------------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------------
def adjust_required(dist):
    """Aplica la distribución de descansos a la demanda original."""
    ajustado = []
    for i, day in enumerate(required_resources):
        factor   = 1 / (1 - dist[i])                     # engorda la demanda
        adj_base = [math.ceil(r * factor) for r in day]  # redondeo hacia ↑
        adj_in   = [math.ceil(r * (1 + IN_OFF))  for r in adj_base]
        adj_out  = [math.ceil(r * (1 + OUT_OFF)) for r in adj_base]
        ajustado.append((adj_in, adj_out))               # (in, out)
    return ajustado

def build_scheduler(adj, seed):
    return MinAbsDifference(
        num_days=7,
        periods=24,
        shifts_coverage=shifts_coverage,
        required_resources=[d[0] for d in adj],          # usamos la parte “in”
        max_period_concurrency=5000,
        max_shift_concurrency=300,
        max_search_time=TIME_SOLVER,
        num_search_workers=8,
        random_seed=seed
    )

def greedy_day_off_assignment(n_turnos, dist):
    """
    Devuelve una lista de días de descanso (nombres) con la proporción dist.
    El orden del resultado coincide con el orden de los turnos recibidos.
    """
    result  = []
    counts  = np.zeros(7, int)
    quota   = (n_turnos * dist).round().astype(int)
    for _ in range(n_turnos):
        idx = np.argmax(quota - counts)                 # día con + cupo pendiente
        result.append(dias_semana[idx])
        counts[idx] += 1
    return result                                       # len == n_turnos

def coverage_pct(sol, dist):
    """
    Calcula la cobertura global respetando el día libre de cada turno.
    No se aplica (1‑rest) porque los descansos son turnos completos fuera.
    """
    if sol['status'] not in ('OPTIMAL', 'FEASIBLE'):
        return 0.0

    # --- mapa turno → día libre, en orden estable de shifts_coverage ----
    shift_order = list(shifts_coverage.keys())
    dayoff_map  = dict(zip(shift_order,
                           greedy_day_off_assignment(len(shift_order), dist)))

    diff_tot, req_tot = 0, sum(map(sum, required_resources))

    # recorre cada día/hora y suma agentes que SÍ trabajan ese día
    for d, day_name in enumerate(dias_semana):
        for h in range(24):
            req  = required_resources[d][h]
            work = 0
            for row in sol['resources_shifts']:         # filas devueltas por solver
                if row['day'] == d and shifts_coverage[row['shift']][h]:
                    if dayoff_map[row['shift']] != day_name:
                        work += row['resources']
            diff_tot += abs(work - req)

    return (1 - diff_tot / req_tot) * 100               # porcentaje

def mutate_dist(base):
    """Perturba la distribución base y asegura MIN_REST_PCT por día."""
    noise = rng.normal(0, PERTURB_NOISE, 7)
    cand  = np.clip(base * (1 + noise), 1e-9, None)
    cand /= cand.sum()

    mask    = cand < MIN_REST_PCT                       # días por debajo
    deficit = (MIN_REST_PCT - cand[mask]).sum()
    cand[mask] = MIN_REST_PCT

    if deficit > 0:                                    # resta a los demás
        surplus = ~mask
        cand[surplus] -= deficit / surplus.sum()
        if (cand < 0).any():                           # ruido extremo → recursivo
            return mutate_dist(base)
    return cand / cand.sum()

def export_reports(sol, dist, tag):
    # 1️⃣  Resultados crudos ------------------------------------------------
    df_res = pd.DataFrame(sol['resources_shifts'])
    df_res.to_excel(f"Result_{tag}.xlsx", index=False)
    df_res.to_csv  (f"Result_{tag}.csv",  sep=';', index=False)

    # 2️⃣  Plan de Contratación --------------------------------------------
    summary = df_res.groupby('shift').agg(resources=('resources', 'sum')).reset_index()
    summary['Personal a Contratar'] = (summary['resources'] / 7).round().astype(int)
    summary['Tipo de Contrato']     = summary['shift'].apply(
        lambda s: 'Full Time (8h)' if s.startswith('FT_') else 'Part Time (4h)')

    # asigna día de descanso coherente con dist
    summary['Día de Descanso'] = greedy_day_off_assignment(len(summary), dist)

    # refrigerio solo full‑time
    summary['Refrigerio'] = summary.apply(
        lambda r: f"Refrigerio {r['shift'].split('_')[-1]}"
        if r['Tipo de Contrato'].startswith('Full') else '-', axis=1)

    summary.rename(columns={'shift': 'Horario'}, inplace=True)
    cols = ['Horario', 'Tipo de Contrato', 'Personal a Contratar',
            'Día de Descanso', 'Refrigerio']
    summary[cols].to_excel(f"Plan_Contratacion_{tag}.xlsx", index=False)
    summary[cols].to_csv  (f"Plan_Contratacion_{tag}.csv",  sep=';', index=False)

    # 3️⃣  Plan global de descansos ----------------------------------------
    total_agents = summary['Personal a Contratar'].sum()
    descansos = (total_agents * dist).round().astype(int)
    while descansos.sum() < total_agents:               # ajusta redondeo
        descansos[np.argmin(descansos)] += 1

    pd.DataFrame({'Día': dias_semana,
                  'Personal Descansando': descansos,
                  'Porcentaje': (dist * 100).round(2),
                  'Demanda del Día': daily_demand}) \
        .to_excel(f"Plan_Descansos_{tag}.xlsx", index=False)

    # 4️⃣  Verificación de cobertura --------------------------------------
    staff_map  = summary.set_index('Horario')['Personal a Contratar'].to_dict()
    dayoff_map = summary.set_index('Horario')['Día de Descanso'].to_dict()

    cov_rows = []
    for d, day_name in enumerate(dias_semana):
        for h in range(24):
            req  = required_resources[d][h]
            work = sum(staff_map[t]
                       for t, vec in shifts_coverage.items()
                       if vec[h] and dayoff_map[t] != day_name)
            cov_rows.append({'Día': d + 1,
                             'Día Semana': day_name,
                             'Hora': f"{h:02}:00",
                             'Requeridos': req,
                             'Asignados': work,
                             'Diferencia': work - req})
    pd.DataFrame(cov_rows).to_excel(f"Verificación_Cobertura_{tag}.xlsx",
                                    index=False)

# ----------------------------------------------------------------
# 5. BÚSQUEDA META‑HEURÍSTICA
# ----------------------------------------------------------------
best_cov  = -1
best_sol  = None
best_dist = None

try:
    iterator = itertools.count() if MAX_ITER is None else range(MAX_ITER)
    for it in iterator:
        dist = mutate_dist(base_rest_dist) if it else base_rest_dist.copy()
        adj  = adjust_required(dist)
        sol  = build_scheduler(adj, SEED_START + it).solve()
        cov  = coverage_pct(sol, dist)
        if cov > best_cov:
            best_cov, best_sol, best_dist = cov, sol, dist.copy()
            print(f"[{it:03}] 🔹 NUEVO BEST {cov:5.2f}%  costo {sol['cost']}")
        else:
            print(f"[{it:03}] cobertura {cov:5.2f}%  (best {best_cov:5.2f}%)")
except KeyboardInterrupt:
    print("\n⏹  Detenido por el usuario")

# ----------------------------------------------------------------
# 6. EXPORTA MEJOR RESULTADO
# ----------------------------------------------------------------
if best_sol:
    tag = time.strftime("%Y%m%d_%H%M%S")
    print(f"\n➡️  Mejor cobertura: {best_cov:5.2f}% — exportando ({tag}) …")
    export_reports(best_sol, best_dist, tag)
    print("✅  Reportes generados.")
else:
    print("No se encontró solución factible.")