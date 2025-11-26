# -*- coding: utf-8 -*-
"""
Agrega respuestas aleatorias a src/source/respuestas_cuestionario.xlsx
usando áreas/puestos de src/source/11. SAN VICENTE MATRIZ VOTME 2024 Faenacción.xlsx.

- Mantiene las columnas que tus páginas esperan (Respuestas + Parametros).
- NO elimina filas existentes; solo agrega nuevas y continúa el ID.
- Genera fechas "vigentes" y "vencidas" (según --expired-ratio).
- Usa Faker (si está instalada) para nombres realistas; si no, usa un fallback.
"""

import os
import math
import argparse
from datetime import datetime, timedelta
from collections import defaultdict
import random
import pandas as pd

# ----------------- CLI -----------------
def parse_args():
    ap = argparse.ArgumentParser(description="Agregar respuestas aleatorias por área.")
    ap.add_argument("--project-root", default=".", help="Raíz del proyecto (donde está src/).")
    ap.add_argument("--matrix-name", default="11. SAN VICENTE MATRIZ VOTME 2024 Faenacción.xlsx",
                    help="Nombre del Excel de matriz dentro de src/source/")
    ap.add_argument("--out-name", default="respuestas_cuestionario.xlsx",
                    help="Nombre del Excel de salida dentro de src/source/")
    ap.add_argument("--add-per-area", type=int, default=12,
                    help="Cantidad de nuevas respuestas a agregar por área si el área NO tiene Total definido.")
    ap.add_argument("--expired-ratio", type=float, default=0.10,
                    help="Proporción de respuestas con fecha vencida (0..1).")
    ap.add_argument("--seed", type=int, default=321, help="Semilla aleatoria para reproducibilidad.")
    return ap.parse_args()

# ----------------- Zonas / headers -----------------
ZONAS = [
    "Cuello","Hombro Derecho","Hombro Izquierdo","Codo/antebrazo Derecho","Codo/antebrazo Izquierdo",
    "Muñeca/mano Derecha","Muñeca/mano Izquierda","Espalda Alta","Espalda Baja",
    "Caderas/nalgas/muslos","Rodillas (una o ambas)","Pies/tobillos (uno o ambos)"
]

def default_headers():
    base = [
        "ID","Fecha","Area","Puesto_de_Trabajo","Nombre_Trabajador","Sexo","Edad","Diestro_Zurdo",
        "Temporadas_previas","N_temporadas","Tiempo_en_trabajo_meses","Actividad_previa","Otra_actividad","Otra_actividad_cual"
    ]
    zona_cols = []
    for z in ZONAS:
        zona_cols += [f"{z}__12m", f"{z}__Incap", f"{z}__Dolor12m", f"{z}__7d", f"{z}__Dolor7d"]
    return base + zona_cols

# ----------------- Faker opcional -----------------
def build_name_generator(seed=321):
    try:
        from faker import Faker
        fk = None
        for loc in ("es_CL", "es_ES", "es_MX", "es"):
            try:
                fk = Faker(loc)
                break
            except Exception:
                fk = None
        if fk is None:
            raise RuntimeError("Faker locale not found")
        Faker.seed(seed)
        def gen(sex):
            # Faker no siempre condiciona por sexo en todos los locales, pero sirve igual
            return fk.name()
        return gen
    except Exception:
        rnd = random.Random(seed)
        nombres_m = ["Juan","Pedro","Carlos","Luis","Jorge","Diego","Andrés","Mauricio","Sebastián",
                     "Felipe","Matías","Gonzalo","Nicolás","Cristóbal","Hernán","Rodrigo","Tomás"]
        nombres_f = ["María","Ana","Carolina","Daniela","Camila","Valentina","Francisca","Fernanda",
                     "Josefina","Antonia","Constanza","Isidora","Catalina","Paz","Trinidad","Sofía"]
        apellidos = ["González","Muñoz","Rojas","Díaz","Pérez","Soto","Contreras","Silva","Martínez",
                     "Sepúlveda","Morales","Gutiérrez","Castro","Vargas","Romero","Herrera","Flores"]
        def gen(sex):
            if sex == "Hombre":
                return f"{rnd.choice(nombres_m)} {rnd.choice(apellidos)}"
            else:
                return f"{rnd.choice(nombres_f)} {rnd.choice(apellidos)}"
        return gen

# ----------------- Utilidades -----------------
def safe_int(x):
    try:
        return int(str(x).split(".")[0])
    except:
        return 0

def random_fecha(rnd, expired_ratio=0.40):
    today = datetime.today().date()
    if rnd.random() < expired_ratio:
        # vencido: > 365 días (entre 400 y 900)
        days = rnd.randint(400, 900)
    else:
        # vigente: dentro de ~10 meses
        days = rnd.randint(0, 300)
    return (today - timedelta(days=days)).strftime("%d/%m/%Y")

def fill_zone_answers(rnd, row):
    for z in ZONAS:
        has12 = rnd.random() < 0.35
        row[f"{z}__12m"] = "SI" if has12 else "NO"
        if has12:
            row[f"{z}__Incap"] = rnd.choice(["","NO","SI","NO","NO"])
            row[f"{z}__Dolor12m"] = str(rnd.randint(1,10))
            last7 = rnd.random() < 0.25
            row[f"{z}__7d"] = "SI" if last7 else "NO"
            row[f"{z}__Dolor7d"] = str(rnd.randint(1,10)) if last7 else ""
        else:
            row[f"{z}__Incap"]    = ""
            row[f"{z}__Dolor12m"] = ""
            row[f"{z}__7d"]       = ""
            row[f"{z}__Dolor7d"]  = ""

# ----------------- Leer matriz: pares (Área, Puesto) + dotación H/M -----------------
def read_matrix_pairs(matrix_path):
    if not os.path.exists(matrix_path):
        return [], {}
    xls = pd.ExcelFile(matrix_path)
    # localizar hoja "inicial"
    sheet = None
    for n in xls.sheet_names:
        if "inicial" in n.lower() or "inicio" in n.lower():
            sheet = n
            break
    if sheet is None:
        sheet = xls.sheet_names[0]

    df = pd.read_excel(matrix_path, sheet_name=sheet, header=None, dtype=object)

    pairs = []
    area_counts = {}
    nrows, ncols = df.shape
    # B=1, C=2, H=7, I=8; datos desde fila 3 (índice 2)
    for r in range(2, nrows):
        area   = df.iat[r,1] if 1 < ncols else None
        puesto = df.iat[r,2] if 2 < ncols else None
        h      = df.iat[r,7] if 7 < ncols else 0
        m      = df.iat[r,8] if 8 < ncols else 0
        if pd.isna(area) and pd.isna(puesto):
            continue
        area   = "" if pd.isna(area)   else str(area).strip()
        puesto = "" if pd.isna(puesto) else str(puesto).strip()
        if area or puesto:
            pairs.append((area, puesto))
        if area:
            try:
                hval = int(float(h)) if h is not None and not pd.isna(h) else 0
            except:
                hval = 0
            try:
                mval = int(float(m)) if m is not None and not pd.isna(m) else 0
            except:
                mval = 0
            c = area_counts.get(area, {"Hombres":0, "Mujeres":0})
            c["Hombres"] += hval
            c["Mujeres"] += mval
            area_counts[area] = c

    # de-dup manteniendo orden
    seen = set()
    uniq_pairs = []
    for a, p in pairs:
        key = (a, p)
        if key not in seen and (a or p):
            uniq_pairs.append(key)
            seen.add(key)

    for a, c in area_counts.items():
        c["Total"] = c["Hombres"] + c["Mujeres"]

    return uniq_pairs, area_counts

# ----------------- Principal -----------------
def main():
    args = parse_args()
    rnd = random.Random(args.seed)
    gen_nombre = build_name_generator(args.seed)

    src_dir = os.path.join(args.project_root, "src", "source")
    os.makedirs(src_dir, exist_ok=True)

    matrix_path = os.path.join(src_dir, args.matrix_name)
    out_path    = os.path.join(src_dir, args.out_name)

    headers = default_headers()

    # Cargar existente
    if os.path.exists(out_path):
        try:
            x2 = pd.ExcelFile(out_path)
            if "Respuestas" in x2.sheet_names:
                df_res = pd.read_excel(out_path, sheet_name="Respuestas", dtype=str)
                for h in headers:
                    if h not in df_res.columns:
                        df_res[h] = ""
                df_res = df_res[headers]
            else:
                df_res = pd.DataFrame(columns=headers)
        except Exception:
            df_res = pd.DataFrame(columns=headers)
    else:
        df_res = pd.DataFrame(columns=headers)

    max_id = df_res["ID"].apply(safe_int).max() if len(df_res) else 0

    # Leer áreas/puestos de la matriz
    pairs, area_counts = read_matrix_pairs(matrix_path)
    if not pairs:
        # Fallback mínimo si no encuentra la matriz
        pairs = [
            ("Colgado", "Operario de colgado"),
            ("Empaque - Chiller", "Operario de empaque - chiller"),
            ("Eviscerado", "Operario de eviscerado"),
            ("Garras", "Operario de garras"),
            ("Sacrificio", "Operario de sacrificio"),
        ]

        area_counts = {
            "Colgado":           {"Hombres": 36, "Mujeres": 0,  "Total": 36},
            "Empaque - Chiller": {"Hombres": 81, "Mujeres": 3,  "Total": 84},
            "Eviscerado":        {"Hombres": 87, "Mujeres": 28, "Total": 115},
            "Garras":            {"Hombres": 23, "Mujeres": 17, "Total": 40},
            "Sacrificio":        {"Hombres": 8,  "Mujeres": 2,  "Total": 10},
        }

    # Puestos por área
    puestos_por_area = defaultdict(list)
    for a, p in pairs:
        if a:
            puestos_por_area[a].append(p or "")

    # ----------------- Generar nuevas respuestas respetando Total por área -----------------
    nuevas = []

    # cuántas respuestas ya existen por área en el Excel
    existing_counts = df_res["Area"].value_counts().to_dict() if not df_res.empty else {}

    # todas las áreas conocidas (por matriz o fallback)
    all_areas = sorted({a for a, _ in pairs if a} | set(area_counts.keys()))

    for area in all_areas:
        info_area = area_counts.get(area, {})

        # Si el área tiene Total definido, lo usamos. Si no, usamos add-per-area.
        target_total = info_area.get("Total")
        if target_total is None or target_total <= 0:
            # comportamiento de respaldo para áreas sin dotación conocida
            n_to_add = args.add_per-area
        else:
            current_total = existing_counts.get(area, 0)
            n_to_add = max(target_total - current_total, 0)

        # si ya llegamos o superamos el total, no agregamos más
        if n_to_add == 0:
            continue

        puestos = [p for p in puestos_por_area.get(area, []) if p] or [
            "Operario", "Supervisor", "Técnico", "Ayudante"
        ]

        hombres = info_area.get("Hombres", 1)
        mujeres = info_area.get("Mujeres", 1)
        total = hombres + mujeres if (hombres + mujeres) > 0 else 1
        prob_m = hombres / total

        for _ in range(n_to_add):
            sex = "Hombre" if rnd.random() < prob_m else "Mujer"
            nombre = gen_nombre(sex)
            puesto = rnd.choice(puestos)

            max_id += 1
            row = {h: "" for h in headers}
            row["ID"] = str(max_id)
            row["Fecha"] = random_fecha(rnd, expired_ratio=args.expired_ratio)
            row["Area"] = area
            row["Puesto_de_Trabajo"] = puesto      # cargo correcto
            row["Nombre_Trabajador"] = nombre
            row["Sexo"] = sex
            row["Edad"] = str(rnd.randint(20, 60))
            row["Diestro_Zurdo"] = rnd.choice(["Diestro/a", "Zurdo/a"])
            prev = rnd.random() < 0.35
            row["Temporadas_previas"] = "SI" if prev else "NO"
            row["N_temporadas"] = str(rnd.randint(1, 6) if prev else 0)
            row["Tiempo_en_trabajo_meses"] = str(rnd.randint(1, 144))
            row["Actividad_previa"] = rnd.choice([
                "Agricultura", "Construcción", "Comercio", "Transporte", "Servicios",
                "Alimentos", "Metalurgia", "Pesca", ""
            ])
            row["Otra_actividad"] = rnd.choice(["NO", "SI", "NO", "NO"])
            row["Otra_actividad_cual"] = "" if row["Otra_actividad"] == "NO" else rnd.choice(
                ["Estudios", "Emprendimiento", "Hogar", "Deportes"]
            )

            fill_zone_answers(rnd, row)
            nuevas.append(row)

    df_new = pd.DataFrame(nuevas, columns=headers) if nuevas else pd.DataFrame(columns=headers)
    df_res = pd.concat([df_res, df_new], ignore_index=True)

    # Construir hoja Parametros
    if df_res.empty:
        df_params = pd.DataFrame(columns=["Área","Hombres","Mujeres","Total"])
    else:
        tmp = df_res.groupby(["Area","Sexo"]).size().reset_index(name="n")
        wid = tmp.pivot_table(index="Area", columns="Sexo", values="n", aggfunc="sum", fill_value=0).reset_index()
        if "Hombre" not in wid.columns:
            wid["Hombre"] = 0
        if "Mujer" not in wid.columns:
            wid["Mujer"] = 0
        wid["Total"] = wid["Hombre"] + wid["Mujer"]
        df_params = wid.rename(
            columns={"Area":"Área","Hombre":"Hombres","Mujer":"Mujeres"}
        )[["Área","Hombres","Mujeres","Total"]].sort_values("Área")

    # Guardar
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_res.to_excel(writer, sheet_name="Respuestas", index=False)
        df_params.to_excel(writer, sheet_name="Parametros", index=False)

    print(f"✔ Listo: agregado(s) {len(df_new)} registro(s) a '{out_path}'. Total ahora: {len(df_res)}")
    print("\nÁreas (Parametros):")
    print(df_params)

if __name__ == "__main__":
    main()
