import io, re, zipfile, csv
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# -------------------- Helpers --------------------

def find_col_like(columns, needle):
    for c in columns:
        if needle.lower() in str(c).lower():
            return c
    return None

def find_any(columns, needles):
    for n in needles:
        c = find_col_like(columns, n)
        if c: return c
    return None

def norm_cuit(val):
    if pd.isna(val):
        return None
    s = re.sub(r"\D", "", str(val))
    return s or None


def cuit_valido(s: str) -> bool:
    d = re.sub(r"\D", "", str(s))
    if len(d) != 11:
        return False
    nums = [int(ch) for ch in d]
    pesos = [5,4,3,2,7,6,5,4,3,2]
    total = sum(p * nums[i] for i, p in enumerate(pesos))
    resto = total % 11
    dv = 11 - resto
    if dv == 11: dv = 0
    elif dv == 10: dv = 9
    return nums[-1] == dv

def parse_cod_comp(tipo_val):
    try:
        s = str(tipo_val or "")
        num = s.split("-")[0].strip()
        return int(num)
    except Exception:
        return None

def parse_letra_from_tipo(tipo_val):
    s = str(tipo_val or "")
    for suf in [" A"," B"," C"," E"]:
        if s.endswith(suf): return suf.strip()
    for L in ["A","B","C","E"]:
        if f" {L}" in s: return L
    return ""

def map_tipo_comp_default(cod):
    if cod in [1,6,11]:  return "FC"
    if cod in [2,7,12]:  return "ND"
    if cod in [3,8,13]:  return "NC"
    return ""

def override_tipo_letra(cod):
    if cod == 19:   return ("FC", "E")
    if cod == 21:   return ("NC", "E")
    if cod == 201:  return ("FC", "A")
    if cod == 202:  return ("ND", "A")
    if cod == 203:  return ("NC", "A")
    return (None, None)

def map_condicion(letra):
    if letra == "A": return "RI"
    if letra == "C": return "MT"
    return "CF"

def fnum(x):
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return 0.0
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            return float(x)
        s = str(x).strip()
        if s == "" or s.lower() == "nan":
            return 0.0
        s = s.replace("\xa0", "").replace(" ", "").replace("$", "").replace("ARS", "").replace("ars", "")
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        has_dot = "." in s
        has_comma = "," in s
        if has_dot and has_comma:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
            return float(s)
        if has_comma:
            frac = s.split(",")[-1]
            if frac.isdigit() and 1 <= len(frac) <= 3:
                s = s.replace(".", "")
                s = s.replace(",", ".")
                return float(s)
            s = s.replace(",", "").replace(".", "")
            return float(s)
        if has_dot:
            frac = s.split(".")[-1]
            if frac.isdigit() and 1 <= len(frac) <= 3:
                s = s.replace(",", "")
                return float(s)
            s = s.replace(".", "")
            return float(s)
        s2 = re.sub(r"[^\d\-]", "", s)
        if s2 in ("", "-", "+"):
            return 0.0
        return float(s2)
    except Exception:
        return 0.0

def sanitize_name(name):
    if not name: return "EMPRESA"
    s = str(name)
    s = re.sub(r"[^\w\s\-\.\(\)]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s or "EMPRESA")[:64]

# --------- CUIT / Empresa extractors (file-like) ---------

def extract_cuit_from_text(text: str):
    if not text:
        return None
    # 1) Buscar 'CUIT ...' permitiendo separadores - _ espacio .
    m = re.search(r"CUIT[^0-9]*([0-9][0-9\-\s_\.]{9,}[0-9])", text, re.IGNORECASE)
    candidates = []
    if m:
        candidates.append(m.group(1))
    # 2) Cualquier bloque con 11 dígitos y separadores, o 11 seguidos
    candidates += [mm.group(1) for mm in re.finditer(r"([0-9][0-9\-\s_\.]{9,}[0-9])", text)]
    candidates += [mm.group(0) for mm in re.finditer(r"(\d{11})", text)]
    for c in candidates:
        d = re.sub(r"\D", "", c)
        if len(d) == 11 and cuit_valido(d):
            return d
    return None
    m = re.search(r"CUIT[^0-9]*([0-9\-]+)", text, re.IGNORECASE)
    candidate = None
    if m:
        candidate = re.sub(r"\D","", m.group(1))
    else:
        m2 = re.search(r"(\d{2}-\d{8}-\d|\d{11})", text)
        if m2:
            candidate = re.sub(r"\D","", m2.group(1))
    if candidate and len(candidate) == 11:
        return candidate
    return None

def extract_empresa_from_text(text: str):
    if not text: return None
    m = re.search(r"(Denominaci[oó]n|Raz[oó]n Social|Razon Social)\s*:\s*(.+)", text, re.IGNORECASE)
    if m:
        val = m.group(2).strip()
        val = re.split(r"CUIT\b", val, flags=re.IGNORECASE)[0].strip(" -:")
        return val if val else None
    return None

def extract_cuit_from_afip_file(file, filename):
    # Intentar PRIMERO por nombre de archivo
    if filename:
        from pathlib import Path as _P
        c_by_name = extract_cuit_from_text(_P(filename).name)
        if c_by_name:
            return c_by_name
    try:
        if filename.lower().endswith((".xlsx",".xlsm",".xls")):
            file.seek(0)
            top = pd.read_excel(file, sheet_name=0, header=None, nrows=8, dtype=str)
            for val in top.fillna("").astype(str).values.flatten():
                c = extract_cuit_from_text(val)
                if c: return c
        else:
            file.seek(0)
            head = file.read(4096).decode("utf-8", errors="ignore")
            for line in head.splitlines()[:8]:
                c = extract_cuit_from_text(line)
                if c: return c
    except Exception:
        pass
    c = extract_cuit_from_text(filename)
    if c: return c
    return "SINCUIT"

def extract_cuit_from_datos_file(file, filename):
    # Intentar PRIMERO por nombre de archivo
    if filename:
        from pathlib import Path as _P
        c_by_name = extract_cuit_from_text(_P(filename).name)
        if c_by_name:
            return c_by_name
    # 1) A1 de la primera hoja
    try:
        file.seek(0)
        top = pd.read_excel(file, sheet_name=0, header=None, nrows=10, dtype=str)
        for val in top.fillna("").astype(str).values.flatten():
            c = extract_cuit_from_text(val)
            if c: 
                return c
    except Exception:
        pass
    # 2) nombre de archivo
    c = extract_cuit_from_text(filename)
    if c: return c
    return None

# -------------------- Readers --------------------

import csv
def sniff_sep(sample: bytes):
    try:
        s = sample.decode("utf-8", errors="ignore")
        dialect = csv.Sniffer().sniff(s, delimiters=[",",";","\t","|"])
        return dialect.delimiter
    except Exception:
        text = s if 's' in locals() else ""
        counts = {",": text.count(","), ";": text.count(";"), "\t": text.count("\t"), "|": text.count("|")}
        delim = max(counts, key=counts.get)
        return delim if counts[delim] > 0 else ","

def read_afip(file, filename) -> pd.DataFrame:
    if filename.lower().endswith((".xlsx",".xlsm",".xls")):
        file.seek(0)
        try:
            df = pd.read_excel(file, sheet_name=0, header=1, dtype=str)
        except Exception:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=0, header=0, dtype=str)
    else:
        file.seek(0)
        sample = file.read(4096)
        sep = sniff_sep(sample)
        for header in (1, 0):
            try:
                file.seek(0)
                df = pd.read_csv(file, header=header, dtype=str, sep=sep)
                if df.shape[1] == 1 and sep != ";":
                    file.seek(0)
                    df = pd.read_csv(file, header=header, dtype=str, sep=";")
                break
            except Exception:
                if header == 0:
                    raise
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_datos(file) -> pd.DataFrame:
    file.seek(0)
    xls = pd.ExcelFile(file)
    datos_sheet = None
    for s in xls.sheet_names:
        if s.strip().lower() == "datos":
            datos_sheet = s
            break
    if datos_sheet is None:
        raise RuntimeError("No se encontró una hoja llamada 'datos'.")
    df = pd.read_excel(xls, sheet_name=datos_sheet)
    upper = {c: str(c).strip().upper() for c in df.columns}
    df.rename(columns=upper, inplace=True)
    if "CUIT" not in df.columns:
        raise RuntimeError("La hoja 'datos' debe tener una columna 'CUIT'.")
    if "CODIGO" not in df.columns:
        df["CODIGO"] = None
    if "CONCEPTO" not in df.columns:
        df["CONCEPTO"] = None
    if "PROVEEDOR" not in df.columns and "CLIENTE" not in df.columns:
        df["PROVEEDOR"] = None
    return df

# -------------------- PV --------------------

def extract_pv(row, c_pv, c_numero):
    import re as _re
    def de_zero(x):
        s = _re.sub(r"\D", "", str(x or ""))
        if s == "":
            return ""
        try:
            return str(int(s))
        except Exception:
            return s

    if c_pv:
        val = row.get(c_pv)
        if val is not None and str(val).strip() != "":
            return de_zero(val)

    if c_numero:
        s = str(row.get(c_numero) or "").strip()
        if "-" in s:
            left = s.split("-")[0].strip()
            return de_zero(left)
        return de_zero(s[:5])

    return ""

# -------------------- Core transform --------------------

def procesar(afip_df: pd.DataFrame, datos_df: pd.DataFrame):
    cols = afip_df.columns
    c_fecha     = find_any(cols, ["Fecha"])
    c_tipo      = find_any(cols, ["Tipo"])
    c_numero    = find_any(cols, ["Número","Numero"])
    c_num_desde = find_any(cols, ["Número Desde","Numero Desde"])
    c_num_hasta = find_any(cols, ["Número Hasta","Numero Hasta"])
    c_pv_col    = find_any(cols, ["Punto de Venta","Pto Vta","Pto. Vta","P.V."])
    c_aut       = find_any(cols, ["Cód. Autorización","CAE"])
    c_tdoc      = find_any(cols, ["Tipo Doc.","Tipo Doc","Tipo Documento"])
    c_ndoc      = find_any(cols, ["Nro. Doc.","Número Doc","N° Doc"])
    c_den       = find_any(cols, ["Denominación","Razón Social","Razon Social"])
    c_tc        = find_any(cols, ["Tipo Cambio","Cotización"])
    c_mon       = find_any(cols, ["Moneda"])
    c_trib      = find_any(cols, ["Otros Tributos"])
    c_no_grav   = find_any(cols, ["Neto No Gravado"])
    c_exen      = find_any(cols, ["Op. Exentas"])
    c_total     = find_any(cols, ["Imp. Total","Importe Total","Total Comprobante"])

    def col_grav(label): return find_any(cols, [f"Neto Grav. IVA {label}"])
    def col_iva(label):  return find_any(cols, [f"IVA {label}"])
    aliq_labels = ["2,5%", "5%", "10,5%", "21%", "27%"]
    aliq_cols = [(lab, col_grav(lab), col_iva(lab)) for lab in aliq_labels]

    datos_df = datos_df.copy()
    datos_df["CUIT_N"] = datos_df["CUIT"].apply(norm_cuit)
    map_cod = dict(zip(datos_df["CUIT_N"], datos_df["CODIGO"]))
    map_con = dict(zip(datos_df["CUIT_N"], datos_df["CONCEPTO"]))

    rows = []
    for _, r in afip_df.iterrows():
        cod = parse_cod_comp(r.get(c_tipo))
        tipo_override, letra_override = override_tipo_letra(cod) if cod is not None else (None, None)
        letra = letra_override or parse_letra_from_tipo(r.get(c_tipo))
        tipo_comp = tipo_override or map_tipo_comp_default(cod)
        condicion = map_condicion(letra)
        pv = extract_pv(r, c_pv_col, c_numero)
        cuit_norm = norm_cuit(r.get(c_ndoc))

        cod_prov_cli = map_cod.get(cuit_norm)
        concepto     = map_con.get(cuit_norm)

        partes = []
        for lab, net_c, iva_c in aliq_cols:
            neto = fnum(r.get(net_c, 0)) if net_c else 0.0
            iva  = fnum(r.get(iva_c, 0)) if iva_c else 0.0
            if (net_c and abs(neto) > 1e-9) or (iva_c and abs(iva) > 1e-9):
                tasa = float(lab.replace("%","").replace(",", "."))
                partes.append({"label": lab, "tasa": tasa, "neto": neto, "iva": iva})

        val_no_grav = fnum(r.get(c_no_grav, 0)) if c_no_grav else 0.0
        val_exen    = fnum(r.get(c_exen, 0)) if c_exen else 0.0
        val_otros   = fnum(r.get(c_trib, 0)) if c_trib else 0.0
        val_total   = fnum(r.get(c_total, 0)) if c_total else 0.0

        comunes = {
            "Fecha": r.get(c_fecha),
            "Cod Comp": cod,
            "Tipo de Comp": tipo_comp,
            "Letra": letra,
            "Tipo": r.get(c_tipo),
            "Punto de Venta": pv,
            "Número Desde": r.get(c_num_desde),
            "Número Hasta": r.get(c_num_hasta),
            "Cód. Autorización": r.get(c_aut),
            "Tipo Doc.": r.get(c_tdoc),
            "Nro. Doc.": r.get(c_ndoc),
            "Denominación": r.get(c_den),
            "Condicion": map_condicion(letra),
            "Cod Prov/Cliente": cod_prov_cli,
            "Concepto": concepto,
            "Tipo Cambio": r.get(c_tc),
            "Moneda": r.get(c_mon),
            "Imp. Neto No Gravado": 0.0,
            "Imp. Op. Exentas": 0.0,
            "Otros Tributos": 0.0,
        }

        if len(partes) == 0:
            out = comunes.copy()
            out["IVA"] = 0.0
            out["Tasa IVA"] = 0.0
            suma_otros = (val_no_grav or 0.0) + (val_exen or 0.0) + (val_otros or 0.0)
            base_grav = max(0.0, (val_total or 0.0) - suma_otros)
            out["Imp. Neto Gravado"]    = round(base_grav, 2)
            out["Imp. Neto No Gravado"] = round(val_no_grav, 2)
            out["Imp. Op. Exentas"]     = round(val_exen, 2)
            out["Otros Tributos"]       = round(val_otros, 2)
            out["Imp. Total"]           = round(base_grav + suma_otros, 2)
            rows.append(out)
            continue

        partes = sorted(partes, key=lambda x: x["tasa"])
        for idx_part, p in enumerate(partes):
            out = comunes.copy()
            out["Imp. Neto Gravado"] = round(p["neto"], 2)
            out["IVA"] = round(p["iva"], 2)
            out["Tasa IVA"] = p["tasa"]
            if idx_part == 0:
                out["Imp. Neto No Gravado"] = round(val_no_grav, 2)
                out["Imp. Op. Exentas"] = round(val_exen, 2)
                out["Otros Tributos"] = round(val_otros, 2)
                out["Imp. Total"] = round(p["neto"] + p["iva"] + val_no_grav + val_exen + val_otros, 2)
            else:
                out["Imp. Total"] = round(p["neto"] + p["iva"], 2)
            rows.append(out)

    result = pd.DataFrame(rows)

    out_cols = [
        "Fecha","Cod Comp","Tipo de Comp","Letra","Tipo",
        "Punto de Venta","Número Desde","Número Hasta","Cód. Autorización",
        "Tipo Doc.","Nro. Doc.","Denominación",
        "Condicion","Cod Prov/Cliente","Concepto","Tipo Cambio","Moneda",
        "Imp. Neto Gravado","Imp. Neto No Gravado","Imp. Op. Exentas",
        "Otros Tributos","IVA","Imp. Total","Tasa IVA"
    ]
    for c in out_cols:
        if c not in result.columns:
            result[c] = np.nan
    result = result[out_cols]
    return result

# -------------------- UI --------------------

st.set_page_config(page_title="AFIP2Import – Cloud", page_icon="🧾", layout="centered")

st.title("AFIP2Import – Cloud")
st.caption("Subí AFIP/ARCA y el archivo de **datos** (hoja `datos`) y te devuelvo el Excel para importar.")

tab1, tab2 = st.tabs(["Individual", "Batch (multi-archivo)"])

with tab1:
    afip = st.file_uploader("Archivo AFIP/ARCA", type=["xlsx","xls","csv"], key="afip1")
    datos = st.file_uploader("Archivo de datos (hoja 'datos')", type=["xlsx","xls"], key="datos1")
    if afip and datos:
        if st.button("Procesar"):
            try:
                afip_df = read_afip(afip, afip.name)
                datos_df = read_datos(datos)
                out = procesar(afip_df, datos_df)
                datos_stem = Path(datos.name).stem
                bio = io.BytesIO()
                try:
                    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                        out.to_excel(writer, index=False)
                except ModuleNotFoundError:
                    st.error("Falta 'openpyxl' en el entorno. Agregalo a requirements.txt (openpyxl==3.1.2) y redeploy.")
                    st.stop()
                bio.seek(0)
                fname = f"importar {datos_stem}.xlsx"
                st.download_button("⬇️ Descargar", data=bio.getvalue(), file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success(f"Listo: {fname}")
            except Exception as e:
                st.error(str(e))

with tab2:
    st.write("👉 **Cargá varios AFIP** y **varios 'datos' por empresa**. La app genera **un archivo por AFIP**, con nombre = `importar <datos>`.")
    afips = st.file_uploader("Varios AFIP/ARCA", type=["xlsx","xls","csv"], accept_multiple_files=True, key="afip2")
    datos_files = st.file_uploader("Varios 'datos' por empresa (obligatorio, A1 con CUIT o en el nombre)", type=["xlsx","xls"], accept_multiple_files=True, key="datos2")

    # Mostrar CUIT detectados
    datos_map = {}
    datos_name = {}
    if datos_files:
        st.subheader("CUIT en 'datos'")
        for df_file in datos_files:
            try:
                c = extract_cuit_from_datos_file(df_file, df_file.name)
            except Exception as e:
                c = None
            st.write(f"- {df_file.name} → CUIT: **{c or 'NO DETECTADO'}**")
            if c:
                try:
                    df = read_datos(df_file)
                    datos_map[c] = df
                    datos_name[c] = Path(df_file.name).stem
                except Exception as e:
                    st.error(f"{df_file.name}: {e}")

    if afips:
        st.subheader("CUIT detectados en AFIP")
        for f in afips:
            try:
                cuit = extract_cuit_from_afip_file(f, f.name)
            except Exception:
                cuit = "SINCUIT"
            st.write(f"- {f.name} → CUIT: **{cuit}**")

    if afips and datos_map and st.button("Procesar lote"):
        try:
            zipbio = io.BytesIO()
            with zipfile.ZipFile(zipbio, "w", zipfile.ZIP_DEFLATED) as z:
                used_names = set()
                for f in afips:
                    try:
                        cuit = extract_cuit_from_afip_file(f, f.name)
                        df_afip = read_afip(f, f.name)
                        datos_df = datos_map.get(cuit)
                        if datos_df is None:
                            st.warning(f"Sin 'datos' para CUIT {cuit} → {f.name} saltado.")
                            continue
                        out = procesar(df_afip, datos_df)
                        datos_stem = datos_name.get(cuit, f"{cuit}")
                        bio = io.BytesIO()
                        try:
                            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                                out.to_excel(writer, index=False)
                        except ModuleNotFoundError:
                            st.error("Falta 'openpyxl' en el entorno. Agregalo a requirements.txt (openpyxl==3.1.2) y redeploy.")
                            st.stop()
                        bio.seek(0)
                        fname = f"importar {datos_stem}.xlsx"
                        base = fname
                        i = 2
                        while fname in used_names:
                            fname = f"{Path(base).stem} ({i}).xlsx"
                            i += 1
                        used_names.add(fname)
                        z.writestr(fname, bio.getvalue())
                    except Exception as ie:
                        z.writestr(f"ERROR_{Path(f.name).stem}.txt", str(ie))
            zipbio.seek(0)
            st.download_button("⬇️ Descargar ZIP", data=zipbio.getvalue(), file_name="AFIP2Import_resultados.zip", mime="application/zip")
            st.success("Listo. (Un archivo por AFIP; ZIP sin carpetas; nombre = 'importar <datos>'.)")
        except Exception as e:
            st.error(str(e))

st.info("El CUIT se detecta del **nombre del archivo** (si existe) o de **A1** de la primera hoja.")
