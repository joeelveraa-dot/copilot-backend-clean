import pandas as pd
import re


def read_csv_mapex(path):
    """
    Lee CSV exportado de Mapex de forma robusta
    """

    df = pd.read_csv(
        path,
        sep=None,          # autodetecta ; o ,
        engine="python",
        encoding="latin-1",
        dtype=str
    )

    # limpiar nombres de columnas
    df.columns = df.columns.astype(str).str.strip()

    # si todo quedó en una sola columna
    if len(df.columns) == 1:
        df = df[df.columns[0]].str.split(";", expand=True)
        df.columns = df.iloc[0]
        df = df[1:]
        df.columns = df.columns.astype(str).str.strip()

    return df


def clean_description(desc):
    """
    Convierte descripciones largas de Mapex en
    descripciones cortas como el Excel original
    """

    if not isinstance(desc, str):
        return desc

    desc = desc.upper()

    replacements = {
        "CAJA": "BOX",
        "PLACA": "PCB",
        "SENSOR": "SENSOR",
        "IMAN": "MAGNET",
        "EPORAL": "RESIN",
        "CUERPO EXTERIOR": "OUTER HOUSING",
        "GUIA": "HOUSING",
        "POLIAMIDA": "POLIAMIDE",
        "ESTAÑO": "TIN",
        "TIRA DE PINES": "PCB TERMINALS"
    }

    for k, v in replacements.items():
        if k in desc:
            return v

    return desc


def extract_bom_from_csv(bom_file, materials_file, catalog_file):

    # -----------------------------
    # LEER BOM
    # -----------------------------
    df = read_csv_mapex(bom_file)

    # usar solo BOM actual
    if "Tipo" in df.columns:
        df = df[df["Tipo"].astype(str).str.strip().str.lower() == "actual"]

    # -----------------------------
    # DETECTAR NIVELES n0-n9
    # -----------------------------
    levels = []

    for col in df.columns:
        name = col.strip().lower()

        if re.match(r"^n\d+$", name):
            levels.append(col)

    levels = sorted(levels, key=lambda x: int(x.strip()[1:]))

    if not levels:
        print("COLUMNAS DETECTADAS:", df.columns.tolist())
        raise ValueError("No se encontraron columnas n0-n9 en el CSV")

    # -----------------------------
    # RECONSTRUIR BOM
    # -----------------------------
    stack = {}
    bom = []

    for _, row in df.iterrows():

        for level_index, level in enumerate(levels):

            value = row.get(level)

            if pd.isna(value) or str(value).strip() == "":
                continue

            part = str(value).strip()

            # limpiar niveles inferiores
            stack = {k: v for k, v in stack.items() if k < level_index}
            stack[level_index] = part

            if level_index == 0:
                continue

            parent = stack.get(level_index - 1)

            if not parent:
                continue

            quantity = row.get("Coeficiente", "1")

            try:
                quantity = float(str(quantity).replace(",", "."))
            except:
                quantity = 1.0

            bom.append({
                "parent": parent,
                "part_number": part,
                "quantity": quantity
            })

    bom_df = pd.DataFrame(bom)

    if bom_df.empty:
        raise ValueError("No se pudo extraer el BOM del CSV")

    bom_df = bom_df.drop_duplicates(
        subset=["parent", "part_number"]
    ).reset_index(drop=True)

    # -----------------------------
    # LEER CSV DE MATERIALES
    # -----------------------------
    materials = read_csv_mapex(materials_file)

    part_col = "Componente"
    desc_col = "Descripcion"
    type_col = "Tipo Comp"

    materials = materials.rename(columns={
        part_col: "part_number",
        desc_col: "Item Description",
        type_col: "Material type"
    })

    materials["part_number"] = materials["part_number"].astype(str).str.strip()
    bom_df["part_number"] = bom_df["part_number"].astype(str).str.strip()

    bom_df = bom_df.merge(
        materials[["part_number", "Item Description", "Material type"]],
        on="part_number",
        how="left"
    )

    # -----------------------------
    # LEER CATALOGO DE ARTICULOS
    # -----------------------------
    catalog = read_csv_mapex(catalog_file)

    # limpiar nombres de columnas
    catalog.columns = (
        catalog.columns
        .str.replace('"', '')
        .str.strip()
    )

    # detectar columna de referencia automáticamente
    part_col = None

    for col in catalog.columns:
        if "ident" in col.lower():
            part_col = col
            break

    if part_col is None:
        print("COLUMNAS CATALOGO:", catalog.columns.tolist())
        raise ValueError("No se pudo detectar la columna de referencia del catálogo")

    catalog = catalog.rename(columns={
        part_col: "part_number",
        "DESCRIPCION": "Item Description"
    })

    catalog["part_number"] = catalog["part_number"].astype(str).str.strip()

    bom_df = bom_df.merge(
        catalog[["part_number", "Item Description"]],
        on="part_number",
        how="left",
        suffixes=("", "_catalog")
    )

    bom_df["Item Description"] = bom_df["Item Description"].fillna(
        bom_df["Item Description_catalog"]
    )

    bom_df = bom_df.drop(columns=["Item Description_catalog"])

    # -----------------------------
    # LIMPIAR DESCRIPCIONES
    # -----------------------------
    bom_df["Item Description"] = bom_df["Item Description"].apply(clean_description)

    # -----------------------------
    # MAPEO TIPO MATERIAL
    # -----------------------------
    material_map = {
        "Componente": "Purchased component",
        "Subconjunto": "Purchased component",
        "BULK": "Raw material",
        "Acabado Bulk": "Raw material"
    }

    bom_df["Material type"] = bom_df["Material type"].map(material_map)

    print("\n===== BOM GENERADO =====")
    print(bom_df)

    return bom_df