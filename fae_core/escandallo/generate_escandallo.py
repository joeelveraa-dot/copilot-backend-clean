import os
import pandas as pd
from .extract_bom import extract_bom_from_csv
from .write_excel import write_bom_to_excel


def generate_escandallo(bom_file, materials_file, catalog_file):

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    template = os.path.join(BASE_DIR, "template.xlsm")

    os.makedirs("temp", exist_ok=True)

    bom_df = extract_bom_from_csv(bom_file, materials_file, catalog_file)

    # eliminar duplicados del BOM
    bom_df = bom_df.drop_duplicates(
        subset=["parent", "part_number"]
    )

    # ordenar jerárquicamente el BOM
    ordered = []

    def walk(parent):
        children = bom_df[bom_df["parent"] == parent]
        for _, row in children.iterrows():
            ordered.append(row)
            walk(row["part_number"])

    root = bom_df["parent"].iloc[0]
    walk(root)

    bom_df = pd.DataFrame(ordered).reset_index(drop=True)

    output_file = write_bom_to_excel(bom_df, template)

    return output_file