# -*- coding: utf-8 -*-
"""
Created on Fri Jul  4 10:53:41 2025

@author: MijkePietersma
"""

import streamlit as st
import numpy as np
import pandas as pd

st.title("FiscFree / Hellorider Analysis")

# 1. User uploads only the FiscFree Excel file
fiscfree_file = st.file_uploader("Upload FiscFree Excel", type=["xlsx"])

# 2. Load the other reference files from disk (bundled in the app)
hellorider = pd.read_excel("data/20250627 - Hellorider - Export.xlsx")
bike_totaal = pd.read_excel("data/20250630 - DRG Dealers - Overzicht.xlsx", skiprows=6)
mail_fiscfree = pd.read_excel("data/mail_fiscfree.xlsx")

if fiscfree_file:

    # Read the uploaded files into dataframes
    fiscfree = pd.read_excel(fiscfree_file)

    # Add a button to run analysis
    if st.button("Run Analysis"):
        # 2. Fast 1‑to‑1 match (vectorised):   Artikelnr ↔ Ean Code
        # ---------------------------------------------------------------------
        # keep the FIRST occurrence of each Ean Code  ➜ guarantees a 1‑to‑1 join
        lookup_cols = ["Ean Code", "Brand", "Msrp Ex Vat", "Name"]
        lookup = (
            hellorider[lookup_cols]
                .drop_duplicates(subset="Ean Code")          # <-- new line
                .rename(columns={
                    "Ean Code":    "Artikelnr",
                    "Brand":       "Brand_hr",
                    "Msrp Ex Vat": "adviesprijs",
                    "Name":        "Naam Hellorider",
                })
        )

        # now the merge is “many‑to‑ONE”; validate='m:1' makes pandas raise
        # if duplicates ever sneak back in.
        fiscfree = fiscfree.merge(
            lookup,
            how="left",
            on="Artikelnr",
            validate="m:1"     # many‑to‑one
        )

        brand_match = fiscfree["Brand_hr"].str.lower() == fiscfree["Merk"].str.lower()
        fiscfree.loc[~brand_match, ["adviesprijs", "Naam Hellorider"]] = np.nan
        fiscfree["Artikelnummer check"] = brand_match


        # ---------------------------------------------------------------------
        # 3. Fallback fuzzy match (Type contained in Name, Brand == Merk)
        #    – only for rows whose price is still NA
        # ---------------------------------------------------------------------
        
        need = fiscfree["adviesprijs"].isna()
        
        type_low  = fiscfree["Type"].str.lower().str.replace(" ", "", regex=False)
        merk_low  = fiscfree["Merk"].str.lower()
        name_low  = hellorider["Name"].str.lower().str.replace(" ", "", regex=False)
        brand_low = hellorider["Brand"].str.lower()
        
        progress_bar = st.progress(0)
        total = need.sum()
        
        for count, i in enumerate(fiscfree[need].index, start=1):
            candidates = hellorider[
                brand_low.eq(merk_low[i]) &
                name_low.str.contains(type_low[i], regex=False)
            ]
            if not candidates.empty:
                fiscfree.at[i, "adviesprijs"]     = candidates["Msrp Ex Vat"].iloc[0]
                fiscfree.at[i, "Naam Hellorider"] = candidates["Name"].iloc[0]
            
            progress_bar.progress(count / total)

        # ---------------------------------------------------------------------
        # 4. Price incl. VAT and period tagging
        # ---------------------------------------------------------------------
        fiscfree["adviesprijs"] *= 1.21
        fiscfree["adviesprijs"]  = fiscfree["adviesprijs"].round(2)

        fiscfree["Besteldatum"] = pd.to_datetime(fiscfree["Besteldatum"])
        fiscfree["periode"]     = pd.NA

        conditions = [
            fiscfree["Besteldatum"].between("2025-01-01", "2025-04-01"),
            fiscfree["Besteldatum"] >= "2025-04-02",
            fiscfree["Besteldatum"].between("2024-01-01", "2024-12-31"),
        ]
        choices = [
            "Tussen 1-1-2025 en 2-4-2025",
            "Vanaf 2-4-2025",
            "2024",
        ]
        fiscfree["periode"] = np.select(conditions, choices, default=pd.NA)

        # ---------------------------------------------------------------------
        # 5. Δ‑marge berekeningen
        # ---------------------------------------------------------------------
        fiscfree["delta"] = np.where(
            fiscfree["adviesprijs"].notna(),
            0.10 * (fiscfree["adviesprijs"] - fiscfree["bedraghoofdproductincl"]),
            np.nan  # or 0, if you prefer defaulting to zero
        )
        fiscfree["Diff >15%"] = (
            fiscfree["bedraghoofdproductincl"] < 0.85 * fiscfree["adviesprijs"]
        )
        fiscfree["Marge delta >15%"] = (
            (fiscfree["adviesprijs"] - fiscfree["bedraghoofdproductincl"]) * 0.10
        ).where(fiscfree["Diff >15%"])
        fiscfree["max_budget"] = fiscfree[[
            "maximaalteverrekenenhoofdproduct",
            "bestelling.verrekeninghoofdproductbedrag"
        ]].max(axis=1, skipna=True)

        fiscfree["bedraghoofd = max_budget"]  = np.isclose(
            fiscfree["bedraghoofdproductincl"].round(2),
            fiscfree["max_budget"].round(2),
        )
        fiscfree["Controle twee condities"] = fiscfree["bedraghoofd = max_budget"] | fiscfree["Diff >15%"]

        # sums & counts per period
        # ---------------------------------------------------------------------
        misgelopen_df = (
            fiscfree
              .groupby("periode")
              .agg(
                  totaal_misgelopen_marge          = ("Marge delta >15%", "sum"),
                  totaal_bestellingen              = ("periode", "size"),
                  aantal_bestellingen_max_budget_gelijk = ("bedraghoofd = max_budget", "sum"),
                  aantal_bestellingen_max_budget_ongelijk = (
                      "bedraghoofd = max_budget", 
                      lambda s: (~s).sum()  # count where False
                  ),
                  aantal_bestellingen_grote_delta = ("Marge delta >15%", "count"),
                  aantal_bestellingen_kleine_delta = (
                      "Marge delta >15%",
                      lambda s: s.isna().sum() + (s <= 0).sum()  # count where <=0 or NaN
                  ),
              )
              .reset_index()
        )


        # ---------------------------------------------------------------------
        misgelopen_df_leverancier = (
            fiscfree
              .groupby("Leveranciervestiging")
              .agg(
                  totaal_misgelopen_marge          = ("Marge delta >15%", "sum"),
                  totaal_bestellingen              = ("periode", "size"),
                  aantal_bestellingen_max_budget_gelijk = ("bedraghoofd = max_budget", "sum"),
                  aantal_bestellingen_max_budget_ongelijk = (
                      "bedraghoofd = max_budget", 
                      lambda s: (~s).sum()  # count where False
                  ),
                  aantal_bestellingen_grote_delta = ("Marge delta >15%", "count"),
                  aantal_bestellingen_kleine_delta = (
                      "Marge delta >15%",
                      lambda s: s.isna().sum() + (s <= 0).sum()  # count where <=0 or NaN
                  ),
              )
              .reset_index()
        )

        # ---------------------------------------------------------------------
        # 6. Marge per leverancier (vanaf 2‑4‑2025 & delta > 1)
        # ---------------------------------------------------------------------
        marge_per_leverancier = (
            fiscfree[
                (fiscfree["periode"] == "Vanaf 2-4-2025") &
                fiscfree["Marge delta >15%"].gt(0)  # checks if > 0 and not null
            ]
            .loc[:, [
                "Leveranciervestiging", "Bestelnummer", "adviesprijs", "bedraghoofdproductincl",
                "Merk", "Type", "Naam Hellorider", "Besteldatum"
            ]]
            .sort_values("Leveranciervestiging")
        )


        # ---------------------------------------------------------------------
        # 7. Fraude‑analyse: exacte match verkoopprijs == max budget
        # ---------------------------------------------------------------------

        bestelling_fraude = fiscfree[
            (fiscfree["periode"] == "Vanaf 2-4-2025") &
            fiscfree["bedraghoofd = max_budget"]
        ].loc[:, ["Leveranciervestiging","Bestelnummer","bedraghoofdproductincl",
                  "max_budget","adviesprijs","Merk","Type","Besteldatum"]]
        bestelling_fraude = bestelling_fraude.sort_values("Leveranciervestiging")

        # ---------------------------------------------------------------------
        # 8. Bike‑Totaal ‘Formule’ veld koppelen
        # ---------------------------------------------------------------------

        bike_totaal["email_clean"]       = bike_totaal["E mail"].str.strip().str.lower()
        mail_fiscfree["email_clean"]     = (
            mail_fiscfree["leverancier_vestiging_email"].str.strip().str.lower()
        )

        # (optional) drop rows whose e‑mail is missing
        bike_totaal  = bike_totaal.dropna(subset=["email_clean"])
        mail_fiscfree = mail_fiscfree.dropna(subset=["email_clean"])

        # ------------------------------------------------------------------
        # 2. Inner‑join on the cleaned address – duplicates are handled automatically
        # ------------------------------------------------------------------
        bike_tot_df = (
            mail_fiscfree[["email_clean", "leverancier_vestiging_naam"]]
              .merge(
                  bike_totaal[["email_clean", "Formule"]],
                  on="email_clean",
                  how="left"          # keep only rows that exist in *both* table        # many‑to‑one: raises if the same e‑mail
                                        # is paired to *different* Formules
              )
              .rename(columns={"leverancier_vestiging_naam": "Naam"})
              .dropna(subset=["Naam"])
              .drop_duplicates(subset=["Naam", "Formule"])  # optional cleanup
              .reset_index(drop=True)
        )

        # bike_tot_df now has perfectly aligned 'Naam' and 'Formule' columns


        def add_formule(df, key="Leveranciervestiging"):
            return df.merge(
                bike_tot_df.rename(columns={"Naam": key}),
                how="left", on=key
            ).assign(Formule = lambda d: d["Formule"].fillna("N.v.t."))

        misgelopen_df_leverancier = add_formule(misgelopen_df_leverancier)
        marge_per_leverancier = add_formule(marge_per_leverancier)
        bestelling_fraude     = add_formule(bestelling_fraude)

        # re‑order columns to match the R output
        keep_first = ["Leveranciervestiging", "Formule"]
        misgelopen_df_leverancier          = misgelopen_df_leverancier[keep_first + [c for c in misgelopen_df_leverancier.columns if c not in keep_first]]
        marge_per_leverancier = marge_per_leverancier[keep_first + [c for c in marge_per_leverancier.columns if c not in keep_first]]
        bestelling_fraude     = bestelling_fraude[keep_first + [c for c in bestelling_fraude.columns if c not in keep_first]]

        # Optional: allow user to download results as Excel
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            misgelopen_df.to_excel(writer, sheet_name="Misgelopen marge per periode", index=False)
            misgelopen_df_leverancier.to_excel(writer, sheet_name="Misgelopen marge per leverancier", index=False)
            marge_per_leverancier.to_excel(writer, sheet_name="Bestellingen verschil >15%")
            bestelling_fraude.to_excel(writer, sheet_name="Bestellingen met fraude")
            fiscfree.to_excel(writer, sheet_name="Alle data Fiscfree", index=False)
            # Add other sheets as needed
        output.seek(0)

        st.download_button(
            label="Download results as Excel",
            data=output,
            file_name="FiscFree_misgelopen_marge_analyse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload all required Excel files to run the analysis.")
