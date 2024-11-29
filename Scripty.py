import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO
import base64
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

def create_pdf(dataframe, filename="tableau.pdf"):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    text_object = c.beginText(40, height - 40)
    text_object.setFont("Helvetica", 10)

    # Adding column headers
    text_object.textLine("Prénom et nom | Opérateur | Répétitions")
    text_object.textLine("----------------------------------------")

    # Adding table data
    for index, row in dataframe.iterrows():
        text_object.textLine(f"{row['Prénom et nom']} | {row['Opérateur']} | {row['Repetitions']}")
    
    c.drawText(text_object)
    c.showPage()
    c.save()

    buffer.seek(0)
    pdf_data = buffer.read()

    with open(filename, "wb") as f:
        f.write(pdf_data)

    return pdf_data

def convert_df_to_xlsx(df):
    # Convert the DataFrame to Excel format
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")

st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        col_prenom_nom = 'Prénom et nom'  # Directly using the column name "Prénom et nom"
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal['Opérateur'].unique()
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année", "Total"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        date_min = pd.to_datetime(df_principal[col_date]).min().date()
        date_max = pd.to_datetime(df_principal[col_date]).max().date()
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=date_min, max_value=date_max, value=date_max)
    
    if st.button("Analyser"):
        df_principal[col_date] = pd.to_datetime(df_principal[col_date])
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrer les données pour la période sélectionnée
        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]

        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        # Calcul des répétitions pour le tableau (toutes les dates)
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        # Ajouter la colonne 'Opérateur' au tableau
        repetitions_tableau = repetitions_tableau.merge(df_principal[['Prénom et nom', 'Opérateur']].drop_duplicates(), 
                                                         on='Prénom et nom', 
                                                         how='left')

        with col2:
            # Afficher le graphique avec les répétitions
            if periode_selectionnee != "Total":
                fig = px.bar(repetitions_graph, x=periode_selectionnee, y='Repetitions', color=col_prenom_nom, barmode='group',
                             title=f"Nombre de rapports d'intervention par {periode_selectionnee.lower()} pour les opérateurs sélectionnés (de {debut_periode} à {fin_periode})")
            else:
                fig = px.bar(repetitions_graph, x=col_prenom_nom, y='Repetitions',
                             title=f"Total des rapports d'intervention pour les opérateurs sélectionnés (de {debut_periode} à {fin_periode})")
            
            # Afficher les valeurs dans le graphique
            fig.update_traces(texttemplate='%{y}', textposition='outside')
            st.plotly_chart(fig)
        
        st.subheader(f"Tableau des répétitions par {periode_selectionnee.lower()} (de toutes les dates)")

        colonnes_affichage = [col_prenom_nom, periode_selectionnee, 'Repetitions', 'Opérateur'] if periode_selectionnee != "Total" else [col_prenom_nom, 'Repetitions', 'Opérateur']
        tableau_affichage = repetitions_tableau[colonnes_affichage]
        
        st.dataframe(tableau_affichage, use_container_width=True)
        
        # Tirage au sort
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                st.dataframe(lignes_tirees, use_container_width=True)
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")

        # Téléchargement des données sous format Excel
        xlsx_data = convert_df_to_xlsx(repetitions_tableau)
        st.download_button(label="Télécharger les répétitions sous format Excel", data=xlsx_data, file_name="repetitions.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # Téléchargement des données sous format PDF
        pdf_data = create_pdf(repetitions_tableau)
        st.download_button(label="Télécharger les répétitions sous format PDF", data=pdf_data, file_name="repetitions.pdf", mime="application/pdf")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
