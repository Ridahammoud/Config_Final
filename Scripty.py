import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

# Fonction pour charger les données
@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

# Fonction pour exporter en format Excel
def to_xlsx(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    processed_data = output.getvalue()
    return processed_data

# Fonction pour exporter en format PDF
def to_pdf(df, filename="repetitions.pdf"):
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=letter)
    textobject = c.beginText(40, 750)
    textobject.setFont("Helvetica", 10)

    for i, row in df.iterrows():
        textobject.textLine(f"{row['Prénom et nom']} - {row['Repetitions']}")  # Exemple, ajustez selon vos colonnes
    c.drawText(textobject)
    c.showPage()
    c.save()
    processed_data = output.getvalue()
    return processed_data

# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")

st.title("📊 Analyse des interventions des opérateurs")

# Téléchargement du fichier
fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        col_prenom_nom = 'Prénom et nom'  # Définir directement la colonne 'Prénom et nom'
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique()
        select_all_operators = st.checkbox("Sélectionner tous les opérateurs")
        operateurs_selectionnes = operateurs if select_all_operators else st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année", "Total"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        date_min = pd.to_datetime(df_principal[col_date]).min().date()
        date_max = pd.to_datetime(df_principal[col_date]).max().date()
        
        # Sélection de la plage de dates pour le graphique
        debut_periode = st.date_input("Début de la période pour le graphique", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période pour le graphique", min_value=debut_periode, max_value=date_max, value=date_max)
    
    if st.button("Analyser"):
        # Conversion des dates et création des nouvelles colonnes
        df_principal[col_date] = pd.to_datetime(df_principal[col_date])
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrage des données en fonction de la période sélectionnée
        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & 
                                (df_principal[col_date].dt.date <= fin_periode)]

        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        # Calcul des répétitions pour le graphique
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        # Calcul des répétitions pour le tableau (toutes les dates)
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        # Graphique avec Plotly
        with col2:
            if periode_selectionnee != "Total":
                fig = px.bar(repetitions_graph, 
                             x=periode_selectionnee, 
                             y='Repetitions', 
                             color=col_prenom_nom, 
                             barmode='group',
                             title=f"Nombre de rapport d'intervention pour les opérateurs sélectionnés (du {debut_periode} au {fin_periode})",
                             text_auto=True)  # Afficher les valeurs sur les barres
            else:
                fig = px.bar(repetitions_graph, 
                             x=col_prenom_nom, 
                             y='Repetitions',
                             title=f"Total des rapports d'intervention pour les opérateurs sélectionnés (du {debut_periode} au {fin_periode})",
                             text_auto=True)  # Afficher les valeurs sur les barres

            st.plotly_chart(fig)
        
        st.subheader(f"Tableau des répétitions par {periode_selectionnee.lower()} (toutes les dates)")
        
        # Ajouter la colonne 'Opérateur' et afficher les résultats dans un tableau
        repetitions_tableau['Opérateur'] = repetitions_tableau[col_prenom_nom].map(df_principal.set_index(col_prenom_nom)['Opérateur'])
        colonnes_affichage = [col_prenom_nom, 'Opérateur', periode_selectionnee, 'Repetitions'] if periode_selectionnee != "Total" else [col_prenom_nom, 'Opérateur', 'Repetitions']
        st.dataframe(repetitions_tableau[colonnes_affichage], use_container_width=True)
        
        # Tirage au sort
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                # Convertir les liens d'images en affichage d'image
                if 'photo' in df_principal.columns:
                    lignes_tirees['photo'] = lignes_tirees['photo'].apply(lambda x: f'<img src="{x}" width="100">')  # Taille de l'image ajustée
                    st.markdown(lignes_tirees.to_html(escape=False), unsafe_allow_html=True)  # Utilisation de markdown pour afficher les images
                else:
                    st.dataframe(lignes_tirees, use_container_width=True)
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")

        # Téléchargement des données
        xlsx_data = to_xlsx(repetitions_tableau)
        st.download_button("Télécharger les données (xlsx)", xlsx_data, "repetitions.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        pdf_data = to_pdf(repetitions_tableau)
        st.download_button("Télécharger les données (PDF)", pdf_data, "repetitions.pdf", "application/pdf")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
