import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Fonction de chargement des données
@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

# Fonction pour convertir un dataframe en fichier XLSX
def convert_df_to_xlsx(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Fonction pour générer un PDF
def generate_pdf(df, filename="tableau.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, "Tableau des répétitions des opérateurs")
    
    y_position = height - 60
    for i, row in df.iterrows():
        text = f"{row['Prénom et nom']} : {row['Repetitions']}"
        c.drawString(30, y_position, text)
        y_position -= 20
    
    c.save()

# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")

st.title("📊 Analyse des interventions des opérateurs")

# Chargement du fichier Excel
fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # Sélection de la colonne pour 'Prénom et nom' et de la colonne de date
        col_prenom_nom = df_principal.columns[4]  # Sélection automatique de la première colonne
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique()
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année", "Total"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        # Tentative de conversion des dates avec gestion des erreurs
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        
        # Gérer les dates invalides et définir les bornes min et max des dates valides
        date_min = df_principal[col_date].min()
        date_max = df_principal[col_date].max()

        # Si les dates sont invalides, on avertit l'utilisateur
        if pd.isna(date_min) or pd.isna(date_max):
            st.warning("Certaines dates dans le fichier sont invalides. Elles ont été ignorées.")
            date_min = date_max = None
        
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)
    
    # Quand le bouton "Analyser" est cliqué
    if st.button("Analyser"):
        # Filtrage des données pour garder seulement les dates valides
        valid_dates_df = df_principal.dropna(subset=[col_date])
        if len(df_principal) != len(valid_dates_df):
            st.warning(f"{len(df_principal) - len(valid_dates_df)} lignes avec des dates invalides ont été ignorées.")
        
        # Ajout des périodes (Jour, Semaine, Mois, Trimestre, Année)
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrer les données pour la période sélectionnée
        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]

        # Choisir les colonnes pour grouper les données
        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        # Calcul des répétitions pour le graphique
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        # Limiter le nombre de données affichées dans le graphique si trop volumineux
        if len(repetitions_graph) > 100:
            repetitions_graph = repetitions_graph.sample(n=100)
        
        # Calcul des répétitions pour le tableau (toutes les dates)
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        # Affichage du graphique avec les valeurs des répétitions et couleurs par opérateur
        with col2:
            fig = px.bar(repetitions_graph, 
                         x=periode_selectionnee if periode_selectionnee != "Jour" else col_prenom_nom,
                         y='Repetitions',
                         barmode='group',
                         color=col_prenom_nom,  # Ajout de la colonne 'Prénom et nom' pour les couleurs
                         title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})")
            fig.update_traces(text=repetitions_graph['Repetitions'], textposition='outside')
            st.plotly_chart(fig)

            # Calcul de la moyenne totale par période
            moyenne_totale = repetitions_graph['Repetitions'].mean()
            st.write(f"### Moyenne totale par période : {moyenne_totale:.2f}")
            
            # Calcul de la moyenne par période et par opérateur
            moyenne_par_operateur = repetitions_graph.groupby([periode_selectionnee, col_prenom_nom])['Repetitions'].mean().reset_index()
            st.write("### Moyenne par période et par opérateur :")
            st.dataframe(moyenne_par_operateur, use_container_width=True)

        # Affichage du tableau des répétitions
        st.subheader(f"Tableau du nombre des rapports d'intervention par {periode_selectionnee.lower()} (toutes les dates)")
        
        colonnes_affichage = [col_prenom_nom, periode_selectionnee, 'Repetitions'] if periode_selectionnee != "Total" else [col_prenom_nom, 'Repetitions']
        tableau_affichage = repetitions_tableau[colonnes_affichage]
        
        st.dataframe(tableau_affichage, use_container_width=True)
        
        # Tirage au sort pour deux lignes par opérateur
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                # Vérification de la validité des URLs des photos
                if pd.notna(lignes_tirees['Photo']).all() and pd.notna(lignes_tirees['Photo 2']).all():
                    lignes_tirees['Photo'] = lignes_tirees['Photo'].apply(lambda x: f'<img src="{x}" width="100"/>')
                    lignes_tirees['Photo 2'] = lignes_tirees['Photo 2'].apply(lambda x: f'<img src="{x}" width="100"/>')
                    st.markdown(lignes_tirees.to_html(escape=False), unsafe_allow_html=True)
                else:
                    st.warning("Certaines photos ne sont pas valides ou manquantes.")
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")
                            

        # Téléchargement du fichier XLSX
        st.subheader("Télécharger le tableau des rapports d'interventions")
        xlsx_data = convert_df_to_xlsx(repetitions_tableau)
        st.download_button(label="Télécharger en XLSX", data=xlsx_data, file
