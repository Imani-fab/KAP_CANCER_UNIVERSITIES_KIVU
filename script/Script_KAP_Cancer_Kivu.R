rm(list = ls()) # Nettoyage de l'environnement

library(dplyr)
library(janitor)
library(stringr)

# ------------------------------------------------------------------------------
# 1. Rechargement propre depuis le disque
# ------------------------------------------------------------------------------
Cancer_raw <- read.csv2(
  "data/Cancer_clean_utf8.csv",
  header = TRUE,
  fill = TRUE,
  check.names = FALSE,
  fileEncoding = "latin1" # Empêche la corruption des caractères accentués
) %>% 
  clean_names()

# ------------------------------------------------------------------------------
# 2. Recodage sécurisé de TOUTES les variables clés
# ------------------------------------------------------------------------------
Cancer_clean <- Cancer_raw %>%
  # Nettoyage des espaces blancs invisibles en début/fin de chaîne
  mutate(across(where(is.character), str_trim)) %>%
  
  # Recodage robuste sans risque de conversion involontaire en NA
  mutate(
    # Sexe
    sexe = case_when(
      sexe %in% c("Feminin_", "Feminin", "Féminin") ~ "Female",
      sexe %in% c("Masculin", "Masculin_")          ~ "Male",
      TRUE ~ as.character(sexe)
    ) %>% factor(levels = c("Female", "Male")),
    
    # Province
    province = case_when(
      province %in% c("Autre_province_", "Autre_province") ~ "Other province",
      province %in% c("Hors_pays_", "Hors_pays")          ~ "Outside DRC",
      province %in% c("Maniema_", "Maniema")              ~ "Maniema",
      province %in% c("Nord_kivu", "Nord_Kivu")           ~ "North Kivu",
      province %in% c("Sud_kivu", "Sud_Kivu")             ~ "South Kivu",
      TRUE ~ as.character(province)
    ) %>% factor(levels = c("Other province", "Outside DRC", "Maniema", "North Kivu", "South Kivu")),
    
    # Statut matrimonial
    statut_matrimonial = case_when(
      grepl("celibataire", statut_matrimonial, ignore.case = TRUE) ~ "Single",
      grepl("Mari", statut_matrimonial, ignore.case = TRUE)        ~ "Married",
      TRUE ~ as.character(statut_matrimonial)
    ) %>% factor(levels = c("Single", "Married")),
    
    # Antécédent personnel
    antecedent_personnel_de_cancer = case_when(
      grepl("Non", antecedent_personnel_de_cancer, ignore.case = TRUE) ~ "No",
      grepl("Oui", antecedent_personnel_de_cancer, ignore.case = TRUE) ~ "Yes",
      TRUE ~ as.character(antecedent_personnel_de_cancer)
    ) %>% factor(levels = c("No", "Yes")),
    
    # Antécédent familial
    antecedent_familial_de_cancer = case_when(
      grepl("Non", antecedent_familial_de_cancer, ignore.case = TRUE) ~ "No",
      grepl("Oui", antecedent_familial_de_cancer, ignore.case = TRUE) ~ "Yes",
      TRUE ~ as.character(antecedent_familial_de_cancer)
    ) %>% factor(levels = c("No", "Yes")),
    
    # Entendu parler du cancer
    avoir_deja_entendu_cancer = case_when(
      avoir_deja_entendu_cancer %in% c("OUI", "Oui") ~ "Yes",
      avoir_deja_entendu_cancer %in% c("NON", "Non") ~ "No",
      TRUE ~ as.character(avoir_deja_entendu_cancer)
    ) %>% factor(levels = c("No", "Yes"))
  )

# ------------------------------------------------------------------------------
# 3. Vérification de la restauration
# ------------------------------------------------------------------------------
head(Cancer_clean[c("sexe", "province", "statut_matrimonial", 
                    "antecedent_personnel_de_cancer", "antecedent_familial_de_cancer")])

# ==============================================================================
# Codebook V2.0 Template : DICTIONNAIRE DE VARIABLES (KAP CANCER)
# ==============================================================================

# Packages ---------------------------------------------------------------------
library(readxl)
library(openxlsx)
library(dplyr)
library(stringr)

# 1. Importation de la liste initiale des variables ---------------------------
dictionary <- read_excel("output/Variable_Names_French.xlsx")

# 2. Construction et enrichissement complet du Codebook ------------------------
dictionary <- dictionary %>%
  
  # A. Génération des noms anglais et étiquettes pour le manuscrit
  mutate(
    english_variable = french_variable,
    english_variable = str_replace_all(
      english_variable,
      c(
        "age"                                              = "age",
        "sexe"                                             = "sex",
        "province"                                         = "province",
        "statut_matrimonial"                               = "marital_status",
        "duree_stage"                                      = "internship_duration",
        "avoir_deja_entendu_cancer"                        = "ever_heard_about_cancer",
        "source_information"                               = "source_of_information",
        "formation_medicale"                               = "medical_training",
        "famille_amis"                                     = "family_friends",
        "litterature_scientifique_livres"                  = "scientific_literature",
        "experience_clinique_durant_mon_stages"            = "clinical_experience",
        "caracteristic"                                    = "definition",
        "type_de_cancer_connu"                             = "known_cancer_types",
        "principaux_facteurs_de_risque"                    = "risk_factors",
        "tabagiste"                                        = "smoking",
        "consomation_d_alocool"                            = "alcohol_consumption",
        "mauvaise_alimentation"                            = "unhealthy_diet",
        "manque_d_exercice_physique"                       = "physical_inactivity",
        "infections"                                       = "infections",
        "genetique_hereditaire"                            = "genetic_hereditary",
        "vieillissement"                                   = "aging",
        "mythes_culturels_sorcellerie"                     = "witchcraft_myths",
        "le_cancer"                                        = "cancer",
        "detection_precoce"                                = "early_detection",
        "vaccination"                                      = "vaccination",
        "depistage"                                        = "screening",
        "arret_du_tabac"                                   = "smoking_cessation",
        "modalites_therapeutiques"                         = "treatment_modalities",
        "chirurgie"                                        = "surgery",
        "chimiotherapie"                                   = "chemotherapy",
        "immunotherapie"                                   = "immunotherapy",
        "therapie_cible"                                   = "targeted_therapy",
        "radiotherapie"                                    = "radiotherapy",
        "traitement_paliatif"                              = "palliative_care",
        "traitement_traditionnel"                          = "traditional_treatment",
        "role_principal_des_soins_palliatifs"              = "main_role_of_palliative_care",
        "antecedent_personnel_de_cancer"                   = "personal_history_of_cancer",
        "antecedent_familial_de_cancer"                    = "family_history_of_cancer",
        "relation_avec_la_personne_atteinte"               = "relationship_with_affected_person",
        "mode_de_traitement_observe_ou_vecu"               = "observed_treatment",
        "evolution_observee"                               = "observed_outcome",
        "perception"                                       = "perception"
      )
    ),
    
    manuscript_label = english_variable %>%
      str_replace_all("_", " ") %>%
      str_to_sentence(),
    
    # B. Domaines KAP
    kap_domain = case_when(
      french_variable %in% c(
        "age", "sexe", "province", "statut_matrimonial", "duree_stage", "university"
      ) ~ "Sociodemographic",
      
      str_detect(
        french_variable, 
        "antecedent|relation_avec|observe|evolution"
      ) ~ "Experience",
      
      str_detect(
        french_variable, 
        "perception|pense_|pensez_vous|priorite|mythes"
      ) ~ "Perception",
      
      str_detect(
        french_variable, 
        "precision|raison|proposee|poposition"
      ) ~ "Open-ended",
      
      TRUE ~ "Knowledge"
    ),
    
    # C. Types de variables
    variable_type = case_when(
      french_variable == "age" ~ "Continuous",
      
      str_detect(
        french_variable, 
        "precision|raison|proposee|poposition"
      ) ~ "Text",
      
      TRUE ~ "Binary/Categorical"
    ),
    
    # D. Règles de cotation (Scoring rules)
    scoring_rule = case_when(
      kap_domain == "Knowledge"  ~ "1=Correct ; 0=Incorrect",
      kap_domain == "Perception" ~ "Positive=1 ; Negative=0",
      TRUE                       ~ "NA"
    ),
    
    # E. Inclusion dans les scores
    include_in_knowledge_score = ifelse(kap_domain == "Knowledge", "YES", "NO"),
    include_in_perception_score = ifelse(kap_domain == "Perception", "YES", "NO"),
    
    # F. Champs à compléter manuellement / placeholders
    categories = "",
    
    # G. Réponses correctes (Connaissances)
    correct_answer = case_when(
      str_detect(french_variable, "^type_de_cancer_connu_") ~ "1",
      
      french_variable == "caracteristic_cancer_une_proliferation_incontrolee_de_cellules_anormales" ~ "1",
      french_variable %in% c(
        "caracteristic_cancer_une_infection_aigue",
        "caracteristic_cancer_une_inflammation_passagere",
        "caracteristic_cancer_une_reaction_allergique"
      ) ~ "0",
      
      str_detect(french_variable, "^principaux_facteurs_de_risque_") ~ "1",
      str_detect(french_variable, "^les_moyens_de_prevention_ou_de_depistage_du_cancer_") ~ "1",
      str_detect(french_variable, "^modalites_therapeutiques_") ~ "1",
      
      # Vrai / Faux
      french_variable == "le_cancer_touche_uniquement_les_personnes_agees" ~ "Faux_",
      french_variable == "la_detection_precoce_ameliore_les_resultats_du_traitement" ~ "VRAI",
      french_variable == "le_tabagisme_est_la_cause_la_plus_frequente_du_cancer" ~ "VRAI",
      french_variable == "certaines_vaccination_peuvent_prevenir_certains_cancers" ~ "VRAI",
      french_variable == "utilisation_d_ecrans_solaires_reduit_le_cancer_de_la_peau" ~ "VRAI",
      french_variable == "le_cancer_est_toujours_hereditaire" ~ "Faux_",
      french_variable == "le_cancer_du_sein_peut_toucher_les_hommes" ~ "VRAI",
      french_variable == "hpylori_peut_conduire_au_cancer_gastrique" ~ "VRAI",
      
      TRUE ~ NA_character_
    ),
    
    # Ajustement des réponses négatives/incorrectes
    correct_answer = ifelse(
      french_variable %in% c(
        "type_de_cancer_connu_aucun",
        "principaux_facteurs_de_risque_aucun",
        "principaux_facteurs_de_risque_je_ne_sais_pas",
        "principaux_facteurs_de_risque_mythes_culturels_sorcellerie",
        "les_moyens_de_prevention_ou_de_depistage_du_cancer_aucun",
        "modalites_therapeutiques_traitement_traditionnel",
        "modalites_therapeutiques_je_ne_sais_pas"
      ),
      "0",
      correct_answer
    ),
    
    # H. Réponses de perception positive
    positive_response = case_when(
      french_variable == "pense_que_le_cancer_peut_etre_gueri" ~ "Oui",
      french_variable == "pensez_vous_que_le_cancer_est_un_probleme_majeur_de_sante_publique_dans_la_sous_region_du_kivu" ~ "Oui",
      french_variable == "niveau_de_priorite_accordee_a_la_lutte_contre_le_cancer_en_rdc" ~ "Elevee",
      french_variable == "influence_des_mythes_et_traditions_sur_la_gestion_du_cancer" ~ "Oui",
      french_variable == "votre_perception_generale_du_cancer" ~ "Positive",
      TRUE ~ ""
    )
  )

# 3. Exportations Excel --------------------------------------------------------

# Export brouillon initial
write.xlsx(
  dictionary,
  file = "output/Codebook_V2_Draft.xlsx",
  rowNames = FALSE
)

# Export version complète avec catégories
write.xlsx(
  dictionary,
  file = "output/Codebook_V2_With_Categories.xlsx",
  rowNames = FALSE
)

# 4. Contrôles et vérifications du Codebook ------------------------------------

# Aperçu rapide de la répartition
table(dictionary$kap_domain)

# Vérification visuelle
dictionary %>%
  select(french_variable, english_variable, kap_domain, correct_answer, positive_response) %>%
  View()

# Vérification externe à partir du fichier nettoyé final
cd <- read_excel("output/Codebook_V2_correct_Analyse.xlsx")

glimpse(cd)
names(cd)

# Répartition des variables par domaine KAP
table(cd$kap_domain)

# Nombre de variables incluses dans les scores
cd %>% count(include_in_knowledge_score)
cd %>% count(include_in_perception_score)


# ==============================================================================
# PIPELINE COMPLÈTEMENT CORRIGÉ (SANS CONFLIT SELECT)
# ==============================================================================

library(dplyr)
library(readxl)
library(openxlsx)
library(stringr)
library(janitor)

# 1. Importation des données
Cancer_raw <- read.csv2(
  "data/Cancer_clean_utf8.csv",
  header = TRUE,
  fill = TRUE,
  check.names = FALSE,
  fileEncoding = "latin1"
) %>% clean_names()

Cancer <- Cancer_raw

# 2. Chargement du Codebook
cd <- read_excel("output/Codebook_V2_correct_Analyse.xlsx")

# ------------------------------------------------------------------------------
# SECTION A : KNOWLEDGE SCORE
# ------------------------------------------------------------------------------

knowledge_items <- cd %>% filter(include_in_knowledge_score == "YES")

# Cotation binaire
for (i in seq_len(nrow(knowledge_items))) {
  var       <- knowledge_items$french_variable[i]
  answer    <- knowledge_items$correct_answer[i]
  score_var <- paste0(var, "_score")
  
  if (var %in% names(Cancer)) {
    Cancer[[score_var]] <- ifelse(
      is.na(Cancer[[var]]),
      NA_real_,
      ifelse(trimws(as.character(Cancer[[var]])) == trimws(answer), 1, 0)
    )
  }
}

knowledge_score_vars <- paste0(knowledge_items$french_variable, "_score")
knowledge_score_vars <- intersect(knowledge_score_vars, names(Cancer))
max_knowledge_score  <- length(knowledge_score_vars)

# Calcul sécurisé avec across()
Cancer <- Cancer %>%
  mutate(
    knowledge_score      = rowSums(across(all_of(knowledge_score_vars)), na.rm = TRUE),
    knowledge_percentage = (knowledge_score / max_knowledge_score) * 100,
    knowledge_level      = case_when(
      knowledge_percentage < 50 ~ "Poor knowledge",
      knowledge_percentage < 65 ~ "Insufficient knowledge",
      knowledge_percentage < 85 ~ "Moderate knowledge",
      TRUE                      ~ "Good knowledge"
    )
  )

# Contrôles Connaissances
summary(Cancer$knowledge_score)
table(Cancer$knowledge_level)

# ------------------------------------------------------------------------------
# SECTION B : PERCEPTION SCORE
# ------------------------------------------------------------------------------

cd_work <- cd %>%
  mutate(
    positive_response = case_when(
      french_variable == "pense_que_le_cancer_peut_etre_gueri" ~ "VRAI",
      french_variable == "pensez_vous_que_la_prevention_le_diagnostic_et_la_prise_en_charge_du_cancer_sont_adequates_en_rdc" ~ "Non",
      french_variable == "niveau_de_priorite_accordee_a_la_lutte_contre_le_cancer_en_rdc" ~ "5",
      french_variable == "pensez_vous_que_le_cancer_est_un_probleme_majeur_de_sante_publique_dans_la_sous_region_du_kivu" ~ "VRAI",
      french_variable == "influence_des_mythes_et_traditions_sur_la_gestion_du_cancer" ~ "Oui_,_beaucoup;Oui_,_moderement_",
      TRUE ~ positive_response
    )
  )

perception_items <- cd_work %>%
  filter(include_in_perception_score == "YES") %>%
  dplyr::select(french_variable, positive_response)

for (i in seq_len(nrow(perception_items))) {
  var       <- perception_items$french_variable[i]
  positive  <- perception_items$positive_response[i]
  score_var <- paste0(var, "_score")
  
  positive_values <- trimws(strsplit(positive, ";")[[1]])
  
  if (var %in% names(Cancer)) {
    Cancer[[score_var]] <- ifelse(
      is.na(Cancer[[var]]) | Cancer[[var]] == "",
      NA_real_,
      ifelse(trimws(as.character(Cancer[[var]])) %in% positive_values, 1, 0)
    )
  }
}

perception_score_vars <- paste0(perception_items$french_variable, "_score")
perception_score_vars <- intersect(perception_score_vars, names(Cancer))
max_perception_score  <- length(perception_score_vars)

# Calcul sécurisé avec across()
Cancer <- Cancer %>%
  mutate(
    perception_score      = rowSums(across(all_of(perception_score_vars)), na.rm = TRUE),
    perception_percentage = (perception_score / max_perception_score) * 100,
    perception_level      = case_when(
      perception_percentage < 50 ~ "Negative perception",
      TRUE                       ~ "Positive perception"
    )
  )

# Contrôles Perception
summary(Cancer$perception_score)
table(Cancer$perception_level)

# ==============================================================================
# SECTION B : MISE À JOUR ET EXPORT DU CODEBOOK PERCEPTION
# ==============================================================================

# 1. Ajustement des réponses positives dans la copie de travail
cd_work <- cd %>%
  mutate(
    positive_response = case_when(
      french_variable == "pense_que_le_cancer_peut_etre_gueri" ~ "VRAI",
      french_variable == "pensez_vous_que_la_prevention_le_diagnostic_et_la_prise_en_charge_du_cancer_sont_adequates_en_rdc" ~ "Non",
      french_variable == "niveau_de_priorite_accordee_a_la_lutte_contre_le_cancer_en_rdc" ~ "5",
      french_variable == "pensez_vous_que_le_cancer_est_un_probleme_majeur_de_sante_publique_dans_la_sous_region_du_kivu" ~ "VRAI",
      french_variable == "influence_des_mythes_et_traditions_sur_la_gestion_du_cancer" ~ "Oui_,_beaucoup;Oui_,_moderement_",
      TRUE ~ positive_response
    )
  )

# 2. Vérification visuelle sécurisée avec dplyr::select
cd_work %>%
  filter(include_in_perception_score == "YES") %>%
  dplyr::select(french_variable, positive_response)

# 3. Export du Codebook corrigé
openxlsx::write.xlsx(
  cd_work,
  "output/Codebook_V2_correct_Analyse_PerceptionFixed.xlsx",
  overwrite = TRUE
)

# 4. Export du jeu de données final enrichi avec tous les scores
openxlsx::write.xlsx(
  Cancer,
  "output/Cancer_Clean_With_Scores.xlsx",
  overwrite = TRUE
)


# ==============================================================================
# SECTION C : SCORE DE PERCEPTION (PERCEPTION SCORE)
# ==============================================================================

# 1. Filtrage des items de perception ajustés (utilisation de dplyr::select)
perception_items <- cd_work %>%
  filter(include_in_perception_score == "YES") %>%
  dplyr::select(french_variable, positive_response)

glimpse(perception_items)

# 2. Cotation binaire pour chaque item de perception (avec gestion des réponses multiples ';')
for (i in seq_len(nrow(perception_items))) {
  var       <- perception_items$french_variable[i]
  positive  <- perception_items$positive_response[i]
  score_var <- paste0(var, "_score")
  
  # Séparation des réponses positives multiples et nettoyage des espaces
  positive_values <- trimws(strsplit(positive, ";")[[1]])
  
  if (var %in% names(Cancer)) {
    Cancer[[score_var]] <- ifelse(
      is.na(Cancer[[var]]) | trimws(as.character(Cancer[[var]])) == "",
      NA_real_,
      ifelse(trimws(as.character(Cancer[[var]])) %in% positive_values, 1, 0)
    )
  } else {
    warning(paste("Variable introuvable dans 'Cancer':", var))
  }
}

# 3. Calcul du score total, du pourcentage et du niveau de perception (sécurisé avec across)
perception_score_vars <- paste0(perception_items$french_variable, "_score")
perception_score_vars <- intersect(perception_score_vars, names(Cancer))
max_perception_score  <- length(perception_score_vars)

Cancer <- Cancer %>%
  mutate(
    perception_score      = rowSums(across(all_of(perception_score_vars)), na.rm = TRUE),
    perception_percentage = (perception_score / max_perception_score) * 100,
    perception_level      = case_when(
      perception_percentage < 50 ~ "Negative perception",
      TRUE                       ~ "Positive perception"
    )
  )

# ==============================================================================
# SECTION D : CONTRÔLES ET VÉRIFICATIONS QUALITÉ
# ==============================================================================

cat("\nMaximum perception score = ", max_perception_score, "\n")

summary(Cancer$perception_score)
summary(Cancer$perception_percentage)
table(Cancer$perception_level)

# Vérification item par item pour le score de perception
for (v in perception_score_vars) {
  cat("\n============================\n")
  cat(v, "\n")
  print(table(Cancer[[v]], useNA = "ifany"))
}



# ==============================================================================
# ANALYSE DESCRIPTIVE : TABLES 1 À 5 POUR LE MANUSCRIT
# ==============================================================================

library(dplyr)
library(gtsummary)
library(flextable)
library(officer)
library(stringr)

# Initialisation du document Word
doc <- read_docx()

# Helper pour mise en forme flextable propre
style_table <- function(ft) {
  ft %>%
    theme_vanilla() %>%
    bold(part = "header") %>%
    autofit()
}

# ==============================================================================
# TABLE 1 : Sociodemographic Characteristics
# ==============================================================================

vars_socio <- cd_work %>%
  filter(kap_domain == "Sociodemographic") %>%
  pull(french_variable)

vars_socio_valid <- intersect(vars_socio, names(Cancer))

table1 <- Cancer %>%
  dplyr::select(all_of(vars_socio_valid)) %>%
  tbl_summary(
    missing = "no",
    statistic = list(
      all_continuous() ~ "{mean} ± {sd}",
      all_categorical() ~ "{n} ({p}%)"
    )
  ) %>%
  bold_labels()

doc <- body_add_par(
  doc,
  "Table 1. Sociodemographic Characteristics of Medical Interns (N = 262)",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(as_flex_table(table1)))

# ==============================================================================
# TABLE 2 : Knowledge Items
# ==============================================================================

knowledge_items <- cd_work %>%
  filter(include_in_knowledge_score == "YES")

table2_list <- list()

for (i in seq_len(nrow(knowledge_items))) {
  var   <- knowledge_items$french_variable[i]
  label <- knowledge_items$manuscript_label[i]
  score_var <- paste0(var, "_score")
  
  if (score_var %in% names(Cancer)) {
    n_correct <- sum(Cancer[[score_var]] == 1, na.rm = TRUE)
    pct_correct <- round(100 * n_correct / nrow(Cancer), 1)
  } else {
    n_correct <- NA
    pct_correct <- NA
  }
  
  table2_list[[i]] <- data.frame(
    `Knowledge Item` = label,
    `Correct Responses (n)` = n_correct,
    `Correct Responses (%)` = pct_correct,
    check.names = FALSE
  )
}

table2 <- bind_rows(table2_list)

doc <- body_add_par(
  doc,
  "Table 2. Cancer Knowledge Among Medical Interns",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(flextable(table2)))

# ==============================================================================
# TABLE 2B : Performance by Knowledge Domain
# ==============================================================================

cd_work <- cd_work %>%
  mutate(
    knowledge_domain = case_when(
      str_detect(french_variable, "avoir_deja_entendu_cancer|source_information") ~ "Cancer Awareness and Information Sources",
      str_detect(french_variable, "caracteristic_cancer") ~ "Definition of Cancer",
      str_detect(french_variable, "type_de_cancer_connu|type_de_cancer_conn") ~ "Cancer Types",
      str_detect(french_variable, "facteurs_de_risque") ~ "Risk Factors",
      str_detect(french_variable, "le_cancer|detection_precoce|hpylori|tabagisme_est_la_cause") ~ "General Cancer Knowledge",
      str_detect(french_variable, "moyens_de_prevention|vaccination|si_vrai_quand") ~ "Prevention and Screening",
      str_detect(french_variable, "modalites_therapeutiques") ~ "Treatment Modalities",
      str_detect(french_variable, "role_principal_des_soins_palliatifs") ~ "Palliative Care",
      str_detect(french_variable, "principales_causes") ~ "Cancer Burden and Barriers",
      TRUE ~ "Unclassified"
    )
  )

knowledge_items_domain <- cd_work %>%
  filter(include_in_knowledge_score == "YES")

# Calcul du % moyen de bonnes réponses par item
knowledge_items_domain$correct_percent <- sapply(
  knowledge_items_domain$french_variable,
  function(x) {
    score_v <- paste0(x, "_score")
    if (score_v %in% names(Cancer)) {
      mean(Cancer[[score_v]] == 1, na.rm = TRUE) * 100
    } else {
      NA_real_
    }
  }
)

domain_summary <- knowledge_items_domain %>%
  group_by(knowledge_domain) %>%
  summarise(
    `Number of Items` = n(),
    `Mean Correct (%)` = round(mean(correct_percent, na.rm = TRUE), 1),
    .groups = "drop"
  ) %>%
  arrange(knowledge_domain) %>%
  rename(`Knowledge Domain` = knowledge_domain)

# Vérification du nombre total d'items (doit égaler 74)
cat("\nNombre total d'items analysés par domaine : ", sum(domain_summary$`Number of Items`), "\n")

doc <- body_add_par(
  doc,
  "Table 2B. Performance by Knowledge Domain",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(flextable(domain_summary)))

# ==============================================================================
# TABLE 3 : Perception Items
# ==============================================================================

perception_items <- cd_work %>%
  filter(include_in_perception_score == "YES")

table3_list <- list()

for (i in seq_len(nrow(perception_items))) {
  var   <- perception_items$french_variable[i]
  label <- perception_items$manuscript_label[i]
  score_var <- paste0(var, "_score")
  
  if (score_var %in% names(Cancer)) {
    n_pos <- sum(Cancer[[score_var]] == 1, na.rm = TRUE)
    pct_pos <- round(100 * n_pos / nrow(Cancer), 1)
  } else {
    n_pos <- NA
    pct_pos <- NA
  }
  
  table3_list[[i]] <- data.frame(
    `Perception Item` = label,
    `Positive Responses (n)` = n_pos,
    `Positive Responses (%)` = pct_pos,
    check.names = FALSE
  )
}

table3 <- bind_rows(table3_list)

doc <- body_add_par(
  doc,
  "Table 3. Perceptions and Attitudes Toward Cancer",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(flextable(table3)))

# ==============================================================================
# TABLE 4 : Cancer-Related Experiences
# ==============================================================================

vars_experience <- cd_work %>%
  filter(kap_domain == "Experience") %>%
  pull(french_variable)

vars_exp_valid <- intersect(vars_experience, names(Cancer))

table4 <- Cancer %>%
  dplyr::select(all_of(vars_exp_valid)) %>%
  tbl_summary(
    missing = "no",
    statistic = all_categorical() ~ "{n} ({p}%)"
  ) %>%
  bold_labels()

doc <- body_add_par(
  doc,
  "Table 4. Cancer-Related Experiences",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(as_flex_table(table4)))

# ==============================================================================
# TABLE 5 : Knowledge and Perception Score Distribution
# ==============================================================================

# Table 5A: Résumé continu des scores
table5a <- data.frame(
  Score = c(
    "Knowledge Score",
    "Knowledge Percentage",
    "Perception Score",
    "Perception Percentage"
  ),
  `Mean ± SD` = c(
    paste0(round(mean(Cancer$knowledge_score, na.rm = TRUE), 2), " ± ", round(sd(Cancer$knowledge_score, na.rm = TRUE), 2)),
    paste0(round(mean(Cancer$knowledge_percentage, na.rm = TRUE), 2), " ± ", round(sd(Cancer$knowledge_percentage, na.rm = TRUE), 2)),
    paste0(round(mean(Cancer$perception_score, na.rm = TRUE), 2), " ± ", round(sd(Cancer$perception_score, na.rm = TRUE), 2)),
    paste0(round(mean(Cancer$perception_percentage, na.rm = TRUE), 2), " ± ", round(sd(Cancer$perception_percentage, na.rm = TRUE), 2))
  ),
  check.names = FALSE
)

# Table 5B: Distribution par catégories
knowledge_dist <- Cancer %>%
  count(knowledge_level) %>%
  mutate(Percent = round(100 * n / sum(n), 1)) %>%
  rename(`Knowledge Level` = knowledge_level, `N` = n, `Percent (%)` = Percent)

perception_dist <- Cancer %>%
  count(perception_level) %>%
  mutate(Percent = round(100 * n / sum(n), 1)) %>%
  rename(`Perception Level` = perception_level, `N` = n, `Percent (%)` = Percent)

doc <- body_add_par(
  doc,
  "Table 5. Knowledge and Perception Score Distribution",
  style = "heading 1"
)
doc <- body_add_flextable(doc, style_table(flextable(table5a)))

doc <- body_add_par(doc, "Knowledge Level Distribution", style = "heading 2")
doc <- body_add_flextable(doc, style_table(flextable(knowledge_dist)))

doc <- body_add_par(doc, "Perception Level Distribution", style = "heading 2")
doc <- body_add_flextable(doc, style_table(flextable(perception_dist)))

# ==============================================================================
# EXPORT DU DOCUMENT WORD & CONTRÔLES
# ==============================================================================

print(doc, target = "output/KAP_Cancer_Tables.docx")

cat("\n[SUCCÈS] Le fichier 'output/KAP_Cancer_Tables.docx' a été créé avec succès.\n")

# Proportions en console
cat("\n--- Distribution Connaissances (%) ---\n")
print(round(prop.table(table(Cancer$knowledge_level)) * 100, 1))

cat("\n--- Distribution Perception (%) ---\n")
print(round(prop.table(table(Cancer$perception_level)) * 100, 1))


#============================================================================
# TABLE 6 : EXPLORATORY ASSOCIATIONS (KNOWLEDGE SCORE)
#============================================================================

library(dplyr)
library(officer)
library(flextable)

# 1. FONCTIONS DE CALCUL CORRIGÉES

# Variables binaires (t-test)
analyse_binary <- function(var) {
  test <- t.test(Cancer$knowledge_score ~ Cancer[[var]])
  Cancer %>%
    filter(!is.na(.data[[var]])) %>%
    group_by(Category = as.character(.data[[var]])) %>%
    summarise(
      n = n(),
      Mean_Knowledge_Score = sprintf("%.2f (%.2f)", mean(knowledge_score, na.rm = TRUE), sd(knowledge_score, na.rm = TRUE)),
      .groups = "drop"
    ) %>%
    mutate(
      Variable = var,
      Test = "Student t-test",
      Statistic = c(as.character(round(as.numeric(test$statistic), 3)), rep("", n() - 1)),
      P_value = c(as.character(signif(test$p.value, 3)), rep("", n() - 1))
    ) %>%
    dplyr::select(Variable, Category, n, Mean_Knowledge_Score, Test, Statistic, P_value)
}

# Variable multinomiale (ANOVA - Province)
analyse_anova <- function(var) {
  fml <- as.formula(paste("knowledge_score ~", var))
  fit <- aov(fml, data = Cancer)
  res <- summary(fit)[[1]]
  
  Cancer %>%
    filter(!is.na(.data[[var]])) %>%
    group_by(Category = as.character(.data[[var]])) %>%
    summarise(
      n = n(),
      Mean_Knowledge_Score = sprintf("%.2f (%.2f)", mean(knowledge_score, na.rm = TRUE), sd(knowledge_score, na.rm = TRUE)),
      .groups = "drop"
    ) %>%
    mutate(
      Variable = var,
      Test = "ANOVA",
      Statistic = c(as.character(round(res[["F value"]][1], 3)), rep("", n() - 1)),
      P_value = c(as.character(signif(res[["Pr(>F)"]][1], 3)), rep("", n() - 1))
    ) %>%
    dplyr::select(Variable, Category, n, Mean_Knowledge_Score, Test, Statistic, P_value)
}

# Variables quantitatives (Spearman)
analyse_spearman <- function(var) {
  test <- cor.test(Cancer$knowledge_score, Cancer[[var]], method = "spearman", exact = FALSE)
  data.frame(
    Variable = var,
    Category = "Spearman rho",
    n = sum(!is.na(Cancer[[var]]) & !is.na(Cancer$knowledge_score)),
    Mean_Knowledge_Score = "—",
    Test = "Spearman",
    Statistic = as.character(round(as.numeric(test$estimate), 3)),
    P_value = as.character(signif(test$p.value, 3)),
    stringsAsFactors = FALSE
  )
}

# 2. CONSTRUCTION DU TABLEAU 6

binary_vars <- c("sexe", "statut_matrimonial", "antecedent_personnel_de_cancer", "antecedent_familial_de_cancer")
quant_vars  <- c("age", "duree_stage", "perception_score")

table6 <- bind_rows(
  lapply(binary_vars, analyse_binary),
  analyse_anova("province"),
  lapply(quant_vars, analyse_spearman)
)

# 3. EXPORT WORD

doc <- read_docx() %>%
  body_add_par("Table 6. Exploratory Associations with Cancer Knowledge Score", style = "heading 1") %>%
  body_add_flextable(
    flextable(table6) %>% 
      theme_vanilla() %>% 
      autofit()
  )

print(doc, target = "output/Table_6_Exploratory_Associations.docx")


#============================================================================
# association between knowledge score and perception score
#===========================================================================


library(dplyr)
library(ggplot2)
library(car)

# ------------------------------------------------------------------------------
# 1. FIGURE 1 : Distribution du score de connaissances
# ------------------------------------------------------------------------------

# Calculs descriptifs sécurisés (gestion des NA)
stats_k <- Cancer %>%
  summarise(
    mean = round(mean(knowledge_score, na.rm = TRUE), 1),
    sd   = round(sd(knowledge_score, na.rm = TRUE), 1),
    med  = round(median(knowledge_score, na.rm = TRUE), 1),
    q1   = round(quantile(knowledge_score, 0.25, na.rm = TRUE), 1),
    q3   = round(quantile(knowledge_score, 0.75, na.rm = TRUE), 1)
  )

annot_fig1 <- sprintf(
  "Moyenne (SD) = %.1f (%.1f)\nMédienne (IQR) = %.1f (%.1f–%.1f)",
  stats_k$mean, stats_k$sd, stats_k$med, stats_k$q1, stats_k$q3
)

fig1 <- ggplot(Cancer, aes(x = knowledge_score)) +
  geom_histogram(
    aes(y = after_stat(count / sum(count) * 100)),
    bins = 20, fill = "grey80", color = "black", linewidth = 0.4
  ) +
  geom_vline(xintercept = stats_k$mean, linetype = "dashed", linewidth = 0.8, color = "red") +
  annotate(
    "text", x = -Inf, y = Inf, label = annot_fig1,
    hjust = -0.05, vjust = 1.2, size = 4
  ) +
  labs(
    title = "Distribution du score de connaissances",
    x = "Score de connaissances",
    y = "Proportion (%)"
  ) +
  theme_classic(base_size = 13) +
  theme(plot.title = element_text(face = "bold", hjust = 0.5))

ggsave("output/model_analysis/Figure_1_Knowledge_Score.png", fig1, width = 7, height = 5, dpi = 300)


# ------------------------------------------------------------------------------
# 2. RÉGRESSION ET DIAGNOSTIC DES RÉSIDUS (Étape corrigée)
# ------------------------------------------------------------------------------

# Modèle linéaire simple
lm_fig <- lm(knowledge_score ~ perception_score, data = Cancer)
summary_lm <- summary(lm_fig)

# Test de normalité sur les RÉSIDUS du modèle (et non sur Y)
res_shapiro <- shapiro.test(residuals(lm_fig))
message("P-value normalité des résidus : ", round(res_shapiro$p.value, 4))


# ------------------------------------------------------------------------------
# 3. FIGURE 2 : Association Perception vs Connaissances
# ------------------------------------------------------------------------------

# Extraction et formatage des paramètres
b0 <- round(coef(lm_fig)[1], 2)
b1 <- round(coef(lm_fig)[2], 2)
r2 <- round(summary_lm$r.squared, 3)
p_val <- summary_lm$coefficients["perception_score", "Pr(>|t|)"]
p_text <- ifelse(p_val < 0.001, "p < 0.001", sprintf("p = %.3f", p_val))

eq_label <- sprintf("Score Connaissances = %.2f + %.2f × Perception\nR² = %.3f | %s", b0, b1, r2, p_text)

fig2 <- ggplot(Cancer, aes(x = perception_score, y = knowledge_score)) +
  geom_jitter(width = 0.15, height = 0, alpha = 0.5, size = 2) +
  geom_smooth(method = "lm", se = TRUE, color = "blue", linewidth = 1) +
  annotate(
    "text", x = -Inf, y = Inf, label = eq_label,
    hjust = -0.05, vjust = 1.2, size = 4.2, fontface = "italic"
  ) +
  labs(
    title = "Association entre Perception et Connaissances",
    x = "Score de perception",
    y = "Score de connaissances"
  ) +
  theme_classic(base_size = 13) +
  theme(plot.title = element_text(face = "bold", hjust = 0.5))

ggsave("output/model_analysis/Figure_2_Perception_vs_Knowledge.png", fig2, width = 7, height = 5, dpi = 300)



# ==============================================================================
# ANALYSIS MULTIVARIÉE : MODÈLE POUR VARIABLE DE COMPTAGE (KNOWLEDGE_SCORE)
# ==============================================================================

# 1. Chargement des packages
library(MASS)        # glm.nb (Binomiale Négative) - Placé en premier pour limiter les conflits
library(dplyr)       # Manipulation de données
library(car)         # Test VIF / GVIF
library(performance) # Overdispersion, Pseudo-R2, Métriques
library(ggplot2)     # Graphiques
library(DHARMa)      # Diagnostics par résidus simulés
library(gtsummary)   # Tableaux de régression
library(flextable)   # Exportation Word (.docx)

# Création du dossier de sortie
dir.create("output/model_analysis", showWarnings = FALSE, recursive = TRUE)

# ------------------------------------------------------------------------------
# 2. Ajustement des Modèles (Poisson & Binomiale Négative)
# ------------------------------------------------------------------------------

# Modèle A: Régression de Poisson
model_poisson <- glm(
  knowledge_score ~ age + sexe + province + statut_matrimonial +
    duree_stage + antecedent_personnel_de_cancer +
    antecedent_familial_de_cancer + perception_score,
  family = poisson(link = "log"),
  data = Cancer
)

# Test de surdispersion pour valider la Poisson
check_overdispersion(model_poisson)

# Modèle B: Régression Binomiale Négative (Modèle Retenu)
model_final <- glm.nb(
  knowledge_score ~ age + sexe + province + statut_matrimonial +
    duree_stage + antecedent_personnel_de_cancer +
    antecedent_familial_de_cancer + perception_score,
  data = Cancer
)

# Test de surdispersion du modèle final
check_overdispersion(model_final)

# ------------------------------------------------------------------------------
# 3. Diagnostic & Colinéarité
# ------------------------------------------------------------------------------

# Multicolinéarité (GVIF ajusté pour variables multinomiales comme province)
vif_final <- car::vif(model_final)
print(vif_final)

# Performance globale du modèle
perf <- performance(model_final)
print(perf)

# ------------------------------------------------------------------------------
# 4. Diagnostics des Résidus Simulés (DHARMa)
# ------------------------------------------------------------------------------

simulation_res <- simulateResiduals(fittedModel = model_final, n = 1000)

# Tests formels
testUniformity(simulation_res) # Test d'uniformité (Kolmogorov-Smirnov)
testDispersion(simulation_res) # Test de dispersion résiduelle
testOutliers(simulation_res)   # Test des valeurs aberrantes

# Exportation sécurisée de la figure DHARMa sans l'erreur de lissage
png(
  "output/model_analysis/Figure_DHARMa_Diagnostics.png",
  width = 10, height = 5, units = "in", res = 300
)

# Option smooth = FALSE désactive le spline problématique tout en gardant tous les résidus et tests visuels
plot(simulation_res, smooth = FALSE)

dev.off()
# ------------------------------------------------------------------------------
# 5. Génération du Tableau de Publication (Style JAMA - English)
# ------------------------------------------------------------------------------

# Configuration du thème gtsummary
theme_gtsummary_journal(journal = "jama")
theme_gtsummary_language(language = "en")

table_publication <- model_final %>%
  tbl_regression(
    exponentiate = TRUE, # Conversion des coefficients en IRR
    pvalue_fun   = ~ style_pvalue(.x, digits = 3),
    label        = list(
      age                            ~ "Age (years)",
      sexe                           ~ "Sex",
      province                       ~ "Province of origin",
      statut_matrimonial             ~ "Marital status",
      duree_stage                    ~ "Internship duration",
      antecedent_personnel_de_cancer ~ "Personal history of cancer",
      antecedent_familial_de_cancer  ~ "Family history of cancer",
      perception_score               ~ "Perception score"
    )
  ) %>%
  bold_p(t = 0.05) %>%
  bold_labels() %>%
  modify_header(
    label    ~ "**Predictors**",
    estimate ~ "**IRR (95% CI)**",
    p.value  ~ "**p-value**"
  ) %>%
  add_glance_source_note(
    label = list(
      nobs ~ "Observations",
      AIC  ~ "Akaike Information Criterion (AIC)"
    ),
    include = c(nobs, AIC)
  ) %>%
  modify_footnote(
    all_stat_cols() ~ "IRR: Incidence Rate Ratio; CI: Confidence Interval. Model Diagnostics: Residual uniformity verified via DHARMa simulated quantile residuals; Overdispersion test passed; Multicollinearity absent (adjusted GVIF < 1.40)."
  )

# Affichage console
table_publication

# ------------------------------------------------------------------------------
# 6. Exportation vers Microsoft Word
# ------------------------------------------------------------------------------

table_publication %>%
  as_flex_table() %>%
  flextable::save_as_docx(
    path = "output/model_analysis/Table_Negative_Binomial_Model.docx"
  )
