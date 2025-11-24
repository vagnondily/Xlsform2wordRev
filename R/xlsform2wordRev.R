# =====================================================================
# XLSForm -> Word (Rendu Word) - Script final et débogué et corrigé V2
# =====================================================================

suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr); library(tidyr)
  library(purrr);  library(officer); library(flextable); library(glue); library(tools)
  library(tibble); library(rlang)
})

# Opérateur de coalescence robuste corrigé pour gérer les vecteurs
`%||%` <- function(a, b) {
  if (is.null(a) || length(a) == 0 || (is.atomic(a) && length(a) == 1 && (is.na(a) || !nzchar(a)))) {
    b
  } else {
    a
  }
}

# ---------------------------------------------------------------------
# PARAMÈTRES DE STYLE POUR UN RENDU PROFESSIONNEL
# ---------------------------------------------------------------------
WFP_BLUE    <- "#0A66C2"   
GREY_TXT    <- "#777777"
WFP_DARK_BLUE <- "#001F3F" 
GREY_BG     <- "#F2F2F2"   
WHITE_TXT   <- "#FFFFFF"   
RED_TXT     <- "#C00000" 
FONT_FAMILY <- "Cambria (Body)"     
LINE_SP     <- 1.0        
INDENT_Q    <- 0.3        
INDENT_C    <- 0.5        

FS_TITLE    <- 14
FS_SUBTITLE <- 12
FS_BLOCK    <- 12
FS_META     <- 10    
FS_HINT     <- 9   
FS_RELV     <- 9   
FS_MISC     <- 10     
FS_Q_BLUE   <- 11  
FS_Q_GREY   <- 9   

EXCLUDE_TYPES <- c("calculate","start","end","today","deviceid","subscriberid","simserial","phonenumber","username","instanceid","end_group", "end group", "end_repeat", "end repeat")
EXCLUDE_NAMES <- c("start","end","today","deviceid","username","instanceID","instanceid")

XLSFORM_PATH   <- NULL
TEMPLATE_DOCX  <- ""  
LOGO_PATH  <- "cp_logo.png"    
get_downloads_dir <- function() {
  if (Sys.info()["sysname"] == "Windows") {
    # Sur Windows, utilise la variable d'environnement USERPROFILE et ajoute 'Downloads'
    return(file.path(Sys.getenv("USERPROFILE"), "Downloads"))
  } else {
    # Sur macOS/Linux, utilise la variable d'environnement HOME et ajoute 'Downloads'
    return(file.path(Sys.getenv("HOME"), "Downloads"))
  }
}

OUTPUT_DIR <- get_downloads_dir()
# Nettoyage supplémentaire au cas où le chemin détecté ne serait pas parfait
OUTPUT_DIR <- normalizePath(OUTPUT_DIR, winslash = "/") 


# ---------------------------------------------------------------------
# Fonctions utilitaires 
# ---------------------------------------------------------------------

generate_output_path <- function(base_path) {
  if (!file.exists(base_path)) return(base_path)
  path_dir <- dirname(base_path)
  path_ext <- tools::file_ext(base_path)
  path_name <- tools::file_path_sans_ext(base_path)
  i <- 1
  new_path <- base_path
  while (file.exists(new_path)) {
    new_name <- paste0(path_name, "_", i, ".", path_ext)
    new_path <- file.path(path_dir, basename(new_name)) 
    i <- i + 1
  }
  return(new_path)
}

detect_label_col <- function(df) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  fr <- cols[grepl("label(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  if ("label" %in% cols_clean) return(cols[cols_clean == "label"])
  ll <- cols[grepl("label(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  NA_character_
}

detect_hint_col <- function(df) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  fr <- cols[grepl("hint(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  if ("hint" %in% cols_clean) return(cols[cols_clean == "hint"])
  ll <- cols[grepl("hint(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  NA_character_
}

library(dplyr)
library(stringr)
# Assurez-vous que la fonction detect_label_col est définie quelque part avant d'exécuter ceci

# Assurez-vous que dplyr et stringr sont chargés : library(dplyr); library(stringr)

translate_relevant <- function(expr, labels, choices, df_survey) {
  if (is.null(expr) || is.na(expr) || !nzchar(trimws(expr))) return(NA_character_)
  txt <- expr
  
  # Nettoyage initial des dataframes (au cas où ils ne seraient pas propres)
  choices <- choices %>% mutate(
    name = tolower(str_trim(as.character(name))),
    list_name = tolower(str_trim(as.character(list_name)))
  )
  df_survey <- df_survey %>% mutate(
    name = tolower(str_trim(as.character(name)))
  )
  
  get_choice_label <- function(code, list_name){
    code_clean <- tolower(str_trim(as.character(code)))
    list_name_clean <- tolower(str_trim(as.character(list_name)))
    r <- choices %>% filter(name == code_clean & list_name == list_name_clean)
    if (nrow(r) > 0) {
      label_col_name_choices <- detect_label_col(choices)
      if (!is.na(label_col_name_choices)) {
        lab_val <- as.character(r[[label_col_name_choices]]) 
        lab <- lab_val %||% as.character(code)
      } else {
        lab <- as.character(r$name) %||% as.character(code)
      }
      return(lab)
    } else { return(as.character(code)) }
  }
  
  # Fonction pour obtenir le label de la variable (Nettoyage HTML inclus)
  get_var_label <- function(v){ 
    val <- tryCatch(labels[[v]], error = function(e) NULL)
    if(is.null(val) || is.na(val) || !nzchar(val)) {
      clean_val <- as.character(v)
    } else {
      # Supprimez tout le HTML ici (<.*?>)
      clean_val <- str_replace_all(as.character(val), "<.*?>", "")
    }
    return(clean_val) 
  }
  
  get_listname_from_varname <- function(var_name, survey_df) {
    var_name_clean <- tolower(str_trim(as.character(var_name)))
    row <- survey_df %>% filter(name == var_name_clean)
    if (nrow(row) > 0) {
      type_val <- tolower(as.character(row$type))
      list_name <- stringr::str_replace(type_val, "^select_(one|multiple)\\s+", "")
      if (list_name == type_val) { return(NA_character_) }
      return(str_trim(list_name))
    } else { return(NA_character_) }
  }
  
  # --- Substitutions principales ---
  # Le reste du code effectue les remplacements d'opérateurs et de fonctions.
  
  # Remplacement pour 'selected()'
  txt <- stringr::str_replace_all(txt, "selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+'\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("Pour '%s', l'option '%s' est sélectionnée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  # Remplacement pour 'not(selected())'
  txt <- stringr::str_replace_all(txt, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+?'\\s*\\)\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("Pour '%s', l'option '%s' n'est PAS sélectionnée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  # Remplacement des variables simples (${nom_variable}) par leur label
  txt <- stringr::str_replace_all(txt, "\\$\\{[^}]+\\}", function(x){ 
    m <- stringr::str_match(x, "\\$\\{([^}]+)\\}")
    vars <- m[,2]
    vapply(vars, get_var_label, character(1)) 
  })
  
  # Remplacement des opérateurs logiques et mathématiques
  # CES REMPLACEMENTS NE DEVRAIENT S'APPLIQUER QU'À LA LOGIQUE RELEVANT, PAS AUX LABELS DE QUESTIONS
  txt <- stringr::str_replace_all(txt, "\\bandand\\b", " et "); 
  txt <- stringr::str_replace_all(txt, "\\bor\\b",  " ou "); 
  txt <- stringr::str_replace_all(txt, "\\bnot\\b", " non "); 
  
  txt <- stringr::str_replace_all(txt, "\\s*=\\s*", " est égal à ");
  txt <- stringr::str_replace_all(txt, "\\s*!=\\s*", " est différent de ");
  txt <- stringr::str_replace_all(txt, "\\s*>\\s*", " est supérieur à ");
  txt <- stringr::str_replace_all(txt, "\\s*>=\\s*", " est supérieur ou égal à ");
  txt <- stringr::str_replace_all(txt, "\\s*<\\s*", " est inférieur à ");
  txt <- stringr::str_replace_all(txt, "\\s*<=\\s*", " est inférieur ou égal à ");
  
  # Note : Les symboles '*' et '?' ne sont pas traduits ici, 
  # mais dans votre exemple, ils ont été traduits car ils étaient dans la mauvaise fonction.
  
  # Remplacement des valeurs binaires Oui/Non
  txt <- stringr::str_replace_all(txt, "'1'", "'Oui'"); 
  txt <- stringr::str_replace_all(txt, "'0'", "'Non'"); 
  
  # Gère count-selected >= 1
  txt <- stringr::str_replace_all(txt, "count-selected\\s*\\(\\s*[^\\)]+\\)\\s*>=\\s*1", function(x){ 
    m <- stringr::str_match(x, "count-selected\\s*\\(\\s*(.+?)\\s*\\)\\s*>=\\s*1")
    v <- m[,2]
    vapply(v, function(z) sprintf("Au moins une option est cochée pour %s", z), character(1)) 
  })
  
  # Nettoyage final pour une meilleure lisibilité
  txt <- stringr::str_replace_all(txt, "\\(\\s*", "\\(");
  txt <- stringr::str_replace_all(txt, "\\s*\\)", "\\)");
  
  return(txt)
}




# ---------------------------------------------------------------------
# Styles texte & paragraphe / Blocs visuels 
# ---------------------------------------------------------------------
fp_txt <- function(color="black", size=11, bold=FALSE, italic=FALSE, underline=FALSE) { fp_text(color = color, font.size = size, bold = bold, italic = italic, underline = underline, font.family = FONT_FAMILY) }
fp_q_blue   <- fp_txt(color = WFP_DARK_BLUE, size = FS_Q_BLUE, bold = FALSE) 
fp_q_grey   <- fp_txt(color = GREY_TXT, size = FS_Q_GREY)             
fp_sec_title <- fp_txt(color = WFP_BLUE, size = FS_TITLE, bold = TRUE, underline = TRUE)
fp_sub_title <- fp_txt(color = WFP_BLUE, size = FS_SUBTITLE, bold = TRUE, underline = TRUE)
fp_block     <- fp_txt(color = WFP_BLUE, size = FS_BLOCK,  bold = TRUE) 
fp_meta      <- fp_txt(color = GREY_TXT, size = FS_META)
fp_hint      <- fp_txt(color = GREY_TXT, size = FS_HINT, italic = TRUE)
fp_relevant  <- fp_txt(color = RED_TXT,  size = FS_RELV)
fp_missing_list <- fp_txt(color = GREY_TXT, size = FS_MISC, italic = TRUE)
p_default <- fp_par(text.align = "left", line_spacing = LINE_SP)  
p_q_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_Q * 72)
p_c_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_C * 72)

names_deja_traites <- character(0)

add_hrule <- function(doc, width = 1){ doc <- body_add_par(doc, ""); return(doc) }
add_band <- function(doc, text, txt_fp){ doc <- body_add_fpar(doc, fpar(ftext(text, txt_fp), fp_p = fp_par(padding.top = 4, padding.bottom = 4, padding.left = INDENT_Q*72))); return(doc) }

add_choice_lines <- function(doc, choices_map, list_name_to_filter, symbol = "○") {
  if (is.na(list_name_to_filter) || list_name_to_filter == "") return(doc)
  list_name_to_filter_clean <- str_replace_all(list_name_to_filter, "\\s", "")
  
  df <- choices_map[[list_name_to_filter_clean]]
  
  if (is.null(df) || nrow(df) == 0) {
    doc <- body_add_fpar(doc, fpar(ftext(glue("Commentaire: La liste de choix '{list_name_to_filter}' est introuvable ou vide dans l'onglet 'choices'."), fp_missing_list), fp_p = p_c_indent_fixed)) 
    return(doc)
  }
  
  df <- df %>% mutate(
    label_final = coalesce(label_col, as.character(name)),
    txt_final = sprintf("%s %s (%s)", symbol, label_final, name)
  )
  
  purrr::walk(df$txt_final, function(text_label) { 
    doc <<- body_add_fpar(doc, fpar(ftext(text_label, fp_txt(size = FS_MISC)), fp_p = p_c_indent_fixed)) 
  }); 
  return(doc)
}

add_placeholder_box <- function(doc, txt = "Réponse : [insérer votre réponse ici]"){
  doc <- body_add_fpar(doc, fpar(ftext(txt, fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_c_indent_fixed))
  doc <- body_add_par(doc, ""); return(doc)
}

# ---------------------------------------------------------------------
# Rendu d’une question
# ---------------------------------------------------------------------
render_question <- function(doc, row, number, label_col_name, hint_col_name, choices_map, lab_map, full_choices_sheet,full_survey_sheet){
  if (is.null(row) || nrow(row) == 0) return(doc)
  q_type <- tolower(row$type %||% "")
  q_name <- row$name %||% ""
  q_lab  <- str_replace_all(row[[label_col_name]] %||% q_name, "<.*?>", "") 
  
  if (is.na(q_name) || q_name %in% EXCLUDE_NAMES) return(doc)
  
  if (q_type == "note") {
    ftext_blue <- ftext(sprintf("Note : %s", q_lab), fp_q_blue)
  } else {
    if (is.null(number) || is.na(number) || !nzchar(as.character(number))) return(doc)
    ftext_blue <- ftext(sprintf("%s. %s", number, q_lab), fp_q_blue)
  }
  ftext_grey <- ftext(sprintf(" (%s – %s)", q_name, q_type), fp_q_grey)
  doc <- body_add_fpar(doc, fpar(ftext_blue, ftext_grey, fp_p = p_q_indent_fixed)) 
  
  rel <- row$relevant %||% NA_character_
  h <- NA_character_
  if (!is.na(hint_col_name) && hint_col_name %in% names(row)) { h <- row[[hint_col_name]] %||% NA_character_ }
  
  if (!is.na(rel) && nzchar(rel)) { tr <- translate_relevant(rel, lab_map, full_choices_sheet, full_survey_sheet); doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) }
  if (!is.na(h) && nzchar(h)) { doc <- body_add_fpar(doc, fpar(ftext(h, fp_hint), fp_p = p_q_indent_fixed)) }
  
  if (str_starts(q_type, "select_one")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir la réponse parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_one\\s+", "", q_type));
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "○");
  } 
  else if (str_starts(q_type, "select_multiple")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir les réponses pertinentes parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_multiple\\s+", "", q_type)); 
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "☐"); 
  } 
  else if (str_detect(q_type, "^note")) { } 
  else { 
    placeholder <- "Réponse : [insérer votre réponse ici]"
    if (str_detect(q_type, "integer")) { placeholder <- "Réponse : [insérer un entier]" } 
    else if (str_detect(q_type, "decimal")) { placeholder <- "Réponse : [insérer un décimal]" }
    else if (str_detect(q_type, "date")) { placeholder <- "Réponse : [insérer une date]" }
    else if (str_detect(q_type, "geopoint")) { placeholder <- "Réponse : [capturer les coordonnées GPS]" }
    else if (str_detect(q_type, "image")) { placeholder <- "Réponse : [prendre une photo]" }
    
    doc <- add_placeholder_box(doc, placeholder)
  }
  
  if (!str_detect(q_type, "^note|select_")) {
    doc <- body_add_fpar(doc, fpar(ftext("________________________________________", fp_txt(size=FS_Q_BLUE)), fp_p = p_c_indent_fixed))
  }
  
  doc <- body_add_par(doc, "")
  return(doc)
}

# ---------------------------------------------------------------------
# Fonction principale 
# ---------------------------------------------------------------------

xlsform_to_wordRev <- function(xlsx = XLSFORM_PATH, output_dir = OUTPUT_DIR, template_docx = TEMPLATE_DOCX, logo_path = LOGO_PATH, doc_title = NULL) {
  message(glue("--- Démarrage du processus de génération Word ---"))
  if (is.null(xlsx)) {
    message("Veuillez sélectionner le fichier XLSForm (.xlsx)...")
    tryCatch({ xlsx <- file.choose() }, error = function(e) { stop("Sélection du fichier annulée ou échouée.") })
    message(glue("Fichier sélectionné : {basename(xlsx)}"))
  }
  
  if (!file.exists(xlsx)) stop("Fichier XLSForm introuvable : ", xlsx)
  message(glue("Lecture de l'XLSForm depuis: {basename(xlsx)}"))
  
  # Les dataframes sont lus ici :
  survey   <- read_excel(xlsx, sheet = "survey")
  choices  <- read_excel(xlsx, sheet = "choices")
  settings <- tryCatch(read_excel(xlsx, sheet = "settings"), error = function(e) NULL)
  
  names(choices) <- tolower(iconv(names(choices), to = "ASCII//TRANSLIT"))
  
  if (is.null(doc_title)) { doc_title <- if (!is.null(settings) && "form_title" %in% names(settings)) settings$form_title %||% "XLSForm – Rendu Word" else "XLSForm – Rendu Word" }
  
  base_filename <- tools::file_path_sans_ext(basename(xlsx))
  safe_title <- str_replace_all(doc_title, "[[:punct:]\\s]+", "_")
  output_filename <- glue("{safe_title}.docx")
  out_docx <- file.path(output_dir, output_filename)
  
  out_docx <- generate_output_path(out_docx)
  if (!dir.exists(output_dir)) { dir.create(output_dir, recursive = TRUE) }
  
  message(glue("Le document sera enregistré sous: {out_docx}"))
  message(glue("Ouverture du document Word. Utilisation du titre: {doc_title}"))
  
  label_col_name <- detect_label_col(survey)
  if (is.na(label_col_name)) stop("Colonne label introuvable.")
  label_col_sym <- sym(label_col_name) 
  hint_col_name <- detect_hint_col(survey)
  
  survey <- survey %>% filter(!duplicated(name))
  if (!"list_name" %in% names(choices)) { stop("La feuille 'choices' doit contenir 'list_name'.") }
  lab_map <- survey %>% select(name, !!label_col_sym) %>% mutate(name = as.character(name)) %>% tibble::deframe()
  
  doc <- if (!is.null(template_docx) && file.exists(template_docx)) read_docx(template_docx) else read_docx()
  doc <- body_set_default_section(doc, prop_section(page_size = page_size(orient = "portrait", width = 8.5, height = 11), page_margins = page_mar(top = 1, bottom = 1, left = 1.0, right = 1.0, header = 0.5, footer = 0.5)))
  
  fp_title_main <- fp_txt(color = RED_TXT, size = 16, bold = TRUE)
  p_title_main <- fp_par(text.align = "center", line_spacing = LINE_SP)
  
  if (!is.null(logo_path) && file.exists(logo_path)) { 
    doc <- body_add_par(doc, "", style = "Normal")
    doc <- body_add_fpar(doc, fpar(external_img(src = logo_path, height = 0.6, width = 0.6, unit = "in"), ftext("  "), ftext(doc_title, prop = fp_title_main), fp_p = p_title_main)) 
  } else { 
    doc <- body_add_fpar(doc, fpar(ftext(doc_title, fp_title_main), fp_p = p_title_main)) 
  }
  doc <- add_hrule(doc)
  
  choices <- choices %>% 
    mutate_all(as.character) %>% 
    mutate_all(~ifelse(is.na(.), NA_character_, .)) %>%
    mutate(list_name = as.character(str_trim(str_replace_all(tolower(list_name), "[[:space:]]+", ""))))
  
  label_col_choices <- detect_label_col(choices)
  if (is.na(label_col_choices)) stop("Colonne label introuvable dans l'onglet 'choices'.")
  
  choices_map <- split(choices, choices$list_name)
  choices_map <- lapply(choices_map, function(df) {
    df %>% mutate(label_col = .data[[label_col_choices]])
  })
  
  level_stack <- list()
  sec_id <- 0L; sub_id <- 0L; q_id <- 0L; names_deja_traites <- character(0) 
  
  message("Début de l'analyse des questions et de la génération du corps du document...")
  for (i in seq_len(nrow(survey))) {
    r <- survey[i, , drop = FALSE]
    # ... (le reste de la boucle pour définir t, qname, etc.) ...
    t_raw <- as.character(r$type %||% "")
    t <- tolower(t_raw)
    qname <- as.character(r$name)
    if (!is.na(qname) && nzchar(qname)) { if (qname %in% names_deja_traites) { next } else { names_deja_traites <- c(names_deja_traites, qname) } }
    
    if (t %in% EXCLUDE_TYPES) next
    if (!is.null(r$name) && any(tolower(r$name) %in% tolower(EXCLUDE_NAMES))) next
    current_number <- NULL
    is_group <- str_starts(t, "begin_group") && !str_starts(t, "begin_repeat")
    is_repeat <- str_starts(t, "begin_repeat") || str_starts(t, "begin repeat")
    is_end <- str_starts(t, "end_group") || str_starts(t, "end group") || str_starts(t, "end_repeat") || str_starts(t, "end repeat")
    
    if (is_group || is_repeat) {
      lbl <- r[[label_col_name]] %||% r$name %||% ""
      
      if (is_group) {
        prev_type <- if (i > 1) tolower(as.character(survey$type[i - 1] %||% "")) else ""
        prev_is_group <- str_starts(prev_type, "begin_group") || str_starts(prev_type, "begin repeat")
        if ( !prev_is_group) { sec_id <- sec_id + 1L; sub_id <- 0L; q_id <- 0L; message(glue("-> Génération Section {sec_id}: {lbl}")); doc <- add_band(doc, glue("Section {sec_id} : {lbl}"), txt_fp = fp_sec_title) } else { sub_id <- sub_id + 1L; q_id <- 0L; message(glue("--> Génération Sous-section {sec_id}.{sub_id}: {lbl}")); doc <- add_band(doc, glue("Sous-section {sec_id}.{sub_id} : {lbl}"), txt_fp = fp_sub_title) }
      } else if (is_repeat) { message(glue("--> Génération Bloc Répétitif: {lbl}")); doc <- body_add_fpar(doc, fpar(ftext(glue("🔁 Bloc : {lbl}"), fp_block), fp_p = p_default)) }
      rel <- r$relevant %||% NA_character_
      if (!is.na(rel) && nzchar(rel)) { 
        # Traduction de l'expression 'relevant' du groupe
        tr <- translate_relevant(rel, lab_map, choices, survey); 
        doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) 
      }
      doc <- add_hrule(doc) ; next
    }
    if (is_end) next
    
    if(t != "note") q_id <- q_id + 1L
    
    current_number <- NULL
    if(t != "note") {
      current_number <- glue("{q_id}")
      if (sub_id > 0) { current_number <- glue("{sec_id}.{sub_id}.{q_id}") } else if (sec_id > 0) { current_number <- glue("{sec_id}.{q_id}") }
    }
    
    # APPEL CORRIGÉ : utilise 'choices' et 'survey' (qui existent localement)
    doc <- render_question(doc, r, current_number, label_col_name, hint_col_name, choices_map, full_choices_sheet = choices, lab_map = lab_map, full_survey_sheet = survey)
  }
  
  print(doc, target = out_docx)
  message(glue("--- Processus terminé ---"))
  final_path_display <- tryCatch(normalizePath(out_docx, winslash = '/'), error = function(e) out_docx)
  message(glue("✅ Document généré : {final_path_display}"))
  invisible(out_docx)
}
