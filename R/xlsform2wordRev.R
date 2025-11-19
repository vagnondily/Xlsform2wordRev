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
OUTPUT_DIR <- file.path(path.expand("~"), "Downloads")


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

translate_relevant <- function(expr, labels, choices) {
  if (is.null(expr) || is.na(expr) || !nzchar(trimws(expr))) return(NA_character_)
  txt <- expr
  
  get_choice_label <- function(code){
    # Filtrer spécifiquement pour le code donné
    r <- choices %>% filter(name == code)
    
    if (nrow(r) > 0) {
      label_col_name_choices <- detect_label_col(choices)
      if (!is.na(label_col_name_choices)) {
        # Correction V2: Extrait la première valeur de la colonne ciblée explicitement
        lab_val <- as.character(r[[label_col_name_choices]])[1] 
        lab <- lab_val %||% as.character(code)
      } else {
        lab <- as.character(r$name)[1] %||% as.character(code)
      }
      return(lab) # Renvoie une seule valeur de longueur 1
    } else return(as.character(code))
  }
  
  get_var_label <- function(v){ 
    val <- tryCatch(labels[[v]], error = function(e) NULL)
    if(is.null(val) || is.na(val) || !nzchar(val)) return(as.character(v)) else return(as.character(val)) 
  }
  
  # Les fonctions de remplacement utilisent vapply qui gère correctement les vecteurs si les sous-fonctions sont scalaires
  # 1. Traitement standard de selected() et not(selected()) (pour select_multiple)
  txt <- stringr::str_replace_all(txt, "selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+'\\s*\\)", function(x){ m <- stringr::str_match(x, "selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)"); vars <- m[,2]; codes <- m[,3]; vapply(seq_along(vars), function(i){ sprintf("%s contient '%s'", get_var_label(vars[i]), get_choice_label(codes[i]))}, character(1)) })
  txt <- stringr::str_replace_all(txt, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+?'\\s*\\)\\s*\\)", function(x){ m <- stringr::str_match(x, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)\\s*\\)"); vars <- m[,2]; codes <- m[,3]; vapply(seq_along(vars), function(i){ sprintf("%s ne contient pas '%s'", get_var_label(vars[i]), get_choice_label(codes[i]))}, character(1)) })
  
  # 2. Remplacement des noms de variables (${...}) par leur label
  txt <- stringr::str_replace_all(txt, "\\$\\{([^}]+)\\}", function(x){ m <- stringr::str_match(x, "\\$\\{([^}]+)\\}"); vars <- m[,2]; vapply(vars, get_var_label, character(1)) })
  
  # 3.Remplacement des codes de choix seuls (ex: '= '1'', '!= '0'') par leur label (pour select_one)
  # Cette étape gère les comparaisons simples.
  txt <- stringr::str_replace_all(txt, "'([^']+)'", function(x){ m <- stringr::str_match(x, "'([^']+)'"); code <- m[,2]; vapply(code, get_choice_label, character(1)) })

  # 4. Traitement des opérateurs logiques et autres expressions
  txt <- stringr::str_replace_all(txt, "\\bandand\\b", "et"); 
  txt <- stringr::str_replace_all(txt, "\\bor\\b",  "ou"); 
  txt <- stringr::str_replace_all(txt, "\\bnot\\b", "non"); 
  
  # Note: Les remplacements spécifiques '1'/'0' ne sont pas nécessaires
  # car '1' et '0' seront maintenant traduits par 'Oui' et 'Non' via get_choice_label() s'ils existent dans vos choix.
  # Je les laisse commentés au cas où vous en auriez besoin.
  # txt <- stringr::str_replace_all(txt, "=\\s*'1'",  " = 'Oui'"); txt <- stringr::str_replace_all(txt, "=\\s*'0'",  " = 'Non'"); txt <- stringr::str_replace_all(txt, "!=\\s*'1'", " ≠ 'Oui'"); txt <- stringr::str_replace_all(txt, "!=\\s*'0'", " ≠ 'Non'"); 
  
  txt <- stringr::str_replace_all(txt, "count-selected\\s*\\(\\s*[^\\)]+\\)\\s*>=\\s*1", function(x){ m <- stringr::str_match(x, "count-selected\\s*\\(\\s*(.+?)\\s*\\)\\s*>=\\s*1"); v <- m[,2]; vapply(v, function(z) sprintf("au moins une option cochée pour %s", z), character(1)) })
  txt
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
render_question <- function(doc, row, number, label_col_name, hint_col_name, choices_map, lab_map, full_choices_sheet){
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
  
  if (!is.na(rel) && nzchar(rel)) { tr <- translate_relevant(rel, lab_map, full_choices_sheet); doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) }
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
    doc <- body_add_fpar(doc, fpar(ftext("___________________________________________________________________", fp_txt(size=FS_Q_BLUE)), fp_p = p_c_indent_fixed))
  }
  
  doc <- body_add_par(doc, "")
  return(doc)
}

# ---------------------------------------------------------------------
# Fonction principale 
# ---------------------------------------------------------------------

xlsform_to_wordRev <- function(xlsx = NULL, output_dir = NULL, template_docx = NULL, logo_path = NULL, doc_title = NULL) {
  message(glue("--- Démarrage du processus de génération Word ---"))
  
  # --- 1. Gestion des paramètres par défaut si non fournis ---
  # Utilise les variables globales définies dans le script si les arguments sont NULL
  xlsx_path_final    <- xlsx %||% XLSFORM_PATH
  output_dir_final   <- output_dir %||% OUTPUT_DIR %||% file.path(path.expand("~"), "Downloads")
  template_docx_final<- template_docx %||% TEMPLATE_DOCX
  logo_path_final    <- logo_path %||% LOGO_PATH

  # Nous allons utiliser ces variables suffixées "_final" pour éviter toute confusion.
  
  # Assurez-vous que le répertoire de sortie existe
  if (!dir.exists(output_dir_final)) { dir.create(output_dir_final, recursive = TRUE) }
  
  # Sélection fichier si non fourni via argument ou variable globale
  if (is.null(xlsx_path_final)) {
    message("Veuillez sélectionner le fichier XLSForm (.xlsx)...")
    tryCatch({ xlsx_path_final <- file.choose() }, error = function(e) { stop("Sélection du fichier annulée ou échouée.") })
    message(glue("Fichier sélectionné : {basename(xlsx_path_final)}"))
  }
  
  if (!file.exists(xlsx_path_final)) stop("Fichier XLSForm introuvable : ", xlsx_path_final)
  message(glue("Lecture de l'XLSForm depuis: {basename(xlsx_path_final)}"))
  
  survey   <- read_excel(xlsx_path_final, sheet = "survey")
  choices  <- read_excel(xlsx_path_final, sheet = "choices")
  settings <- tryCatch(read_excel(xlsx_path_final, sheet = "settings"), error = function(e) NULL)
  
  names(choices) <- tolower(iconv(names(choices), to = "ASCII//TRANSLIT"))
  
  if (is.null(doc_title)) { doc_title <- if (!is.null(settings) && "form_title" %in% names(settings)) settings$form_title %||% "XLSForm – Rendu Word" else "XLSForm – Rendu Word" }
  
  base_filename <- tools::file_path_sans_ext(basename(xlsx_path_final))
  safe_title <- str_replace_all(doc_title, "[[:punct:]\\s]+", "_")
  output_filename <- glue("{safe_title}.docx")
  
  # Utilisez output_dir_final pour construire le chemin complet du fichier de sortie
  out_docx <- file.path(output_dir_final, output_filename)
  
  out_docx <- generate_output_path(out_docx)
  
  message(glue("Le document sera enregistré sous: {out_docx}"))
  message(glue("Ouverture du document Word. Utilisation du titre: {doc_title}"))
  
  # ... (Le reste de la fonction reste identique jusqu'à la fin) ...
  
  print(doc, target = out_docx)
  message(glue("--- Processus terminé ---"))
  final_path_display <- tryCatch(normalizePath(out_docx, winslash = '/'), error = function(e) out_docx)
  message(glue("✅ Document généré : {final_path_display}"))
  
  # --- Code pour ouvrir automatiquement le dossier d'output ---
  
  # Détecte le système d'exploitation et utilise la commande appropriée
  if (Sys.info()["sysname"] == "Windows") {
    # Utilisez output_dir_final ici
    shell.exec(output_dir_final) 
  } else if (Sys.info()["sysname"] == "Darwin") {
    system(glue("open {shQuote(output_dir_final)}"))
  } else if (.Platform$OS.type == "unix") {
    system(glue("xdg-open {shQuote(output_dir_final)}"))
  }
  # -----------------------------------------------------------
  invisible(out_docx)
}

