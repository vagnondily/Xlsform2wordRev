# XLSForm vers Word (Rendu Professionnel)

Ce projet fournit un script R (`xlsform_to_word.R`) qui automatise la conversion des fichiers de définition de formulaire XLSForm (utilisés par des plateformes comme ODK, KoboToolbox) en un document Microsoft Word (.docx) lisible et bien formaté.

Le script est conçu pour générer une documentation de haute qualité du questionnaire, incluant :

*   Numérotation hiérarchique des sections et questions.
*   Formatage spécifique pour les différents types de questions (texte, numérique, date, sélection).
*   Inclusion visuelle des listes de choix (options).
*   Traduction des logiques de pertinence (`relevant logic`).
*   Personnalisation des styles au besoin.

## Prérequis

Le script nécessite R ainsi que les packages suivants. Vous pouvez les installer via la console R :

```R
install.packages(c("readxl", "dplyr", "stringr", "tidyr", "purrr", "officer", "flextable", "glue", "tools", "tibble", "rlang"))

## Installation (GitHub)

```r
install.packages("remotes")
remotes::install_github("vagnondily/Xlsform2wordRev")

```

## Installation (locale)

```r
install.packages("remotes")
remotes::install_local("/chemin/vers/xlsform2wordRev", upgrade = "never")
```

> Sous Windows et R ≥ 4.5, l’installation depuis les sources nécessite **Rtools45**.

## Utilisation

```R
library(xlsform2wordRev)
xlsform_to_wordRev()

```
## Supprimer complètement le package (au besoin)
```R
remove.packages("xlsform2wordRev")
.rs.restartR()  # dans RStudio

```
