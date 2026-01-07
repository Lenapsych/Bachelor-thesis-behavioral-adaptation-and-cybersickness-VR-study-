library(tidyverse)
library(readxl)
library(janitor)
library(lmtest)
library(broom)
library(writexl)
library(patchwork)
library(ggplot2)

# Sample laden 
sample_h3 <- read_csv("sample_h3_export.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

sample_h3 %>% summarise(n = n(), n_codes = n_distinct(code)) %>% print()

# GEQ laden, Skalen bilden
geq_raw <- read_excel("Let's try again - GEQ_fertig_h3.xlsx") %>%
  clean_names() %>%
  mutate(code = as.character(code))

geq_num <- geq_raw %>%
  mutate(across(starts_with("geq_"),
                ~ as.numeric(gsub(",", ".", as.character(.x)))))

# Skalen pro Zeile
geq_with_scales <- geq_num %>%
  rowwise() %>%
  mutate(
    geq_immersion  = mean(c_across(starts_with("geq_imm")),     na.rm = TRUE),
    geq_pos_affect = mean(c_across(starts_with("geq_pos_aff")), na.rm = TRUE),
    geq_neg_affect = mean(c_across(starts_with("geq_neg_aff")), na.rm = TRUE),
    geq_challenge  = mean(c_across(starts_with("geq_chall")),   na.rm = TRUE)
  ) %>%
  ungroup()

# Skalen pro Person
geq_scales <- geq_with_scales %>%
  group_by(code) %>%
  summarise(
    geq_immersion  = mean(geq_immersion,  na.rm = TRUE),
    geq_pos_affect = mean(geq_pos_affect, na.rm = TRUE),
    geq_neg_affect = mean(geq_neg_affect, na.rm = TRUE),
    geq_challenge  = mean(geq_challenge,  na.rm = TRUE),
    .groups = "drop"
  )

geq_scales %>% summarise(n = n(), n_codes = n_distinct(code)) %>% print()


# Sample + GEQ + SSQ zusammenführen
daten_h3 <- sample_h3 %>%
  inner_join(geq_scales, by = "code") %>%
  filter(
    !is.na(ssq_post_x),
    !is.na(ssq_post_v2),
    !is.na(geq_immersion),
    !is.na(geq_pos_affect),
    !is.na(geq_neg_affect),
    !is.na(geq_challenge)
  )

daten_h3 %>% summarise(n = n(), n_codes = n_distinct(code)) %>% print()

valid_codes <- daten_h3$code

# behaviorale Adaptions-Daten einlesen
blocks_keep <- c("MISC 1", "MISC 2", "MISC 3", "MISC 4", "MISC 5")

v2 <- read_csv("V2_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

x <- read_csv("X_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

v2_perf <- v2 %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    interval_num = as.numeric(gsub("MISC ", "", interval)),
    treffer_v2   = c_collected,
    versuche_v2  = triggers,
    perf_v2      = treffer_v2 / pmax(versuche_v2, 1)
  ) %>%
  select(code, interval_num, perf_v2) %>%
  filter(code %in% valid_codes)

x_perf <- x %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    interval_num = as.numeric(gsub("MISC ", "", interval)),
    treffer_x    = t_hit,
    versuche_x   = shoot,
    perf_x       = treffer_x / pmax(versuche_x, 1)
  ) %>%
  select(code, interval_num, perf_x) %>%
  filter(code %in% valid_codes)

# Check: hat jede Person 5 Blöcke?
v2_perf %>% count(code) %>% count(n) %>% print()
x_perf  %>% count(code) %>% count(n) %>% print()

# Performanzsteigerung (Slope) pro Person berechnen
calc_slope <- function(df, perf_col, slope_name) {
  df %>%
    group_by(code) %>%
    summarise(
      !!slope_name := {
        d <- cur_data()
        # nur berechnen, wenn mind. 2 Punkte vorhanden
        if (sum(!is.na(d[[perf_col]])) >= 2) {
          coef(lm(d[[perf_col]] ~ d$interval_num))[2]
        } else {
          NA_real_
        }
      },
      .groups = "drop"
    )
}

slope_v2 <- calc_slope(v2_perf, "perf_v2", rlang::sym("slope_v2"))
slope_x  <- calc_slope(x_perf,  "perf_x",  rlang::sym("slope_x"))

# an daten_h3 hängen
daten_h3 <- daten_h3 %>%
  left_join(slope_v2, by = "code") %>%
  left_join(slope_x,  by = "code")

daten_h3 %>% summarise(
  n = n(),
  n_slope_v2 = sum(!is.na(slope_v2)),
  n_slope_x  = sum(!is.na(slope_x))
) %>% print()

# 6) Deskriptive Statistik (Slope & Prädiktoren)
des_h3 <- daten_h3 %>%
  summarise(
    n               = n(),
    slope_x_m       = mean(slope_x, na.rm = TRUE),
    slope_x_sd      = sd(slope_x,   na.rm = TRUE),
    slope_v2_m      = mean(slope_v2, na.rm = TRUE),
    slope_v2_sd     = sd(slope_v2,   na.rm = TRUE),
    ssq_x_m         = mean(ssq_post_x,  na.rm = TRUE),
    ssq_x_sd        = sd(ssq_post_x,    na.rm = TRUE),
    ssq_v2_m        = mean(ssq_post_v2, na.rm = TRUE),
    ssq_v2_sd       = sd(ssq_post_v2,   na.rm = TRUE),
    imm_m           = mean(geq_immersion,  na.rm = TRUE),
    imm_sd          = sd(geq_immersion,    na.rm = TRUE),
    pos_m           = mean(geq_pos_affect, na.rm = TRUE),
    pos_sd          = sd(geq_pos_affect,   na.rm = TRUE),
    neg_m           = mean(geq_neg_affect, na.rm = TRUE),
    neg_sd          = sd(geq_neg_affect,   na.rm = TRUE),
    chall_m         = mean(geq_challenge,  na.rm = TRUE),
    chall_sd        = sd(geq_challenge,    na.rm = TRUE)
  )

des_h3 %>% print()

# Zentrierung (für Interaktionen)

daten_h3 <- daten_h3 %>%
  mutate(
    imm_c   = scale(geq_immersion,   center = TRUE, scale = FALSE)[,1],
    pos_c   = scale(geq_pos_affect,  center = TRUE, scale = FALSE)[,1],
    neg_c   = scale(geq_neg_affect,  center = TRUE, scale = FALSE)[,1],
    chall_c = scale(geq_challenge,   center = TRUE, scale = FALSE)[,1],
    ssq_x_c  = scale(ssq_post_x,  center = TRUE, scale = FALSE)[,1],
    ssq_v2_c = scale(ssq_post_v2, center = TRUE, scale = FALSE)[,1]
  )

# Datensplits pro Spiel (nur Fälle mit Outcome)
daten_x  <- daten_h3 %>% filter(!is.na(slope_x),  !is.na(ssq_x_c))
daten_v2 <- daten_h3 %>% filter(!is.na(slope_v2), !is.na(ssq_v2_c))


# Regressionen (H3): Slope ~ SSQ * GEQ

# Ballwurfspiel X
model_h3_x_imm   <- lm(slope_x ~ ssq_x_c * imm_c,   data = daten_x)
model_h3_x_pos   <- lm(slope_x ~ ssq_x_c * pos_c,   data = daten_x)
model_h3_x_neg   <- lm(slope_x ~ ssq_x_c * neg_c,   data = daten_x)
model_h3_x_chall <- lm(slope_x ~ ssq_x_c * chall_c, data = daten_x)

# Hexenspiel V2
model_h3_v2_imm   <- lm(slope_v2 ~ ssq_v2_c * imm_c,   data = daten_v2)
model_h3_v2_pos   <- lm(slope_v2 ~ ssq_v2_c * pos_c,   data = daten_v2)
model_h3_v2_neg   <- lm(slope_v2 ~ ssq_v2_c * neg_c,   data = daten_v2)
model_h3_v2_chall <- lm(slope_v2 ~ ssq_v2_c * chall_c, data = daten_v2)

# Kurz-Output
summary(model_h3_x_imm)
summary(model_h3_v2_imm)

# Annahmeprüfungen (Residual-Normalität + Homoskedastizität)

check_assumptions <- function(model, model_name = "") {
  cat("\n=========================================\n")
  cat(" Annahmprüfung für:", model_name, "\n")
  cat("=========================================\n\n")
  
  sh <- shapiro.test(residuals(model))
  cat("Shapiro-Wilk (Residuen):\n"); print(sh); cat("\n")
  
  bp <- bptest(model)
  cat("Breusch-Pagan (Homoskedastizität):\n"); print(bp); cat("\n")
  
  # Diagnoseplots
  oldpar <- par(no.readonly = TRUE)
  par(mfrow = c(2,2))
  plot(model, main = model_name)
  par(oldpar)
  
  cat("\n--- ENDE MODELL:", model_name, " ---\n\n")
}

# Annahmen X
check_assumptions(model_h3_x_imm,   "Ballwurf X – SSQ×Immersion (Slope)")
check_assumptions(model_h3_x_pos,   "Ballwurf X – SSQ×PosAffekt (Slope)")
check_assumptions(model_h3_x_neg,   "Ballwurf X – SSQ×NegAffekt (Slope)")
check_assumptions(model_h3_x_chall, "Ballwurf X – SSQ×Challenge (Slope)")

# Annahmen V2
check_assumptions(model_h3_v2_imm,   "Hexe V2 – SSQ×Immersion (Slope)")
check_assumptions(model_h3_v2_pos,   "Hexe V2 – SSQ×PosAffekt (Slope)")
check_assumptions(model_h3_v2_neg,   "Hexe V2 – SSQ×NegAffekt (Slope)")
check_assumptions(model_h3_v2_chall, "Hexe V2 – SSQ×Challenge (Slope)")


# Interaktionsplots 
make_interaction_plot_3lines <- function(data, model, x_var, mod_var,
                                         title, x_lab, y_lab, legend_title) {
  
  # Moderator: Mittelwert & SD
  m  <- mean(data[[mod_var]], na.rm = TRUE)
  sd <- sd(data[[mod_var]],   na.rm = TRUE)
  
  # 3 Level: -1 SD, Mittelwert, +1 SD
  mod_levels <- c(m - sd, m, m + sd)
  mod_names  <- c("-1 SD", "Mittelwert", "+1 SD")
  
  # Vorhersage-Gitter
  grid <- expand.grid(
    x = seq(min(data[[x_var]], na.rm = TRUE),
            max(data[[x_var]], na.rm = TRUE),
            length.out = 80),
    mod = mod_levels
  )
  names(grid) <- c(x_var, mod_var)
  
  grid$pred <- predict(model, newdata = grid)
  
  # Legendengruppen
  grid$mod_level <- factor(
    mod_names[match(grid[[mod_var]], mod_levels)],
    levels = mod_names
  )
  
  ggplot(grid, aes(x = .data[[x_var]], y = pred, color = mod_level, group = mod_level)) +
    geom_line(linewidth = 1.2) +
    labs(
      title = title,
      x = x_lab,
      y = y_lab,
      color = legend_title
    ) +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(face = "bold", size = 20, hjust = 0.5),
      axis.title = element_text(size = 14),
      axis.text  = element_text(size = 12),
      legend.title = element_text(size = 14),
      legend.text  = element_text(size = 12),
      legend.position = "right",
      panel.grid.minor = element_blank()
    )
}


# 4 Plots Ballwurf X
p_x_imm <- make_interaction_plot_3lines(
  data = daten_x, model = model_h3_x_imm,
  x_var = "ssq_x_c", mod_var = "imm_c",
  title = "A)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Immersion"
)

p_x_pos <- make_interaction_plot_3lines(
  data = daten_x, model = model_h3_x_pos,
  x_var = "ssq_x_c", mod_var = "pos_c",
  title = "B)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Pos. Affekt"
)

p_x_neg <- make_interaction_plot_3lines(
  data = daten_x, model = model_h3_x_neg,
  x_var = "ssq_x_c", mod_var = "neg_c",
  title = "C)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Neg. Affekt"
)

p_x_ch <- make_interaction_plot_3lines(
  data = daten_x, model = model_h3_x_chall,
  x_var = "ssq_x_c", mod_var = "chall_c",
  title = "D)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Challenge"
)

fig_x <- (p_x_imm | p_x_pos) / (p_x_neg | p_x_ch) +
  plot_annotation(
    title = "Interaktionsplots H3 – Ballwurfspiel X: SSQ × GEQ-Subskalen",
    theme = theme(plot.title = element_text(face = "bold", size = 16, hjust = 0.5))
  )

fig_x

# 4 Plots Hexenspiel V2
p_v2_imm <- make_interaction_plot_3lines(
  data = daten_v2, model = model_h3_v2_imm,
  x_var = "ssq_v2_c", mod_var = "imm_c",
  title = "A)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Immersion"
)

p_v2_pos <- make_interaction_plot_3lines(
  data = daten_v2, model = model_h3_v2_pos,
  x_var = "ssq_v2_c", mod_var = "pos_c",
  title = "B)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Pos. Affekt"
)

p_v2_neg <- make_interaction_plot_3lines(
  data = daten_v2, model = model_h3_v2_neg,
  x_var = "ssq_v2_c", mod_var = "neg_c",
  title = "C)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Neg. Affekt"
)

p_v2_ch <- make_interaction_plot_3lines(
  data = daten_v2, model = model_h3_v2_chall,
  x_var = "ssq_v2_c", mod_var = "chall_c",
  title = "D)",
  x_lab = "CS (SSQ post, zentriert)",
  y_lab = "Performanzveränderung (Slope)",
  legend_title = "Challenge"
)

fig_v2 <- (p_v2_imm | p_v2_pos) / (p_v2_neg | p_v2_ch) +
  plot_annotation(
    title = "Interaktionsplots H3 – Hexenspiel V2: SSQ × GEQ-Subskalen",
    theme = theme(plot.title = element_text(face = "bold", size = 16, hjust = 0.5))
  )

fig_v2

# Export PNG
ggsave("H3_Interaktion_Ball_X_.png", fig_x,  width = 12, height = 8, dpi = 300)
ggsave("H3_Interaktion_Hexe_V2_.png", fig_v2, width = 12, height = 8, dpi = 300)


# Export: Deskriptiv + Modelle + Koeffizienten + Annahmen
models <- list(
  "X_Immersion"  = model_h3_x_imm,
  "X_PosAffekt"  = model_h3_x_pos,
  "X_NegAffekt"  = model_h3_x_neg,
  "X_Challenge"  = model_h3_x_chall,
  "V2_Immersion" = model_h3_v2_imm,
  "V2_PosAffekt" = model_h3_v2_pos,
  "V2_NegAffekt" = model_h3_v2_neg,
  "V2_Challenge" = model_h3_v2_chall
)

coef_table <- purrr::map_dfr(models, broom::tidy,  .id = "Modell")
glance_tbl <- purrr::map_dfr(models, broom::glance, .id = "Modell")

assumption_table <- purrr::map_dfr(
  names(models),
  function(mname) {
    mod <- models[[mname]]
    sh <- shapiro.test(residuals(mod))
    bp <- bptest(mod)
    cooks <- cooks.distance(mod)
    
    tibble(
      Modell       = mname,
      Shapiro_W    = unname(sh$statistic),
      Shapiro_p    = sh$p.value,
      BP_statistic = unname(bp$statistic),
      BP_p         = bp$p.value,
      Cooks_max    = max(cooks, na.rm = TRUE)
    )
  }
) %>%
  mutate(across(where(is.numeric), ~ round(.x, 3)))

write_xlsx(
  list(
    "Deskriptiv"    = des_h3,
    "Koeffizienten" = coef_table,
    "Modelle"       = glance_tbl,
    "Annahmen"      = assumption_table
  ),
  path = "H3_Ergebnisse_gesammelt_Neu.xlsx"
)