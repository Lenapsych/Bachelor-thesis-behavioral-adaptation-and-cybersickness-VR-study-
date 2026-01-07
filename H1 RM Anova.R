library(tidyverse)  # Daten einlesen/manipulieren + ggplot2
library(readxl)     # Excel-Dateien einlesen 
library(janitor)    # clean_names() für einheitliche Spaltennamen
library(afex)       # RM-ANOVA + Sphärizität + GG-Korrektur 
library(writexl)    # Ergebnisse als Excel ausgeben


# Basisdaten laden + finales Sample definieren (N=59)

daten <- read_csv("Deskriptive_Statistik.csv", show_col_types = FALSE) %>%
  clean_names() %>%   # sicherheitshalber
  mutate(
    code            = as.character(code),
    hex_perf_total  = as.numeric(gsub(",", ".", as.character(hex_perf_total))),
    ball_perf_total = as.numeric(gsub(",", ".", as.character(ball_perf_total)))
  )

# Ausschlussliste laden (enthält Codes, die raus müssen)
excl <- read_excel("Ausschluss_gesamt.xlsx") %>%
  clean_names() %>%
  mutate(code = as.character(code))

excluded_codes <- excl %>%
  distinct(code) %>%
  pull(code)

# finales Sample: alle, die NICHT ausgeschlossen sind und gültige Performanzwerte haben
sample_h12 <- daten %>%
  filter(!code %in% excluded_codes) %>%
  filter(!is.na(hex_perf_total), !is.na(ball_perf_total))

# Kontrolle
sample_h12 %>%
  summarise(N = n(), n_codes = n_distinct(code)) %>%
  print()

valid_codes <- sample_h12$code  


# Intervall-Daten (Blockdaten) einlesen und Performanz berechnen
# V2 = Hexenspiel V2, X = Ballwurfspiel
v2 <- read_csv("V2_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

x <- read_csv("X_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

blocks_keep <- c("MISC 1", "MISC 2", "MISC 3", "MISC 4", "MISC 5")

# Hexenspiel V2: Erfolg = c_collected, Versuch = triggers 
v2_perf <- v2 %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    interval_num = as.integer(gsub("MISC ", "", interval)),      # Block 1..5
    treffer      = c_collected,
    versuche     = triggers,
    perf         = treffer / pmax(versuche, 1)                   # Treffer/Versuch
  ) %>%
  select(code, interval_num, perf) %>%
  filter(code %in% valid_codes)

# Ballwurf X: Erfolg = t_hit, Versuch = shoot
x_perf <- x %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    interval_num = as.integer(gsub("MISC ", "", interval)),      # Block 1..5
    treffer      = t_hit,
    versuche     = shoot,
    perf         = treffer / pmax(versuche, 1)
  ) %>%
  select(code, interval_num, perf) %>%
  filter(code %in% valid_codes)


# Datensatz für RM-ANOVA bauen 
perf_long <- bind_rows(
  v2_perf %>% mutate(game = "Hexenspiel V2"),
  x_perf  %>% mutate(game = "Ballwurfspiel X")
) %>%
  mutate(
    code = as.character(code),
    # Für RM-ANOVA muss der Zeitblock als Faktor behandelt werden:
    interval_num = factor(interval_num, levels = 1:5)
  )

# Datensätze pro Spiel
v2_data <- perf_long %>% filter(game == "Hexenspiel V2")
x_data  <- perf_long %>% filter(game == "Ballwurfspiel X")

# Kontrollchecks: balanciert? 
v2_data %>% count(interval_num) %>% print()
x_data  %>% count(interval_num) %>% print()


# RM-ANOVA (afex) pro Spiel
# Greenhouse–Geisser Korrektur, weil Sphärizität verletzt ist
aov_v2 <- afex::aov_ez(
  id = "code",
  dv = "perf",
  within = "interval_num",
  data = v2_data,
  anova_table = list(correction = "GG", es = "ges")
)

aov_x <- afex::aov_ez(
  id = "code",
  dv = "perf",
  within = "interval_num",
  data = x_data,
  anova_table = list(correction = "GG", es = "ges")
)

# Output in der Konsole
aov_v2
aov_x

# ANOVA Tabelle
library(stringr)
library(dplyr)

anova_v2_export <- afex::nice(aov_v2) %>%
  filter(Effect == "interval_num") %>%
  mutate(
    df1 = as.numeric(str_trim(str_split(df, ",", simplify = TRUE)[,1])),
    df2 = as.numeric(str_trim(str_split(df, ",", simplify = TRUE)[,2]))
  ) %>%
  transmute(
    Spiel = "Hexenspiel V2",
    df1, df2,
    MSE,
    F,
    p   = p.value,
    ges
  )

anova_x_export <- afex::nice(aov_x) %>%
  filter(Effect == "interval_num") %>%
  mutate(
    df1 = as.numeric(str_trim(str_split(df, ",", simplify = TRUE)[,1])),
    df2 = as.numeric(str_trim(str_split(df, ",", simplify = TRUE)[,2]))
  ) %>%
  transmute(
    Spiel = "Ballwurfspiel X",
    df1, df2,
    MSE,
    F,
    p   = p.value,
    ges
  )

anova_tbl <- bind_rows(anova_v2_export, anova_x_export)
anova_tbl


# Voraussetzungen/Diagnostik
# Normalverteilung der Residuen (Shapiro-Wilk auf Modellresiduen)
lm_v2 <- lm(perf ~ interval_num + code, data = v2_data)
lm_x  <- lm(perf ~ interval_num + code, data = x_data)

sh_v2 <- shapiro.test(residuals(lm_v2))
sh_x  <- shapiro.test(residuals(lm_x))

sh_v2
sh_x

# Ausreißer prüfen über Δ-Performanz (Block 5 - Block 1)
delta_v2 <- v2_data %>%
  filter(interval_num %in% c("1", "5")) %>%
  pivot_wider(names_from = interval_num, values_from = perf, names_prefix = "B") %>%
  mutate(delta = B5 - B1, z = as.numeric(scale(delta)))

delta_x <- x_data %>%
  filter(interval_num %in% c("1", "5")) %>%
  pivot_wider(names_from = interval_num, values_from = perf, names_prefix = "B") %>%
  mutate(delta = B5 - B1, z = as.numeric(scale(delta)))

summary(abs(delta_v2$z) > 3)
summary(abs(delta_x$z) > 3)


# Ergänzende Analyse: Delta-Tests gegen 0 
# t-Test (zweiseitig)
t_delta_v2_two <- t.test(delta_v2$delta, mu = 0)
t_delta_x_two  <- t.test(delta_x$delta,  mu = 0)

# gerichteter Test (alternative="greater" = mittlere Verbesserung > 0)
t_delta_v2_one <- t.test(delta_v2$delta, mu = 0, alternative = "greater")
t_delta_x_one  <- t.test(delta_x$delta,  mu = 0, alternative = "greater")

t_delta_v2_two
t_delta_x_two
t_delta_v2_one
t_delta_x_one


# Plot: Mittelwerte ± Standardfehler über Blöcke pro Spiel
plot_data <- perf_long %>%
  group_by(game, interval_num) %>%
  summarise(
    mean_perf = mean(perf, na.rm = TRUE),
    sd_perf   = sd(perf, na.rm = TRUE),
    n         = n(),
    se_perf   = sd_perf / sqrt(n),
    .groups   = "drop"
  )

ggplot(plot_data, aes(x = interval_num, y = mean_perf, group = 1)) +
  geom_line(linewidth = 1) +
  geom_point(size = 2) +
  geom_errorbar(aes(ymin = mean_perf - se_perf, ymax = mean_perf + se_perf),
                width = 0.15) +
  facet_wrap(~ game) +
  labs(
    title = "Performanzverlauf über die Zeitblöcke",
    x     = "Block (je 2 Minuten)",
    y     = "Performanz (Treffer / Versuche)"
  ) +
  theme_minimal(base_size = 12)


# Ergebnisse exportieren 
make_result_row <- function(delta_vec, spiel_label) {
  tt <- t.test(delta_vec, mu = 0)  # zweiseitig
  tibble(
    Spiel       = spiel_label,
    n           = sum(!is.na(delta_vec)),
    delta_mean  = mean(delta_vec, na.rm = TRUE),
    delta_sd    = sd(delta_vec, na.rm = TRUE),
    t_delta     = unname(tt$statistic),
    df_delta    = unname(tt$parameter),
    p_delta     = unname(tt$p.value)
  )
}

# Delta-Tabelle
tabelle_delta <- bind_rows(
  make_result_row(delta_v2$delta, "Hexenspiel V2"),
  make_result_row(delta_x$delta,  "Ballwurf-Spiel X")
)


# Voraussetzungen 
assumptions_tbl <- tibble(
  Spiel = c("Hexenspiel V2", "Ballwurfspiel X"),
  Shapiro_W = c(unname(sh_v2$statistic), unname(sh_x$statistic)),
  Shapiro_p = c(unname(sh_v2$p.value),   unname(sh_x$p.value)),
  Outlier_abs_z_gt_3 = c(
    sum(abs(delta_v2$z) > 3, na.rm = TRUE),
    sum(abs(delta_x$z)  > 3, na.rm = TRUE)
  )
)

# Gesamt-Export
write_xlsx(
  list(
    H1_ANOVA        = anova_tbl,
    H1_Delta        = tabelle_delta,
    H1_Assumptions  = assumptions_tbl,
    H1_Plot_Data    = plot_data
  ),
  path = "H1_Ergebnisse_01.01.xlsx"
)

# Sample für später speichern 
save(sample_h12, file = "sample_h12_export_Wiederholung.01.01.RData")
