library(tidyverse)  # dplyr/tidyr/ggplot2
library(readxl)     
library(janitor)    # clean_names()
library(plm)        # Panelmodelle (Fixed Effects / within)
library(lmtest)     # Tests 
library(sandwich)   # robuste Varianz-Kovarianz-Matrizen
library(broom)      
library(writexl)    # Excel-Export


# Sample laden (n = 59)
sample_h12 <- read_csv("sample_h12_export.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code)) %>%
  filter(!is.na(code))

sample_h12 %>%
  summarise(n = n(), n_codes = n_distinct(code)) %>%
  print()

# MW numerisch machen
sample_h12 <- sample_h12 %>%
  mutate(
    misc_mean_v2 = as.numeric(gsub(",", ".", as.character(misc_mean_v2))),
    misc_mean_x  = as.numeric(gsub(",", ".", as.character(misc_mean_x)))
  )


# Deskriptive Statistik 
des_hex <- sample_h12 %>%
  summarise(
    n         = n(),
    perf_mean = mean(hex_perf_total, na.rm = TRUE),
    perf_sd   = sd(hex_perf_total,   na.rm = TRUE),
    misc_mean = mean(misc_mean_v2,   na.rm = TRUE),
    misc_sd   = sd(misc_mean_v2,     na.rm = TRUE)
  )

des_ball <- sample_h12 %>%
  summarise(
    n         = n(),
    perf_mean = mean(ball_perf_total, na.rm = TRUE),
    perf_sd   = sd(ball_perf_total,   na.rm = TRUE),
    misc_mean = mean(misc_mean_x,     na.rm = TRUE),
    misc_sd   = sd(misc_mean_x,       na.rm = TRUE)
  )

des_hex; des_ball


# Performanz blockweise einlesen + berechnen
v2 <- read_csv("V2_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

x  <- read_csv("X_Analysis.csv", show_col_types = FALSE) %>%
  clean_names() %>%
  mutate(code = as.character(code))

blocks_keep <- c("MISC 1","MISC 2","MISC 3","MISC 4","MISC 5")

# Hexe V2: Erfolg = c_collected, Versuch = triggers
v2_perf <- v2 %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    block       = as.integer(gsub("MISC ", "", interval)), # 1..5
    treffer_v2  = c_collected,
    versuche_v2 = triggers,
    perf_v2     = treffer_v2 / pmax(versuche_v2, 1)
  ) %>%
  select(code, block, perf_v2)

# Ball X: Erfolg = t_hit, Versuch = shoot
x_perf <- x %>%
  filter(interval %in% blocks_keep) %>%
  mutate(
    block       = as.integer(gsub("MISC ", "", interval)),
    treffer_x   = t_hit,
    versuche_x  = shoot,
    perf_x      = treffer_x / pmax(versuche_x, 1)
  ) %>%
  select(code, block, perf_x)

# Nur Codes aus finalem Sample
valid_codes <- sample_h12$code
v2_perf <- v2_perf %>% filter(code %in% valid_codes)
x_perf  <- x_perf  %>% filter(code %in% valid_codes)


# MISC blockweise (Zuordnung: 2/4/6/8min/post -> Block 1..5)
misc_v2_long <- sample_h12 %>%
  select(code, misc2min_v2, misc4min_v2, misc6min_v2, misc8min_v2, mis_cpost_v2) %>%
  mutate(across(-code, ~ as.numeric(gsub(",", ".", as.character(.))))) %>%
  pivot_longer(cols = -code, names_to = "misc_time", values_to = "misc_score") %>%
  mutate(
    block = case_when(
      misc_time == "misc2min_v2"  ~ 1L,
      misc_time == "misc4min_v2"  ~ 2L,
      misc_time == "misc6min_v2"  ~ 3L,
      misc_time == "misc8min_v2"  ~ 4L,
      misc_time == "mis_cpost_v2" ~ 5L,
      TRUE ~ NA_integer_
    )
  ) %>%
  filter(!is.na(block))

misc_x_long <- sample_h12 %>%
  select(code, misc2min_x, misc4min_x, misc6min_x, misc8min_x, mis_cpost_x) %>%
  mutate(across(-code, ~ as.numeric(gsub(",", ".", as.character(.))))) %>%
  pivot_longer(cols = -code, names_to = "misc_time", values_to = "misc_score") %>%
  mutate(
    block = case_when(
      misc_time == "misc2min_x"  ~ 1L,
      misc_time == "misc4min_x"  ~ 2L,
      misc_time == "misc6min_x"  ~ 3L,
      misc_time == "misc8min_x"  ~ 4L,
      misc_time == "mis_cpost_x" ~ 5L,
      TRUE ~ NA_integer_
    )
  ) %>%
  filter(!is.na(block))


# Paneldaten bauen (MISC + Performanz pro Block)
panel_v2 <- misc_v2_long %>%
  left_join(v2_perf, by = c("code", "block")) %>%
  filter(!is.na(misc_score), !is.na(perf_v2))

panel_x <- misc_x_long %>%
  left_join(x_perf, by = c("code", "block")) %>%
  filter(!is.na(misc_score), !is.na(perf_x))

# Kontrolle: sollte balanciert sein (59 Personen x 5 Blöcke = 295 Zeilen)
panel_v2 %>% summarise(n = n(), codes = n_distinct(code), blocks = n_distinct(block)) %>% print()
panel_x  %>% summarise(n = n(), codes = n_distinct(code), blocks = n_distinct(block)) %>% print()


# Voraussetzungen
# Within-Variation prüfen 
within_var_v2 <- panel_v2 %>%
  group_by(code) %>%
  summarise(sd_misc = sd(misc_score, na.rm = TRUE),
            sd_perf = sd(perf_v2,    na.rm = TRUE),
            .groups = "drop")

within_var_x <- panel_x %>%
  group_by(code) %>%
  summarise(sd_misc = sd(misc_score, na.rm = TRUE),
            sd_perf = sd(perf_x,     na.rm = TRUE),
            .groups = "drop")

# Wie viele Personen haben Variation?
within_var_v2 %>% summarise(
  n_codes = n(),
  n_sd_misc_gt0 = sum(sd_misc > 0, na.rm = TRUE),
  n_sd_perf_gt0 = sum(sd_perf > 0, na.rm = TRUE)
) %>% print()

within_var_x %>% summarise(
  n_codes = n(),
  n_sd_misc_gt0 = sum(sd_misc > 0, na.rm = TRUE),
  n_sd_perf_gt0 = sum(sd_perf > 0, na.rm = TRUE)
) %>% print()

# Lineare Beziehung prüfen (Scatterplots)
ggplot(panel_v2, aes(x = misc_score, y = perf_v2)) +
  geom_point(alpha = .5) +
  geom_smooth(method = "lm", se = FALSE) +
  facet_wrap(~ block) +
  labs(title = "Hexenspiel V2: Zusammenhang MISC und Performanz (pro Block)",
       x = "MISC (CS)",
       y = "Performanz (Treffer/Versuch)") +
  theme_minimal()

ggplot(panel_x, aes(x = misc_score, y = perf_x)) +
  geom_point(alpha = .5) +
  geom_smooth(method = "lm", se = FALSE) +
  facet_wrap(~ block) +
  labs(title = "Ballwurfspiel X: Zusammenhang MISC und Performanz (pro Block)",
       x = "MISC (CS)",
       y = "Performanz (Treffer/Versuch)") +
  theme_minimal()


# Panelmodell
pdata_v2 <- pdata.frame(panel_v2, index = c("code", "block"))
pdata_x  <- pdata.frame(panel_x,  index = c("code", "block"))

model_h2_v2_panel <- plm(perf_v2 ~ misc_score, data = pdata_v2, model = "within")
model_h2_x_panel  <- plm(perf_x  ~ misc_score, data = pdata_x,  model = "within")

summary(model_h2_v2_panel)
summary(model_h2_x_panel)

# Robuste (cluster) Standardfehler auf Personenebene 
robust_v2 <- coeftest(model_h2_v2_panel, vcov = vcovHC(model_h2_v2_panel, type = "HC1", cluster = "group"))
robust_x  <- coeftest(model_h2_x_panel,  vcov = vcovHC(model_h2_x_panel,  type = "HC1", cluster = "group"))

robust_v2
robust_x

# Tests auf Heteroskedastizität / serielle Korrelation 
bp_v2 <- bptest(model_h2_v2_panel)
bp_x  <- bptest(model_h2_x_panel)
bp_v2; bp_x


# Ergebnisse in Excel schreiben 
extract_robust <- function(ct, spiel_label) {
  tibble(
    Spiel = spiel_label,
    b     = ct["misc_score", 1],
    SE    = ct["misc_score", 2],
    t     = ct["misc_score", 3],
    p     = ct["misc_score", 4]
  )
}

regress_h2_panel <- bind_rows(
  extract_robust(robust_v2, "Hexenspiel V2"),
  extract_robust(robust_x,  "Ballwurfspiel X")
) %>%
  mutate(across(where(is.numeric), as.numeric))

# R² (within) aus glance- Funktion (aus paket broom)
glance_v2 <- broom::glance(model_h2_v2_panel)
glance_x  <- broom::glance(model_h2_x_panel)

regress_h2_panel <- regress_h2_panel %>%
  mutate(
    R2 = c(glance_v2$r.squared, glance_x$r.squared)
  )

# Voraussetzungen-Zusammenfassung als Tabelle 
assumptions_h2 <- tibble(
  Spiel = c("Hexenspiel V2", "Ballwurfspiel X"),
  n_panel = c(nrow(panel_v2), nrow(panel_x)),
  n_codes = c(n_distinct(panel_v2$code), n_distinct(panel_x$code)),
  n_blocks = c(n_distinct(panel_v2$block), n_distinct(panel_x$block)),
  bp_p = c(bp_v2$p.value, bp_x$p.value)
)

# Excel Datei auswerfen
write_xlsx(
  list(
    H2_deskriptiv   = bind_rows(des_hex %>% mutate(Spiel="Hexenspiel V2"),
                                des_ball %>% mutate(Spiel="Ballwurfspiel X")),
    H2_panel_robust = regress_h2_panel,
    H2_assumptions  = assumptions_h2
  ),
  "H2_Panel_Ergebnisse.01.01.xlsx"
)


# Tabelle als PNG auswerfen
library(gridExtra)
library(grid)

tabelle_h2_panel <- regress_h2_panel %>%
  mutate(across(where(is.numeric), ~ round(., 2))) %>%
  rename(B = b)

tt <- ttheme_minimal(
  core = list(
    fg_params = list(cex = 0.9, fontface = "plain"),
    bg_params = list(fill = "white", col = NA)
  ),
  colhead = list(
    fg_params = list(cex = 0.9, fontface = "italic"),
    bg_params = list(fill = "white", col = NA)
  )
)

tbl <- tableGrob(tabelle_h2_panel, rows = NULL, theme = tt)

title <- textGrob("Ergebnisse Panel-Analyse H2", gp = gpar(fontsize = 14, fontface = "bold"))

png("H2_Panel_Tabelle.01.01.png", width = 2000, height = 800, res = 300)
grid.newpage()
pushViewport(viewport(layout = grid.layout(2, 1, heights = unit(c(1, 5), "null"))))

pushViewport(viewport(layout.pos.row = 1, layout.pos.col = 1))
grid.draw(title)
upViewport()

pushViewport(viewport(layout.pos.row = 2, layout.pos.col = 1))
grid.draw(tbl)

grid.lines(x = unit(c(0, 1), "npc"), y = unit(1, "npc"), gp = gpar(lwd = 3))
grid.lines(x = unit(c(0, 1), "npc"), y = unit(0.87, "npc"), gp = gpar(lwd = 2))
grid.lines(x = unit(c(0, 1), "npc"), y = unit(0, "npc"), gp = gpar(lwd = 3))

upViewport(2)
dev.off()
