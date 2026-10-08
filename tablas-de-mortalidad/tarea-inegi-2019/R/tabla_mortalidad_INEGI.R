# =====================================================================
#  TABLA DE MORTALIDAD DE MÉXICO, 2019 (ambos sexos) CON DATOS DE INEGI
#  Formato de la plantilla de la profesora: x, qx, lx, dx, Lx, Tx, ex
#  (+ px y la tabla de conmutación al 5 % como en tu código TMM_64_67.R)
# =====================================================================
#  Datos (carpeta "datos", ver LEEME):
#    inegi_defunciones_2019_por_edad_sexo.csv   <- EDR (defunciones ocurridas en 2019)
#    inegi_nacimientos_2019_por_sexo.csv        <- ENR (nacimientos ocurridos en 2019)
#    inegi_censo2020_poblacion_por_edad_sexo.csv<- Censo 2020 (población por edad)
#
#  Cómo usarlo en RStudio:
#    1) Pon este archivo y la carpeta "datos" juntas en una carpeta.
#    2) Session > Set Working Directory > To Source File Location
#    3) Source (o Ctrl+Shift+S).  Los resultados salen en la carpeta "salida".
# =====================================================================
options(scipen = 999)

# ---- 0. Parámetros -----------------------------------------------------
carpeta_datos  <- "datos"
carpeta_salida <- file.path(getwd(), "salida")   # cámbiala si quieres, p. ej. "C:/Users/spart/Desktop/tabla_mortalidad"
radix          <- 100000
suavizar       <- TRUE      # media móvil de 5 edades (10 a 97 años) para quitar la atracción de edades redondas del Censo
tasa_conm      <- 0.05      # tasa de interés de la tabla de conmutación (tu código usa 5 %)
token          <- Sys.getenv("INEGI_TOKEN", unset = "PEGA_AQUI_TU_TOKEN")   # solo se usa en la parte del API (sección 7)

redondea <- function(x) floor(x + 0.5)    # ROUND de Excel (round() de R redondea .5 al par)

# ---- 1. Leer los datos de INEGI -----------------------------------------
def <- read.csv(file.path(carpeta_datos, "inegi_defunciones_2019_por_edad_sexo.csv"), stringsAsFactors = FALSE)
nac <- read.csv(file.path(carpeta_datos, "inegi_nacimientos_2019_por_sexo.csv"))
pob <- read.csv(file.path(carpeta_datos, "inegi_censo2020_poblacion_por_edad_sexo.csv"), stringsAsFactors = FALSE)

EDAD <- 0:100
i_def <- match(as.character(EDAD), def$edad)
i_pob <- match(as.character(EDAD), pob$edad)

# ---- 2. Defunciones, población y nacimientos por edad -----------------------
defun   <- as.numeric(def$hombres[i_def] + def$mujeres[i_def] + def$sexo_no_especificado[i_def])  # defunciones ocurridas en 2019 (as.numeric evita el desbordamiento de enteros)
ne      <- as.numeric(sum(def[def$edad == "NE", c("hombres", "mujeres", "sexo_no_especificado")]))  # edad no especificada
defun_aj <- defun * (sum(defun) + ne) / sum(defun)          # reparte la edad no especificada en proporción
poblacion <- as.numeric(pob$hombres[i_pob] + pob$mujeres[i_pob])        # Censo 2020
nacimientos <- as.numeric(sum(nac$hombres, nac$mujeres, nac$sexo_no_especificado))

# media móvil de 5 edades (x-2 ... x+2): el elemento x está en la posición x + 1
media_movil <- function(v, ini = 10, fin = 97) {
  o <- v
  for (x in ini:fin) o[x + 1] <- mean(v[(x - 1):(x + 3)])
  o
}
D_us <- if (suavizar) media_movil(defun_aj) else defun_aj
P_us <- if (suavizar) media_movil(poblacion) else poblacion

# ---- 3. qx -------------------------------------------------------------------
#   q(0)  = defunciones de menores de 1 año / nacimientos            (nota técnica de la plantilla)
#   q(x)  = defunciones / (población + defunciones / 2),  x = 1..99
#   q(100)= 1   (100 y más: grupo abierto)
qx <- D_us / (P_us + D_us / 2)
qx[1]   <- D_us[1] / nacimientos
qx[101] <- 1

# ---- 4. Tabla de mortalidad (misma lógica que tu código: lx -> dx -> qx -> px) --
lx <- numeric(101); dx <- numeric(101)
lx[1] <- radix
for (i in 1:101) {
  dx[i] <- redondea(lx[i] * qx[i])
  if (i < 101) lx[i + 1] <- lx[i] - dx[i]
}
Lx <- redondea(c(lx[-1], 0) + dx / 2)      # Lx = l(x+1) + dx/2   (a la edad 100, l101 = 0)
Tx <- rev(cumsum(rev(Lx)))                 # Tx = Lx + L(x+1) + ... + L100
ex <- Tx / lx
px <- 1 - qx

# ---- 5. Tabla de conmutación al 5 % (lo que hacías con sum(Dx[1:85]), sum(Dx[2:85]), ...) ----
Dx_5 <- lx / ((1 + tasa_conm)^EDAD)
Nx_5 <- rev(cumsum(rev(Dx_5)))             # Nx = Dx + D(x+1) + ... + D100 (todas las sumas de golpe)

TMM_INEGI_2019 <- data.frame(x = EDAD, qx = round(qx, 6), lx = lx, dx = dx, Lx = Lx, Tx = Tx,
                             ex = round(ex, 4), px = round(px, 6), Dx_5pct = round(Dx_5, 4), Nx_5pct = round(Nx_5, 2))
datos_usados <- data.frame(x = EDAD, defunciones_2019 = defun, defunciones_edad_NE_repartida = round(defun_aj, 1),
                           defunciones_usadas = round(D_us, 1), poblacion_censo_2020 = poblacion, poblacion_usada = round(P_us, 1))

cat("\n===== TABLA DE MORTALIDAD, MÉXICO 2019 (INEGI) =====\n")
print(head(TMM_INEGI_2019, 10), row.names = FALSE)
cat("   ...\n")
print(tail(TMM_INEGI_2019, 5), row.names = FALSE)
cat(sprintf("\nControl: suma de dx = %s (debe ser %s)\n", sum(dx), radix))
cat(sprintf("Nacimientos 2019: %s | Defunciones de menores de 1 año: %s | q(0) = %.6f\n", format(nacimientos, big.mark = ","),
            format(round(D_us[1]), big.mark = ","), qx[1]))
cat(sprintf("e0 = %.2f años | e65 = %.2f años | l65/l0 = %.1f %% | l80/l0 = %.1f %%\n", ex[1], ex[66], 100 * lx[66] / radix, 100 * lx[81] / radix))

# ---- 6. Exportar a tu computadora -------------------------------------------------
dir.create(carpeta_salida, showWarnings = FALSE, recursive = TRUE)
write.csv (TMM_INEGI_2019, file.path(carpeta_salida, "tabla_mortalidad_INEGI_2019.csv"),      row.names = FALSE, fileEncoding = "UTF-8")
write.csv2(TMM_INEGI_2019, file.path(carpeta_salida, "tabla_mortalidad_INEGI_2019_excel_es.csv"), row.names = FALSE, fileEncoding = "UTF-8")  # coma decimal: se abre directo en Excel en español
if (requireNamespace("openxlsx", quietly = TRUE)) {
  openxlsx::write.xlsx(list("Tabla de mortalidad" = TMM_INEGI_2019, "Datos usados" = datos_usados),
                       file.path(carpeta_salida, "tabla_mortalidad_INEGI_2019.xlsx"), overwrite = TRUE)
} else {
  message("Para exportar a .xlsx instala el paquete: install.packages('openxlsx')")
}

# gráficos de la tabla (complemento del análisis)
png(file.path(carpeta_salida, "graficos_tabla_mortalidad.png"), width = 1400, height = 1000, res = 150)
par(mfrow = c(2, 2), mar = c(4.2, 4.5, 3, 1))
plot(EDAD, lx, type = "l", lwd = 2, col = "#2a78d6", main = "Sobrevivientes lx", xlab = "Edad", ylab = "lx"); abline(h = radix / 2, lty = 3)
plot(EDAD[1:100], qx[1:100], type = "l", lwd = 2, col = "#2a78d6", log = "y", main = "Probabilidad de morir qx (escala log)", xlab = "Edad", ylab = "qx")
plot(EDAD[1:100], dx[1:100], type = "l", lwd = 2, col = "#2a78d6", main = "Defunciones dx", xlab = "Edad", ylab = "dx")
plot(EDAD, ex, type = "l", lwd = 2, col = "#2a78d6", main = "Esperanza de vida ex", xlab = "Edad", ylab = "ex (anios)")
dev.off()
cat("\nArchivos guardados en:", normalizePath(carpeta_salida), "\n")

# ---- 7. (Opcional) API de INEGI: comprobación de la población del Censo 2020 ------------------
#  El API del Banco de Indicadores trae series AGREGADAS (p. ej. población total). No trae defunciones por
#  edad simple, por eso la tabla se calcula con los CSV de arriba. Aquí solo se verifica que la población
#  total del CSV coincida con la que publica INEGI (indicador 1002000001, Censo 2020).
consulta_api <- function(indicador, area = "00") {
  url <- sprintf("https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/%s/es/%s/false/BISE/2.0/%s?type=json",
                 indicador, area, token)
  obs <- jsonlite::fromJSON(url)$Series$OBSERVATIONS[[1]]
  data.frame(anio = as.integer(obs$TIME_PERIOD), valor = as.numeric(obs$OBS_VALUE))
}
if (requireNamespace("jsonlite", quietly = TRUE)) {
  pob_api <- tryCatch(consulta_api("1002000001"), error = function(e) NULL)
  if (!is.null(pob_api)) {
    p2020 <- pob_api$valor[pob_api$anio == 2020]
    p_csv <- sum(pob$hombres, pob$mujeres)
    cat(sprintf("\nAPI INEGI, población total 2020: %s | suma del CSV del Censo: %s | %s\n",
                format(p2020, big.mark = ","), format(p_csv, big.mark = ","), ifelse(p2020 == p_csv, "COINCIDEN", "NO coinciden")))
  } else cat("\n(No se pudo consultar el API; revisa tu conexión a internet.)\n")
} else message("Para usar el API instala: install.packages('jsonlite')")
