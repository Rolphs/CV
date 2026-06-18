# Cómo operar este repo desde una sesión de Claude / Cowork

Esta nota documenta el método recomendado para que cualquier sesión de Claude
gestione este repositorio sin fricción.

## Método recomendado: shell del Mac + `gh` (NO GitHub Desktop)

`gh` ya está autenticado en el llavero (keychain) de macOS a nivel de máquina,
así que `git push`, `gh repo ...` y `gh pr ...` funcionan sin volver a hacer login.

**Instrucción para la sesión:**

1. Pide acceso a la carpeta del proyecto (`~/Documents/GitHub/CV`).
2. Opera por línea de comandos con `git` y la CLI `gh`.
3. Verifica antes de empezar:

   ```bash
   gh auth status
   ```

### Detalle importante sobre el shell

Hay dos tipos de shell según la herramienta:

- **Shell del Mac** (ej. la herramienta *Control your Mac* / `osascript` con
  `do shell script`): corre en macOS real, ve el llavero y `gh`. **Usar esta.**
  `gh` vive en `/opt/homebrew/bin`, así que exporta el PATH:

  ```bash
  do shell script "export PATH=/opt/homebrew/bin:/usr/local/bin:$PATH; cd ~/Documents/GitHub/CV && git push"
  ```

- **Shell sandbox (Linux)**: algunas sesiones solo tienen un shell aislado que
  NO ve el llavero ni tiene `gh`. Desde ahí `git push` no autentica. Si una
  sesión solo tiene este shell (y no `osascript`/Control your Mac), entonces hay
  dos opciones: (a) usar un Personal Access Token, o (b) recurrir a GitHub
  Desktop con control de pantalla. Ambas son el plan B.

## Notas de seguridad

- Nunca se hace `gh auth login` automáticamente: ese paso, que implica
  credenciales, lo corre el usuario una sola vez (ya hecho).
- Este repo es público. No subir secretos, tokens ni datos sensibles que no se
  quieran exponer.

## Estructura

Ver `README.md` para la estructura del proyecto y el flujo editorial.
La base de datos maestra para consumo del LLM está en
`data/master/cv_master.json`.
