name: Bot Apuestas V2.2

on:
  workflow_dispatch:

permissions:
  contents: write

jobs:
  ejecutar-bot:
    runs-on: ubuntu-latest
    environment: BOT

    steps:
      - name: Descargar repositorio
        uses: actions/checkout@v4

      - name: Instalar Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Instalar dependencias
        run: |
          python -m pip install --upgrade pip
          pip install -r requisitos.txt

      - name: Comprobar archivos
        run: |
          echo "Archivos del repositorio:"
          ls -la
          echo "Python:"
          python --version

      - name: Ejecutar API-Football
        env:
          API_FOOTBALL_KEY: ${{ secrets.API_FOOTBALL_KEY }}
        run: |
          python api_v22.py \
            --date "$(date -u +%Y-%m-%d)" \
            --xlsx "Excel_V2_2_Motor_Automatico.xlsx" \
            --details

      - name: Guardar cambios
        run: |
          git config user.name "Bot V2.2"
          git config user.email "bot-v22@users.noreply.github.com"
          git add Excel_V2_2_Motor_Automatico.xlsx
          git diff --cached --quiet || git commit -m "Actualizar datos V2.2"
          git push
