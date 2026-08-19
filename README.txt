# Conector API-Football → Excel V2.2

## Qué hace
1. Consulta API-Football.
2. Descarga los partidos de una fecha.
3. Puede consultar estadísticas y cuotas por fixture.
4. Escribe los partidos en `BASE_DATOS` y `EVALUADOR` del Excel V2.2.
5. NO inventa probabilidades: la probabilidad V2.2 se mantiene como dato del modelo.

## Seguridad
La API key NO está dentro del código. Debe guardarse como variable de entorno.

### Windows PowerShell
```powershell
$env:API_FOOTBALL_KEY="TU_CLAVE"
python api_v22.py --date 2026-08-20 --xlsx Excel_V2_2_Bot_Apuestas.xlsx --details
```

### Windows CMD
```cmd
set API_FOOTBALL_KEY=TU_CLAVE
python api_v22.py --date 2026-08-20 --xlsx Excel_V2_2_Bot_Apuestas.xlsx --details
```

### macOS/Linux
```bash
export API_FOOTBALL_KEY="TU_CLAVE"
python3 api_v22.py --date 2026-08-20 --xlsx Excel_V2_2_Bot_Apuestas.xlsx --details
```

## Instalación
```bash
pip install -r requirements.txt
```

## Primera prueba
Sin detalles, para gastar menos cuota:
```bash
python api_v22.py --date 2026-08-20 --xlsx Excel_V2_2_Bot_Apuestas.xlsx
```

Con estadísticas/cuotas:
```bash
python api_v22.py --date 2026-08-20 --xlsx Excel_V2_2_Bot_Apuestas.xlsx --details
```

## Próxima versión
La siguiente etapa debe:
- seleccionar automáticamente mercados disponibles;
- calcular medias de últimos 5/10 partidos;
- alimentar los criterios 0–1 de V2.2;
- calcular la probabilidad propia del modelo;
- comparar contra la cuota;
- generar ALERTA FUERTE/MEDIA;
- guardar histórico;
- y finalmente enviar la alerta a Telegram.

No se recomienda apostar automáticamente. Primero se debe validar el modelo con un histórico suficiente.
