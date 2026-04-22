import pandas as pd

def ejecutar_validacion():
    # 1. Nombres de tus archivos (Asegúrarse de que se llamen exactamente igual)
    file_sistema = 'BD1.xlsx'
    file_nuevo = 'BASE2.xlsx'

    try:
        # 2. Cargar los archivos
        df_sistema = pd.read_excel(file_sistema)
        df_nuevo = pd.read_excel(file_nuevo)

        # 3. Normalizar columnas, Si uno dice 'Vehiculo' y otro 'Placa', los unimos
        mapeo = {'Vehiculo': 'Placa', 'Placa ': 'Placa'} # Por si hay espacios locos
        df_sistema.rename(columns=mapeo, inplace=True)
        df_nuevo.rename(columns=mapeo, inplace=True)

        # 4. Limpieza de fechas
        for df in [df_sistema, df_nuevo]:
            df['Fecha inicio Real'] = pd.to_datetime(df['Fecha inicio Real'], errors='coerce')
            df['Fecha Fin Real'] = pd.to_datetime(df['Fecha Fin Real'], errors='coerce')
            # Eliminamos filas que no tengan fechas válidas
            df.dropna(subset=['Fecha inicio Real', 'Fecha Fin Real'], inplace=True)

        # 5. Cruzar por Placa y detectar choques
        # Esto crea una tabla donde comparamos cada viaje nuevo con los del sistema
        cruce = df_nuevo.merge(df_sistema, on='Placa', suffixes=('_NUEVO', '_SISTEMA'))

        # Lógica (InicioA < FinB) AND (FinA > InicioB)
        filtro_choque = (
            (cruce['Fecha inicio Real_NUEVO'] < cruce['Fecha Fin Real_SISTEMA']) & 
            (cruce['Fecha Fin Real_NUEVO'] > cruce['Fecha inicio Real_SISTEMA'])
        )

        reporte_errores = cruce[filtro_choque].copy()

        # 6. Resultados
        if not reporte_errores.empty:
            # Limpiamos el reporte para que sea legible
            columnas_finales = [
                'Placa', 'Fecha inicio Real_NUEVO', 'Fecha Fin Real_NUEVO', 
                'Fecha inicio Real_SISTEMA', 'Fecha Fin Real_SISTEMA'
            ]
            reporte_errores[columnas_finales].to_excel('CRUCES_DETECTADOS.xlsx', index=False)
            print(f"Se detectaron {len(reporte_errores)} cruces de horario, revisar archivo exportado en la carpeta.")
        else:
            print("No hay cruces.")

    except Exception as e:
        print(f"ERROR TÉCNICO: {e}")
        print("Verifica que los nombres de las columnas en tu Excel sean 'Placa', 'Fecha inicio Real' y 'Fecha Fin Real'.")

if __name__ == "__main__":
    ejecutar_validacion()
