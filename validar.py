import pandas as pd
from datetime import timedelta

def auditoria_logistica():
    file_sistema = 'Exporte_tmland.xlsx'
    file_nuevo = 'Plano.xlsx'
    try:
        df_sistema = pd.read_excel(file_sistema)
        df_nuevo = pd.read_excel(file_nuevo)

        # Limpieza de nombres de columnas
        for df in [df_sistema, df_nuevo]:
            df.columns = df.columns.str.strip()
        
        mapeo = {'Vehiculo': 'Placa'}
        df_sistema.rename(columns=mapeo, inplace=True)
        df_nuevo.rename(columns=mapeo, inplace=True)

        # Conversión de formatos de fecha
        cols_fecha = ['Fecha inicio Real', 'Fecha Fin Real']
        for df in [df_sistema, df_nuevo]:
            for col in cols_fecha:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            df.dropna(subset=cols_fecha, inplace=True)

        df_total = pd.concat([df_sistema, df_nuevo])
        
        # Esto extrae el viaje que terminó más tarde en ESE día específico
        max_por_dia = df_total.groupby(['Placa', df_total['Fecha Fin Real'].dt.date])['Fecha Fin Real'].max().to_dict()

        def obtener_sugerencia_segura(placa, fecha_conflicto):
            """Busca el último hueco libre del mismo día del conflicto."""
            dia_exacto = fecha_conflicto.date()
            
            # Busca el último servicio de ESE día.
            ultimo_fin_del_dia = max_por_dia.get((placa, dia_exacto), fecha_conflicto)
            
            inicio_seguro = ultimo_fin_del_dia + timedelta(minutes=1)
            fin_seguro = inicio_seguro + timedelta(minutes=5)
            
            # Actualizamos la memoria de ese día específico por si la misma placa tiene más choques hoy
            max_por_dia[(placa, dia_exacto)] = fin_seguro
            return inicio_seguro, fin_seguro

        # VALIDACIÓN INTERNA (BD2 contra sí misma)
        df_int = df_nuevo.sort_values(by=['Placa', 'Fecha inicio Real']).copy()
        df_int['Fin_Anterior'] = df_int.groupby('Placa')['Fecha Fin Real'].shift(1)
        
        errores_internos = df_int[df_int['Fecha inicio Real'] <= df_int['Fin_Anterior']].copy()

        if not errores_internos.empty:
            sug_ini, sug_fin = [], []
            for _, fila in errores_internos.iterrows():
                ini, fin = obtener_sugerencia_segura(fila['Placa'], fila['Fecha inicio Real'])
                sug_ini.append(ini)
                sug_fin.append(fin)
            
            errores_internos['Sugerencia_Inicio_Libre'] = sug_ini
            errores_internos['Sugerencia_Fin_Libre'] = sug_fin
            
            # Conversión a texto formato limpio
            columnas_a_texto = ['Fecha inicio Real', 'Fecha Fin Real', 'Fin_Anterior', 'Sugerencia_Inicio_Libre', 'Sugerencia_Fin_Libre']
            for col in columnas_a_texto:
                if col in errores_internos.columns:
                    errores_internos[col] = errores_internos[col].dt.strftime('%Y/%m/%d %H:%M')
            
            errores_internos.to_excel('ERRORES_INTERNOS_EN_EL_PLANO.xlsx', index=False)
            print(f" {len(errores_internos)} cruces dentro del plano")
        else:
            print("No tiene cruces dentro del plano.")

        #VALIDACIÓN EXTERNA (PLANO contra EXPORTE TMLAND)
        cruce = df_nuevo.merge(df_sistema, on='Placa', suffixes=('_NUEVO', '_SISTEMA'))
        
        filtro_externo = (
            (cruce['Fecha inicio Real_NUEVO'] <= cruce['Fecha Fin Real_SISTEMA']) & 
            (cruce['Fecha Fin Real_NUEVO'] >= cruce['Fecha inicio Real_SISTEMA'])
        )
        
        errores_externos = cruce[filtro_externo].copy()

        if not errores_externos.empty:
            sug_ini_ext, sug_fin_ext = [], []
            for _, fila in errores_externos.iterrows():
                # Evaluamos el conflicto usando el día del viaje NUEVO que falló
                ini, fin = obtener_sugerencia_segura(fila['Placa'], fila['Fecha inicio Real_NUEVO'])
                sug_ini_ext.append(ini)
                sug_fin_ext.append(fin)
                
            errores_externos['Sugerencia_Inicio_Libre'] = sug_ini_ext
            errores_externos['Sugerencia_Fin_Libre'] = sug_fin_ext

            # Conversión a texto formato limpio
            columnas_ext_texto = [
                'Fecha inicio Real_NUEVO', 'Fecha Fin Real_NUEVO', 
                'Fecha inicio Real_SISTEMA', 'Fecha Fin Real_SISTEMA', 
                'Sugerencia_Inicio_Libre', 'Sugerencia_Fin_Libre'
            ]
            for col in columnas_ext_texto:
                if col in errores_externos.columns:
                    errores_externos[col] = errores_externos[col].dt.strftime('%Y/%m/%d %H:%M')

            errores_externos.to_excel('ERRORES_CONTRA_SISTEMA.xlsx', index=False)
            print(f"{len(errores_externos)} cruces contra TMLAND.")
        else:
            print("No hay cruces contra el sistema.")

    except Exception as e:
        print(f"ERROR CRÍTICO: {e}")

if __name__ == "__main__":
    auditoria_logistica()
