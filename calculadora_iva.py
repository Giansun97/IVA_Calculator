import pandas as pd
import os
import numpy as np
import openpyxl
#import xlrd


def procesar_ventas(path: str) -> pd.DataFrame:
    """
    Procesa archivos de ventas en formato Excel para extraer información de interés.

    :param path: str, ruta del directorio donde se encuentran los archivos de ventas.
    :return: ventas: pd.DataFrame, DataFrame con la información de interés de los archivos de ventas.

    """

    # Define una lista vacía donde almacenarás los archivos Excel que contengan 'MCE' en el nombre
    lista_ventas = []
    ventas_list = []

    # Itera sobre los archivos en el directorio y agrega los que cumplan con el criterio a la lista
    for venta in os.listdir(path):
        if venta.endswith('.xlsx') and 'MCE' in venta:
            lista_ventas.append(venta)

    # Iterar sobre cada archivo dentro de la lista de ventas
    for venta in lista_ventas:
        # Leer cada archivo dentro de la ruta de archivos de ventas
        df_ventas = pd.read_excel(os.path.join(path, venta),
                                  skiprows=1)

        contribuyente, cuit_ventas, iva_debito = limpiar_data(df_ventas, venta)

        # Crear DataFrame con la info
        ventas = pd.DataFrame({'CUIT': cuit_ventas,
                               'Contribuyente': contribuyente,
                               'IVA debito': iva_debito},
                              index=[0])

        # Concateno los resultados en una lista
        ventas_list.append(ventas)

    # Concateno los resultados al data frame de ventas definitivo
    ventas = pd.concat(ventas_list, ignore_index=True)

    return ventas


def limpiar_data(df, venta):
    # Extraer el cuit del nombre del archivo
    cuit_ventas = venta.split('-')[3].strip()
    # Extraer el nombre de contribuyente del nombre del archivo
    contribuyente = venta.split('-')[4].strip().replace('.xlsx', '')
    df['IVA'] = df['IVA'] * df['Tipo Cambio']
    df.loc[df['Tipo'].str.contains('Nota de Crédito', na=False), 'IVA'] *= -1
    iva_debito = df['IVA'].sum()
    return contribuyente, cuit_ventas, iva_debito


def procesar_compras(path: str) -> pd.DataFrame:
    # Define una lista vacía donde almacenarás los archivos Excel que contengan 'MCE' en el nombre
    archivos_compras = []
    compras_list = []

    # Itera sobre los archivos en el directorio y agrega los que cumplan con el criterio a la lista
    for compra in os.listdir(path):
        if compra.endswith('.xlsx') and 'MCR' in compra:
            archivos_compras.append(compra)

    # Iterar sobre cada archivo dentro de la lista de ventas
    for compra in archivos_compras:
        # Leer cada archivo dentro de la ruta de archivos de ventas
        df_compras = pd.read_excel(os.path.join(path, compra),
                                   skiprows=1)

        contribuyente, cuit_compra, iva_credito = limpiar_data(df_compras, compra)

        # Crear DataFrame con la info
        compras = pd.DataFrame({'CUIT': cuit_compra, 'Contribuyente': contribuyente, 'IVA credito': iva_credito},
                               index=[0])

        # Concateno los resultados en una lista
        compras_list.append(compras)

    # Concateno los resultados al data frame de ventas definitivo
    compras = pd.concat(compras_list, ignore_index=True)

    return compras


def procesar_saldos_anteriores(excel_saldos_anteriores: str):
    # Leemos el archivo excel
    saldos_anteriores = pd.read_excel(excel_saldos_anteriores)

    return saldos_anteriores


def procesar_retenciones(path: str):
    archivos_retenciones = []
    retenciones_list = []

    # Itera sobre los archivos en el directorio y agrega los que cumplan con el criterio a la lista
    for retencion in os.listdir(path):
        if retencion.endswith('.xls') and 'Mis Retenciones' in retencion:
            archivos_retenciones.append(retencion)

    # Iterar sobre cada archivo dentro de la lista de ventas
    for retencion in archivos_retenciones:
        # Leer cada archivo dentro de la ruta de archivos de ventas
        mis_retenciones = pd.read_excel(os.path.join(path, retencion))

        # Extraer el cuit del nombre del archivo
        cuit_ret = retencion.split("-")[3].strip()

        # Extraer el nombre de contribuyente del nombre del archivo
        contribuyente = retencion.split("-")[4].strip().replace('.xls', '')

        total_ret = mis_retenciones['Importe Ret./Perc.'].sum()

        # Crear DataFrame con la info
        retenciones = pd.DataFrame({'CUIT': cuit_ret, 'Contribuyente': contribuyente, 'Total Ret': total_ret},
                                   index=[0])

        # Concateno los resultados en una lista
        retenciones_list.append(retenciones)

    retenciones = pd.concat(retenciones_list, ignore_index=True)

    retenciones['CUIT'] = retenciones['CUIT'].astype(np.int64)

    return retenciones


def mostrar_resultados(ventas: pd.DataFrame, compras: pd.DataFrame, saldos_anteriores, retenciones) -> pd.DataFrame:
    # Creamos el dataframe de resultados
    saldo_del_periodo = pd.merge(ventas,
                                 compras[['CUIT', 'IVA credito']],
                                 on='CUIT',
                                 how='left')

    saldo_del_periodo['CUIT'] = saldo_del_periodo['CUIT'].astype(np.int64)

    # Si se selecciono el archivo de los saldos anteriores
    if saldos_anteriores is not None:
        resultados = pd.merge(saldos_anteriores,
                              saldo_del_periodo[['CUIT', 'IVA debito', 'IVA credito']], on='CUIT',
                              how='outer')

        resultados['Contribuyente'].fillna(resultados['CUIT'].map(saldo_del_periodo.set_index('CUIT')['Contribuyente']),
                                           inplace=True)

        # Llenamos los valores faltantes con ceros
        resultados.fillna(0, inplace=True)

        # Calculamos el saldo tecnico del periodo
        resultados['Saldo Tecnico Del Periodo'] = (
                resultados['IVA debito'] -
                resultados['IVA credito'] -
                resultados['Saldo Tecnico Periodo Anterior']
        )

        if retenciones is not None:
            resultados = pd.merge(resultados,
                                  retenciones[['CUIT', 'Total Ret']],
                                  on='CUIT',
                                  how='outer')

            # Llenamos los valores faltantes con ceros
            resultados.fillna(0, inplace=True)

            # Calculamos el saldo de libre del periodo
            resultados['Saldo Libre Disponibilidad del Periodo'] = (
                    resultados['Saldo de Libre Disponibilidad Periodo Anterior'] +
                    resultados['Total Ret']
            )

            print(resultados)

            resultados['Saldo Del Periodo'] = resultados['Saldo Tecnico Del Periodo'] - resultados[
                'Saldo Libre Disponibilidad del Periodo']

            resultados['Resultado'] = 0

            # Condiciones donde el resultado se interpreta como saldo a favor o como saldo a pagar
            resultados.loc[resultados['Saldo Del Periodo'] > 0, 'Resultado'] = 'IVA Saldo a pagar'
            resultados.loc[resultados['Saldo Del Periodo'] < 0, 'Resultado'] = 'IVA Saldo a favor'

        else:
            # Calculamos el saldo total del periodo
            resultados['Saldo Del Periodo'] = (
                    resultados['Saldo Tecnico Del Periodo'] -
                    resultados['Saldo de Libre Disponibilidad Periodo Anterior']
            )

            resultados['Resultado'] = 0

            # Condiciones donde el resultado se interpreta como saldo a favor o como saldo a pagar
            resultados.loc[resultados['Saldo Del Periodo'] > 0, 'Resultado'] = 'IVA Saldo a pagar'
            resultados.loc[resultados['Saldo Del Periodo'] < 0, 'Resultado'] = 'IVA Saldo a favor'

    # Si no se ingreso el archivo de los saldos anteriores
    else:
        resultados = saldo_del_periodo[['CUIT', 'Contribuyente', 'IVA debito', 'IVA credito', ]]

        # Llenamos los valores faltantes con ceros
        resultados.fillna(0, inplace=True)

        resultados['Saldo Tecnico Del Periodo'] = (
                resultados['IVA debito'] -
                resultados['IVA credito']
        )

        if retenciones is not None:
            resultados = pd.merge(resultados,
                                  retenciones[['CUIT', 'Total Ret']],
                                  on='CUIT',
                                  how='outer')

            # Llenamos los valores faltantes con ceros
            resultados['Total Ret'].fillna(0, inplace=True)

            # Calculamos el saldo del periodo
            resultados['Saldo Del Periodo'] = (
                    resultados['Saldo Tecnico Del Periodo'] -
                    resultados['Total Ret']
            )

            resultados['Resultado'] = 0

            # Condiciones donde el resultado se interpreta como saldo a favor o como saldo a pagar
            resultados.loc[resultados['Saldo Del Periodo'] > 0, 'Resultado'] = 'IVA Saldo a pagar'
            resultados.loc[resultados['Saldo Del Periodo'] < 0, 'Resultado'] = 'IVA Saldo a favor'

    return resultados
