import os
import logging
import pandas as pd
from flask import Flask, render_template, request, redirect, send_file, url_for
from werkzeug.utils import secure_filename
import uuid
from pathlib import Path

app = Flask(__name__,
            template_folder=os.path.join(os.path.dirname(__file__),
                                         'templates'))

# Configuração de Logging
logging.basicConfig(filename='error.log', level=logging.ERROR)
DOWNLOADS_FOLDER = str(Path.home() / "Downloads")
UPLOAD_FOLDER = 'uploads'

# Criar os diretórios se não existirem
if not os.path.exists(DOWNLOADS_FOLDER):
    os.makedirs(DOWNLOADS_FOLDER)
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Definir pesos para avaliações 360º e 180º
pesos_360 = {
    (0, 1, 1, 1): [0.0, 0.555555556, 0.222222222, 0.222222222],
    (1, 0, 1, 1): [0.2, 0.0, 0.4, 0.4],
    (1, 1, 0, 1): [0.125, 0.625, 0.0, 0.25],
    (1, 1, 1, 0): [0.125, 0.625, 0.25, 0.0],
    (1, 0, 0, 1): [0.333333333, 0.0, 0.0, 0.666666667],
    (1, 0, 1, 0): [0.333333333, 0.0, 0.666666667, 0.0],
    (0, 1, 0, 1): [0.0, 0.714285714, 0.0, 0.285714286],
    (0, 1, 1, 0): [0.0, 0.714285714, 0.285714286, 0.0],
    (1, 1, 0, 0): [0.166666667, 0.833333333, 0.0, 0.0],
    (0, 0, 1, 1): [0.0, 0.0, 0.5, 0.5],
    (0, 0, 0, 1): [0.0, 0.0, 0.0, 1.0],
    (0, 0, 1, 0): [0.0, 0.0, 1.0, 0.0],
    (0, 1, 0, 0): [0.0, 1.0, 0.0, 0.0],
    (1, 0, 0, 0): [1.0, 0.0, 0.0, 0.0],
    (0, 0, 0, 0): [0.0, 0.0, 0.0, 0.0],
    (1, 1, 1, 1): [0.1, 0.5, 0.2, 0.2]
}

pesos_180 = {
    (1, 1, 0, 0): [0.1, 0.9, 0.0, 0.0],
    (1, 0, 0, 0): [1.0, 0.0, 0.0, 0.0],
    (0, 1, 0, 0): [0.0, 1.0, 0.0, 0.0],
    (0, 0, 0, 0): [0.0, 0.0, 0.0, 0.0]
}

pesos_nota_final = {
    "Comportamentos da Função - Coordenador Pedagógico e Supervisor Bilíngue":
    11,
    "Comportamentos da Função - Diretor de Ensino": 11,
    "Comportamentos da Função - Diretor de Unidade": 12,
    "Comportamentos da Função - Diretor Geral": 11,
    "Comportamentos da Função - Diretor Regional/Operações": 11,
    "Comportamentos da Função - Trainee": 3,
    "Comportamentos da Função - BP": 9,
    "Comportamentos de Liderança": 9,
    "Comportamentos de Liderança Externa": 2,
    "Comportamentos Grupo Salta": 12
}


# Função para contar o número de avaliações por tipo e determinar os pesos aplicados
def contar_avaliacoes_e_pesos(grupo, tipo_avaliacao):
    contagem = {
        'Autoavaliação (Contagem)':
        len(grupo[grupo['Relacionamento atual na avaliação'] ==
                  'Autoavaliação']),
        'Líder (Contagem)':
        len(grupo[grupo['Relacionamento atual na avaliação'] ==
                  'Líder > Liderado']),
        'Pares (Contagem)':
        len(grupo[grupo['Relacionamento atual na avaliação'] == 'Pares']),
        'Liderados (Contagem)':
        len(grupo[grupo['Relacionamento atual na avaliação'] ==
                  'Liderado > Líder'])
    }

    tipos_avaliacao = (int(contagem['Autoavaliação (Contagem)'] > 0),
                       int(contagem['Líder (Contagem)'] > 0),
                       int(contagem['Pares (Contagem)'] > 0),
                       int(contagem['Liderados (Contagem)'] > 0))

    # Verifica o tipo de avaliação (360 ou 180) e aplica os pesos correspondentes
    if tipo_avaliacao == '180':
        pesos_aplicados = pesos_180.get(tipos_avaliacao, [0, 0, 0, 0])
    else:
        pesos_aplicados = pesos_360.get(tipos_avaliacao, [0, 0, 0, 0])

    # Adiciona os pesos ao dicionário de contagem
    contagem.update({
        'Autoavaliação (Peso %)': pesos_aplicados[0],
        'Líder (Peso %)': pesos_aplicados[1],
        'Pares (Peso %)': pesos_aplicados[2],
        'Liderados (Peso %)': pesos_aplicados[3]
    })

    return pd.Series(contagem)


# Função para calcular a média por pergunta (comportamento)
def calcular_media_por_pergunta(df, tipo_avaliacao):
    resultado_pesos = df.groupby(['Nome', 'Email', 'Nome (Pergunta)']).apply(
        lambda g: contar_avaliacoes_e_pesos(g, tipo_avaliacao)).reset_index()

    # Calcula a média de nota para cada tipo de avaliação
    medias_autoavaliacao = df[
        df['Relacionamento atual na avaliação'] == 'Autoavaliação'].groupby([
            'Nome', 'Email', 'Nome (Pergunta)'
        ])['Nota'].mean().reset_index(name='Média Autoavaliacao')
    medias_lider = df[
        df['Relacionamento atual na avaliação'] == 'Líder > Liderado'].groupby(
            ['Nome', 'Email',
             'Nome (Pergunta)'])['Nota'].mean().reset_index(name='Média Lider')
    medias_pares = df[
        df['Relacionamento atual na avaliação'] == 'Pares'].groupby(
            ['Nome', 'Email',
             'Nome (Pergunta)'])['Nota'].mean().reset_index(name='Média Pares')
    medias_liderados = df[df['Relacionamento atual na avaliação'] ==
                          'Liderado > Líder'].groupby([
                              'Nome', 'Email', 'Nome (Pergunta)'
                          ])['Nota'].mean().reset_index(name='Média Liderados')

    # Junta todas as médias calculadas no resultado de médias principal
    resultado_medias = df.groupby(['Nome', 'Email', 'Nome (Pergunta)'
                                   ])['Nota'].mean().reset_index(name='Média')

    # Merge com as médias de cada tipo de relacionamento
    resultado_medias = resultado_medias.merge(
        medias_autoavaliacao,
        on=['Nome', 'Email', 'Nome (Pergunta)'],
        how='left')
    resultado_medias = resultado_medias.merge(
        medias_lider, on=['Nome', 'Email', 'Nome (Pergunta)'], how='left')
    resultado_medias = resultado_medias.merge(
        medias_pares, on=['Nome', 'Email', 'Nome (Pergunta)'], how='left')
    resultado_medias = resultado_medias.merge(
        medias_liderados, on=['Nome', 'Email', 'Nome (Pergunta)'], how='left')

    # Substitui valores NaN por 0 para evitar erros no cálculo da média ponderada
    resultado_medias[['Média Autoavaliacao', 'Média Lider', 'Média Pares', 'Média Liderados']] = \
        resultado_medias[['Média Autoavaliacao', 'Média Lider', 'Média Pares', 'Média Liderados']].fillna(0)

    # Mescla os resultados de pesos e médias para obter o DataFrame completo
    df_merged = pd.merge(resultado_pesos,
                         resultado_medias,
                         on=['Nome', 'Email', 'Nome (Pergunta)'])

    # Calcular a média ponderada considerando os pesos de cada tipo de avaliação
    df_merged['Média Ponderada Pergunta'] = (
        (df_merged['Média Autoavaliacao'] *
         df_merged['Autoavaliação (Peso %)']) +
        (df_merged['Média Lider'] * df_merged['Líder (Peso %)']) +
        (df_merged['Média Pares'] * df_merged['Pares (Peso %)']) +
        (df_merged['Média Liderados'] * df_merged['Liderados (Peso %)']))

    # Retornar todas as colunas necessárias, incluindo as contagens, pesos e as médias simples
    return df_merged[[
        'Nome', 'Email', 'Nome (Pergunta)', 'Autoavaliação (Contagem)',
        'Líder (Contagem)', 'Pares (Contagem)', 'Liderados (Contagem)',
        'Autoavaliação (Peso %)', 'Líder (Peso %)', 'Pares (Peso %)',
        'Liderados (Peso %)', 'Média', 'Média Ponderada Pergunta',
        'Média Autoavaliacao', 'Média Lider', 'Média Pares', 'Média Liderados'
    ]]


# Função para calcular a média por tópico (grupo de comportamento)
    # Função para calcular a média por tópico (grupo de comportamento)
def calcular_media_por_topico(df, tipo_avaliacao):
        resultado_pesos = df.groupby(['Nome', 'Email', 'Nome (Tópico)']).apply(
            lambda g: contar_avaliacoes_e_pesos(g, tipo_avaliacao)).reset_index()

        medias_autoavaliacao = df[df['Relacionamento atual na avaliação'] == 'Autoavaliação'].groupby(
            ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Autoavaliacao')
        medias_lider = df[df['Relacionamento atual na avaliação'] == 'Líder > Liderado'].groupby(
            ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Lider')
        medias_pares = df[df['Relacionamento atual na avaliação'] == 'Pares'].groupby(
            ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Pares')
        medias_liderados = df[df['Relacionamento atual na avaliação'] == 'Liderado > Líder'].groupby(
            ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Liderados')

        resultado_medias = df.groupby(['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média')
        resultado_medias = resultado_medias.merge(medias_autoavaliacao, on=['Nome', 'Email', 'Nome (Tópico)'], how='left')
        resultado_medias = resultado_medias.merge(medias_lider, on=['Nome', 'Email', 'Nome (Tópico)'], how='left')
        resultado_medias = resultado_medias.merge(medias_pares, on=['Nome', 'Email', 'Nome (Tópico)'], how='left')
        resultado_medias = resultado_medias.merge(medias_liderados, on=['Nome', 'Email', 'Nome (Tópico)'], how='left')

        # Preenche valores NaN com 0 para evitar erros no cálculo da média ponderada
        resultado_medias[['Média Autoavaliacao', 'Média Lider', 'Média Pares', 'Média Liderados']] = \
            resultado_medias[['Média Autoavaliacao', 'Média Lider', 'Média Pares', 'Média Liderados']].fillna(0)

        df_merged = pd.merge(resultado_pesos, resultado_medias, on=['Nome', 'Email', 'Nome (Tópico)'])

        df_merged['Média Ponderada Tópico'] = (
            (df_merged['Média Autoavaliacao'] * df_merged['Autoavaliação (Peso %)']) +
            (df_merged['Média Lider'] * df_merged['Líder (Peso %)']) +
            (df_merged['Média Pares'] * df_merged['Pares (Peso %)']) +
            (df_merged['Média Liderados'] * df_merged['Liderados (Peso %)'])
        )

        return df_merged
def calcular_media_interfaces(df, tipo_avaliacao):
    # 1. Calcule as médias das avaliações de Pares e Liderados para cada tópico individualmente
    medias_pares = df[df['Relacionamento atual na avaliação'] == 'Pares'].groupby(
        ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Pares')
    medias_liderados = df[df['Relacionamento atual na avaliação'] == 'Liderado > Líder'].groupby(
        ['Nome', 'Email', 'Nome (Tópico)'])['Nota'].mean().reset_index(name='Média Liderados')

    # 2. Mescle as médias de Pares e Liderados para garantir que cada tópico tenha suas médias calculadas separadamente
    df_interfaces = pd.merge(medias_pares, medias_liderados, on=['Nome', 'Email', 'Nome (Tópico)'], how='outer')

    # 3. Calcule a média das interfaces (Pares e Liderados) para cada tópico
    # Se um tipo de avaliação está ausente, considera apenas o outro tipo de avaliação disponível
    def calcular_media_interface(row):
        if pd.notna(row['Média Pares']) and pd.notna(row['Média Liderados']):
            return (row['Média Pares'] + row['Média Liderados']) / 2
        elif pd.notna(row['Média Pares']):
            return row['Média Pares']
        elif pd.notna(row['Média Liderados']):
            return row['Média Liderados']
        return None  # Caso em que não há nenhuma avaliação para o tópico

    df_interfaces['Media_Liderados_Pares'] = df_interfaces.apply(calcular_media_interface, axis=1)

    # 4. Adicione o peso específico para cada tópico, usando o dicionário de pesos
    df_interfaces['Peso Específico'] = df_interfaces['Nome (Tópico)'].map(pesos_nota_final)

    # Verificação intermediária - exiba os valores após o cálculo de médias e aplicação de pesos
    print("Médias por Tópico e Pesos Aplicados:\n", df_interfaces[['Nome', 'Email', 'Nome (Tópico)', 'Media_Liderados_Pares', 'Peso Específico']])

    # 5. Calcule a média ponderada para cada tópico (média das interfaces multiplicada pelo peso do tópico)
    df_interfaces['Média Ponderada Tópico'] = df_interfaces['Media_Liderados_Pares'] * df_interfaces['Peso Específico']

    # Verificação intermediária - exiba o DataFrame com médias ponderadas por tópico
    print("Médias Ponderadas por Tópico:\n", df_interfaces[['Nome', 'Email', 'Nome (Tópico)', 'Média Ponderada Tópico']])

    # 6. Agrupe por colaborador para calcular a média final ponderada
    # A média final é a soma das médias ponderadas dos tópicos dividida pela soma dos pesos dos tópicos
    df_resultado_final = df_interfaces.groupby(['Nome', 'Email']).apply(lambda x: pd.Series({
        'Peso Total Grupos': x['Peso Específico'].sum(),
        'Média Final Ponderada': x['Média Ponderada Tópico'].sum() / x['Peso Específico'].sum()
    })).reset_index()

    # Verificação final - exiba o resultado final por colaborador
    print("Resultado Final por Colaborador:\n", df_resultado_final)

    return df_resultado_final

    # Uso correto nas chamadas
    # Em vez de df_topico, use df_merged para armazenar o resultado de calcular_media_por_topico



# Função para calcular a média final
def calcular_media_final(df_topico):
    df_topico['Peso Específico'] = df_topico['Nome (Tópico)'].map(
        pesos_nota_final)
    df_topico['Média Ponderada Final Grupo'] = df_topico[
        'Média Ponderada Tópico'] * df_topico['Peso Específico']
    df_final = df_topico.groupby(['Nome', 'Email']).apply(lambda x: pd.Series({
        'Autoavaliacao_Contagem':
        x['Autoavaliação (Contagem)'].sum(),
        'Lider_Contagem':
        x['Líder (Contagem)'].sum(),
        'Pares_Contagem':
        x['Pares (Contagem)'].sum(),
        'Liderados_Contagem':
        x['Liderados (Contagem)'].sum(),
        'Autoavaliacao_Peso':
        x['Autoavaliação (Peso %)'].mean(),
        'Lider_Peso':
        x['Líder (Peso %)'].mean(),
        'Pares_Peso':
        x['Pares (Peso %)'].mean(),
        'Liderados_Peso':
        x['Liderados (Peso %)'].mean(),
        'Peso Total Grupos':
        x['Peso Específico'].sum(),
        'Média Final':
        x['Média Ponderada Final Grupo'].sum() / x['Peso Específico'].sum()
    })).reset_index()

    return df_final




@app.route('/')
def index():
    return render_template('index.html')


@app.route('/processar', methods=['GET', 'POST'])
def processar():
    try:
        if 'file_avaliacoes' not in request.files:
            return "Nenhum arquivo foi enviado. Por favor, selecione um arquivo e tente novamente.", 400

        file_avaliacoes = request.files['file_avaliacoes']
        tipo_avaliacao = request.form.get('tipo_avaliacao')
        nivel_calculo = request.form.get('nivel')

        if file_avaliacoes.filename == '':
            return "Nenhum arquivo foi selecionado. Por favor, escolha um arquivo e tente novamente.", 400

        filename = secure_filename(
            str(uuid.uuid4()) + "_" + file_avaliacoes.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file_avaliacoes.save(file_path)

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            logging.error("Erro ao carregar o arquivo Excel", exc_info=True)
            return "Erro ao carregar o arquivo. Verifique se é um arquivo Excel válido.", 400

        if nivel_calculo == 'Por Competência':
            df_resultado = calcular_media_por_pergunta(
                pd.read_excel(file_path), tipo_avaliacao)
        elif nivel_calculo == 'Por Grupo de Competência':
            df_resultado = calcular_media_por_topico(pd.read_excel(file_path),
                                                     tipo_avaliacao)
        elif nivel_calculo == 'Calcular Média Final':
            # Calcula primeiro a média por tópico e depois a média final
            df_topico = calcular_media_por_topico(pd.read_excel(file_path),
                                                  tipo_avaliacao)
            df_resultado = calcular_media_final(df_topico)
        elif nivel_calculo == 'Cálculo de Interfaces':
            # Chama a função específica para o cálculo de interfaces
            df_resultado = calcular_media_interfaces(pd.read_excel(file_path), tipo_avaliacao)

        else:
            print("Opção de cálculo não corresponde às condições.")
            return "Opção de cálculo inválida"

        output_file = os.path.join(DOWNLOADS_FOLDER,
                                   'resultado_media_ponderada.xlsx')
        df_resultado.to_excel(output_file, index=False)
        os.remove(file_path)

        return redirect(
            url_for('sucesso', filename=os.path.basename(output_file)))

    except Exception as e:
        logging.error("Erro durante o processamento", exc_info=True)
        return "Ocorreu um erro durante o processamento. Por favor, verifique o log para mais detalhes.", 500


@app.route('/sucesso')
def sucesso():
    filename = request.args.get('filename')
    return render_template('sucesso.html', filename=filename)


@app.route('/download/<filename>')
def download(filename):
    try:
        return send_file(os.path.join(DOWNLOADS_FOLDER, filename),
                         as_attachment=True)
    except Exception as e:
        logging.error(f"Erro durante o download: {e}")
        return "Erro durante o download do arquivo."


if __name__ == '__main__':
    port = 5000
    host = '127.0.0.1'  # Define o host como localhost
    print(f"Servidor rodando em http://{host}:{port}")
    app.run(host=host, port=port)
