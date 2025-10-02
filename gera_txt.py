def generate_txts_from_xls(shopping_escolhido, tipo_faturamento):
    import os
    import pyexcel as p
    from datetime import date
    import glob
    import unicodedata
    import shutil


    def remove_accents(input_str):
        nfkd_form = unicodedata.normalize('NFKD', input_str)
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    shopping_siglas = {
        "Shopping Montserrat": "SMS",
        "Shopping da Ilha": "SDI",
        "Shopping Rio Poty": "SRP",
        "Shopping Metrópole": "SMT",
        "Shopping Moxuara": "SMO",
        "Shopping Mestre Álvaro": "SMA"
    }

    meses_portugues = {
        1: "JANEIRO",
        2: "FEVEREIRO",
        3: "MARÇO",
        4: "ABRIL",
        5: "MAIO",
        6: "JUNHO",
        7: "JULHO",
        8: "AGOSTO",
        9: "SETEMBRO",
        10: "OUTUBRO",
        11: "NOVEMBRO",
        12: "DEZEMBRO"
    }

    hoje = date.today()
    mes_atual = hoje.month
    ano_atual = hoje.year

    if "post" in tipo_faturamento.lower():
        folder_type = "POSTECIPADO"
    elif "ante" in tipo_faturamento.lower():
        folder_type = "ANTECIPADO"
    else:
        folder_type = "ATÍPICAS"

    ano_path = ano_atual
    mes_path = mes_atual

    nome_mes = meses_portugues.get(mes_path, "JANEIRO")
    sigla = shopping_siglas.get(shopping_escolhido, "ILHA")

    pasta_cargas = fr"\\192.168.18.2\csc.financeiro\C-Faturamento e CR\RPA\{sigla}\{ano_path}\{nome_mes}\{folder_type}\Planilha de Cargas"
    output_dir = fr"\\192.168.18.2\csc.financeiro\C-Faturamento e CR\RPA\{sigla}\{ano_path}\{nome_mes}\{folder_type}\Arquivos Cargas"

    print(output_dir)

    # Cria pasta de saída se não existir
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # Remove arquivos existentes na pasta
    for item in os.listdir(output_dir):
        path = os.path.join(output_dir, item)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception as e:
            print(f"Não foi possível remover {path}: {e}")

    arquivos_planilha = glob.glob(os.path.join(pasta_cargas, f"{sigla}_Importar_Encargos*.xls*"))
    if not arquivos_planilha:
        print(f"Nenhum arquivo Excel encontrado em {pasta_cargas} que inicie com {sigla}_Importar_Encargos")
        return output_dir, 0

    primeiro_arquivo = arquivos_planilha[0]
    sheets = p.get_book(file_name=primeiro_arquivo)

    file_count = 0

    for sheet_name in sheets.sheet_names():
        sheet = sheets[sheet_name]
        data = list(sheet)
        # ignora abas sem conteúdo
        if not data:
            print(f"A aba {sheet_name} está vazia e será ignorada.")
            continue

        headers = data[0]
        if ('AINC;MINC;NLUC;LOJA;NSEQLOJ;NCONREDPLC' in headers
            and 'VALOR' in headers):
            col_index = headers.index('AINC;MINC;NLUC;LOJA;NSEQLOJ;NCONREDPLC')
            valor_index = headers.index('VALOR')
            column_data = [
                row[col_index]
                for row in data[1:]
                if len(row) > max(col_index, valor_index)
                and isinstance(row[valor_index], (int, float))
                and row[valor_index] > 0
            ]

            # Apenas cria arquivo se houver dados
            if column_data:
                normalized_sheet_name = remove_accents(sheet_name)
                txt_file_name = os.path.join(output_dir, f"{normalized_sheet_name}.txt")
                with open(txt_file_name, 'w', encoding='utf-8') as txt_file:
                    for line in column_data:
                        txt_file.write(f"{line}\n")
                file_count += 1
                print(f"Arquivo {txt_file_name} criado com sucesso.")
            else:
                print(f"A aba {sheet_name} não tem dados válidos e será ignorada.")
        else:
            print(f"A aba {sheet_name} não possui as colunas necessárias.")

    return output_dir, file_count
