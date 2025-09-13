import csv

def separate_fiorino_loads(csv_file_path):
    MAX_WEIGHT = 500.0
    MAX_CUBAGE = 1.5
    
    orders = []
    try:
        with open(csv_file_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t') # Assuming tab-separated based on provided data
            header = next(reader) # Read header row

            # Find column indices
            try:
                col_cod_rota = header.index('Cod_Rota')
                col_quilos_saldo = header.index('Quilos_Saldo')
                col_cubagem = header.index('Cubagem')
                col_num_pedido = header.index('Num_Pedido')
            except ValueError as e:
                print(f"Erro: Coluna não encontrada no cabeçalho do CSV: {e}")
                return

            for row in reader:
                if len(row) > max(col_cod_rota, col_quilos_saldo, col_cubagem, col_num_pedido):
                    cod_rota = row[col_cod_rota].strip()
                    if cod_rota == '11101':
                        try:
                            quilos_saldo_str = row[col_quilos_saldo].strip().replace(',', '.')
                            cubagem_str = row[col_cubagem].strip().replace(',', '.')
                            
                            quilos_saldo = float(quilos_saldo_str)
                            cubagem = float(cubagem_str)
                            num_pedido = row[col_num_pedido].strip()

                            orders.append({
                                'num_pedido': num_pedido,
                                'quilos_saldo': quilos_saldo,
                                'cubagem': cubagem
                            })
                        except ValueError:
                            # Skip rows with invalid numeric data
                            continue
    except FileNotFoundError:
        print(f"Erro: O arquivo '{csv_file_path}' não foi encontrado.")
        print("Por favor, certifique-se de que o arquivo CSV está no mesmo diretório ou forneça o caminho completo e correto.")
        return
    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo CSV: {e}")
        print("Certifique-se de que o arquivo é um CSV de texto simples e que o delimitador está correto (atualmente configurado para tabulação).")
        return

    fiorino_loads = []
    current_load = {'weight': 0.0, 'cubage': 0.0, 'orders': []}

    for order in orders:
        if (current_load['weight'] + order['quilos_saldo'] <= MAX_WEIGHT and
            current_load['cubagem'] + order['cubagem'] <= MAX_CUBAGE):
            
            current_load['weight'] += order['quilos_saldo']
            current_load['cubagem'] += order['cubagem']
            current_load['orders'].append(order['num_pedido'])
        else:
            # Start a new load
            fiorino_loads.append(current_load)
            current_load = {
                'weight': order['quilos_saldo'],
                'cubagem': order['cubagem'],
                'orders': [order['num_pedido']]
            }

    # Add the last load if it's not empty
    if current_load['orders']:
        fiorino_loads.append(current_load)

    print(f"Para a rota 11101, é possível fazer {len(fiorino_loads)} cargas de Fiorino.")
    for i, load in enumerate(fiorino_loads):
        print(f"\nCarga de Fiorino {i+1}:")
        print(f"  Peso Total: {load['weight']:.2f} kg")
        print(f"  Cubagem Total: {load['cubagem']:.2f}")
        print(f"  Pedidos: {', '.join(load['orders'])}")

if __name__ == "__main__":
    # Certifique-se de que 'PEDIDOS-USAR ESSE.csv' é um arquivo de texto simples
    # e está no mesmo diretório deste script, ou forneça o caminho completo.
    # Se você continuar tendo problemas para ler o arquivo, pode ser necessário
    # abri-lo em um editor de texto e salvá-lo como um CSV delimitado por tabulação
    # ou copiar o conteúdo e colá-lo diretamente no script (substituindo a leitura do arquivo).
    separate_fiorino_loads('PEDIDOS-USAR ESSE.csv')
