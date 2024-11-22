from pathlib import Path

# Caminhos dos arquivos
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / 'data'
FORMULARIO_ORIGEM = DATA_DIR / '../data/Dados do cliente (USAR PARA PREENCHER).xlsx'
FORMULARIO_DESTINO = DATA_DIR / '../data/Formulário Vazio (DEVE SER PREENCHIDO).xlsx'
# Mapeamento de células (origem: nome do campo)
MAPEAMENTO = {
    'B7': 'Número do Cliente',
    'F7': 'Conta Contrato',
    'B8': 'Titular da UC',
    'P8': 'Nascimento',
    'B9': 'CPF / CNPJ',
    'L9': 'RG',
    'P9': 'Data Emissão',
    'B10': 'Endereço',
    'P10': 'Número',
    'B11': 'Bairro',
    'L11': 'Município',
    'B12': 'CEP',
    'L12': 'Estado',
    'P12': 'Complemento',
    'B13': 'Grupo',
    'L13': 'Subgrupo',
    'P13': 'Classe',
    'B16': 'Latitude',
    'L16': 'Latitude Grau',
    'P16': 'Área',
    'B17': 'Longitude',
    'L17': 'Longitude Grau',
    'P17': 'Abscissa (X)',
    'P18': 'Ordenada (Y)',
    'F3': 'Projeto',
    'P3': 'Concessionária',
    'B4': 'Integrador',
    'L4': 'E-mail',
    'P4': 'Telefone',
    'B21': 'Modalidade',          # 2.1 MODALIDADE
    'B29': 'Tipo de Conexão',     # 3.1 TIPO DE CONEXÃO
    'F29': 'Tipo de Ramal',       # 3.2 TIPO DE RAMAL
    'L29': 'Tensão',              # 3.3 TENSÃO
    'B30': 'Disjuntor',           # 3.4 DISJUNTOR
    'F30': 'Tipo',                # 3.5 TIPO
    'L30': 'Carga',               # 3.6 CARGA
    'B31': 'Fase/Neutro',         # 3.7 FASE/NEUTRO
    'L31': 'Terra',               # 3.8 TERRA
    'B33': 'Tipo de Solicitação', # 3.9 TIPO DE SOLICITAÇÃO
    'B34': 'Disjuntor Anterior',  # 3.10 DISJUNTOR ANTERIOR
    'B35': 'Observação',          # 3.11 OBSERVAÇÃO
    'B38': 'Fabricante',          # 4.1 FABRICANTE
    'F38': 'Modelo',              # 4.2 MODELO
    'B39': 'Quantidade',          # 4.3 QUANTIDADE
    'F39': 'Potência (Wp)',       # 4.4 POTÊNCIA (Wp)
    'L39': 'Potência Total (kWp)',# 4.5 POTÊNCIA TOTAL (kWp)
    'B40': 'Isc (A)',             # 4.6 Isc (A)
    'F40': 'Vcco (V)',            # 4.7 Vcco (V)
    'B41': 'Altura (mm)',         # 4.8 ALTURA (mm)
    'F41': 'Largura (mm)',        # 4.9 LARGURA (mm)
    'L41': 'Área (m²)',           # 4.10 ÁREA (m²)
}