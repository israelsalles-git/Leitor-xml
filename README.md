# Monitor de XML para Excel - Controle de Notas Fiscais
Aplicação para monitorar uma pasta com arquivos XML de notas fiscais, extrair informações relevantes e gerar um relatório consolidado em Excel, com controle de notas duplicadas.

Funcionalidades
Monitoramento automático de pasta em busca de novos arquivos XML

Extração dos seguintes dados das notas fiscais:

Número da nota fiscal

Data de emissão

Nome do cliente

Valor total da nota

Controle de notas repetidas (atualiza dados existentes)

Geração de relatório em Excel consolidado

Interface gráfica amigável com CustomTkinter

Requisitos
Python 3.8+

Bibliotecas necessárias (instaladas automaticamente com requirements.txt):

customtkinter

pandas

watchdog

openpyxl

lxml

Instalação
Clone o repositório:

bash
git clone https://github.com/seu-usuario/xml-monitor.git
cd xml-monitor
Crie e ative um ambiente virtual (recomendado):

bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate  # Windows
Instale as dependências:

bash
pip install -r requirements.txt
Uso
Execute o aplicativo com:

bash
python xml_monitor.py
Na interface:

Selecione a pasta para monitorar (contendo os XMLs)

Defina o arquivo Excel de saída

Clique em "Iniciar Monitoramento" para começar

Configurações
Monitoramento automático: A aplicação detecta novos arquivos XML adicionados à pasta

Processamento manual: Botão para processar todos os XMLs existentes na pasta

Atualização de notas: Se uma nota já existir no relatório, seus dados serão atualizados

Estrutura do Projeto
text
xml-monitor/
├── xml_monitor.py        # Código principal da aplicação
├── requirements.txt      # Dependências do projeto
├── README.md             # Este arquivo
└── relatorio_notas.xlsx  # Arquivo de saída (gerado automaticamente)
Contribuição
Contribuições são bem-vindas! Siga os passos:

Faça um fork do projeto

Crie uma branch para sua feature (git checkout -b feature/incrivel)

Commit suas mudanças (git commit -m 'Adiciona feature incrível')

Push para a branch (git push origin feature/incrivel)

Abra um Pull Request

Licença
Distribuído sob a licença MIT. Veja LICENSE para mais informações.

Contato
Israel salles de oliveira - sallesisrael66@gmail.com
