# Controle-de-Entrada
📊 Sistema de Controle de Pedidos e Envios – Automação de Bases
Sistema desenvolvido em Python com foco em automação da gestão de pedidos e entregas de múltiplas bases de clientes. Ideal para setores de logística, cobrança, controle de retirada e atendimento comercial, o projeto permite tratar dados de diferentes planilhas de forma padronizada, com leitura inteligente, filtros dinâmicos, geração de relatórios e possibilidade de envio automatizado de e-mails.

⚙️ Funcionalidades principais
📁 Suporte a múltiplos clientes e bases, com organização por módulos.

📊 Leitura de planilhas .xlsx com tratamento automático de colunas e datas.

📌 Filtro de pedidos não entregues, fora do prazo ou com status específico.

📤 Integração com Outlook para envio de cobranças automatizadas (opcional).

🌒 Interface com tema escuro e componentes estilizados em Tkinter.

🧠 Modularizado para facilitar manutenção e reaproveitamento de funções.

🗂 Estrutura
arduino
Copiar
Editar
modulos/
│
├── cliente_1/
│   ├── config/
│   │   └── config_cliente1.py
│   └── cliente1_nf.py
│
├── cliente_2/
│   ├── config/
│   │   └── config_cliente2.py
│   └── cliente2_nf.py
│
└── utils/
    ├── filtros.py
    ├── conversoes.py
    └── interface.py
🚀 Tecnologias usadas
Python 3.10+

pandas

tkinter

openpyxl

win32com (Outlook Automation)

🔧 Como usar
Defina o módulo do cliente com suas configurações específicas.

Execute o script principal de leitura/automação.

Os dados serão processados automaticamente com base na configuração de cada base.

👨‍💼 Ideal para
Empresas ou profissionais que precisam gerenciar grandes volumes de pedidos com automações como:

Controle de entregas por base e cliente.

Identificação de pedidos vencidos ou não retirados.

Envio de e-mails de cobrança por base.

Geração de relatórios para promotores ou responsáveis.


