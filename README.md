# Envio de Mensagens Automáticas via WhatsApp Web
Este projeto tem como objetivo automatizar o envio de mensagens personalizadas para uma lista de clientes extraída de uma planilha Excel. Utilizando a biblioteca openpyxl para ler dados de uma planilha, webbrowser para abrir o WhatsApp Web e pyautogui para automatizar o envio de mensagens, o programa facilita o envio em massa de mensagens personalizadas para números de telefone armazenados na planilha.

Funcionalidades
Abrir planilha de clientes: Através de uma interface gráfica, o usuário pode selecionar uma planilha .xlsx contendo dados de clientes (nome, telefone e data de vencimento).
Exibição de dados: Após carregar a planilha, os dados dos clientes são exibidos em uma tabela dentro da interface.
Envio de mensagens personalizadas: O usuário pode digitar uma mensagem personalizada e, ao clicar no botão de envio, as mensagens serão enviadas automaticamente através do WhatsApp Web.
Controle de erros: Caso ocorra algum erro durante o envio de uma mensagem, o programa registra o erro em um arquivo CSV e continua com os demais clientes.
Calendário: Um calendário é integrado à interface para que o usuário possa marcar a data de envio das mensagens.

Tecnologias Utilizadas:

Tkinter para a criação da interface gráfica.

openpyxl para leitura de planilhas Excel.

pyautogui para automação da interação com o WhatsApp Web.

webbrowser para abrir links do WhatsApp Web.

tkcalendar para integração do calendário.

Observações:

O usuário deve garantir que o WhatsApp Web esteja configurado corretamente e totalmente carregado antes de iniciar o envio das mensagens.
A automatização requer que o navegador esteja aberto e que o WhatsApp Web esteja carregado para o envio ser realizado sem interrupções.

Requisitos
Python 3.x
Bibliotecas:
openpyxl
pyautogui
tkinter
tkcalendar

Para instalar as dependências, execute:

pip install openpyxl pyautogui tkcalendar
