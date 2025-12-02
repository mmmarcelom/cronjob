Essa é uma solução que criei para automatizar um relatório para um cliente.  
Coloquei ele aqui como portfólio, demonstrando conhecimentos sobre APIs REST, Github actions, google scripts, webhooks e etc.
O trabalho todo foi feito em aproximadamente 4 horas. 
Eu utilizando o ChatGPT para fazer o heavy lifting, mas todo o código foi revisado e adaptado.

Usei nesse projeto:
- Python e selenium
- Javascript no google scripts e as classes do Google Docs
- Github Actions

## Meu problema era o seguinte:

1. O aplicativo de ERP não tinha uma API aberta.
2. O relatório que era criado não tinha o formato adequado.

Relatório criado:
<img width="1150" height="552" alt="image" src="https://github.com/user-attachments/assets/1a7927fc-2bc7-48ec-a8c3-232cd83437b9" />

Relatório desejado:
<img width="929" height="538" alt="image" src="https://github.com/user-attachments/assets/8f0284f5-280c-43f0-ab8d-e08263bdf7ec" />

## Solução:

Usar o google scripts para automatizar.  

## Desafios:
- Fazer login e pegar o token da API de relatórios
- Converter o xlsx e fazer o upload do documento gerado
- Formatar o documento da maneira adequada
- Tornar tudo isso automático

## Login

Como o aplicativo não retornava o user token quando eu fazia login usando o google scripts, decidi utilizar um cron job no github para fazer login usando selenium e enviar o token para o script através de um webhook no próprio google scripts.  
Dessa forma, posso gerar um novo token a cada algumas horas e armazenar o token de forma relativamente segura.  
O usuário e senha estão no secrets do Github e o token está na propriedades do script, dessa maneira, só fica visivel para quem tem acesso a planilha dentro da organização.  

## Converter o xlsx

A resposta da API foi:  
````
{
   "Message":null,  
   "Success":true,  
   "Data":"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=iso-8859-1;base64,UEsDBBQAAAgIAM1wgluyUxgBvQAAACUBAAAPAAAAeGwvd29ya2Jvb2sueG1sjY/BboMwDIZfJfK9BKaKtQjoZZdeuydwwSkRJEZxaHn8ZWxqr719/m1Z31+fVjepOwWx7BsoshwU+Y57628NLNHsDnBq67V6cBivzKNK916qtYEhxrnSWrqBHErGM/m0M"  
}
````
Para transformar essa resposta em uma planilha, salvei a planilha temporariamente no drive e copiei os dados.  

## Formatar a planilha

Para essa parte bastou trabalhar o array para completar as linhas vazias com o produto da linha de cima e remover as linhas que não são usadas.  
A cada execução, eu também limpo a planilha e sobrescrevo os dados.  
Como eu utilizei a planilha formatada do google e usei a planilha ao invés do intervalo na tabela dinâmica, o sistema sempre gera a planilha corretamente.  
<img width="797" height="495" alt="image" src="https://github.com/user-attachments/assets/6ccb6693-c66c-44ea-af27-f9d5419109f9" />

## Tornar tudo automático

Para facilitar a vida do usuário, criei um menu no google scripts para atualizar tudo com um clique.
<img width="1034" height="276" alt="image" src="https://github.com/user-attachments/assets/64efcfd4-7cab-45d4-9f7f-70f4f783655f" />


