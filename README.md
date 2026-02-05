# Ghost of Gov

Automação de processos para a plataforma SouGov.sp.gov.br, desenvolvida em Python,
com foco em **execução confiável, padronização operacional e redução de tarefas manuais**.

O projeto está em evolução para um **sistema web multiusuário**, com **filas de execução**
e suporte a **múltiplos logins**, deixando de ser uma automação local.

---

## Visão geral

O Ghost of Gov automatiza o fluxo de acesso, navegação e download de holerites
(contracheques) da plataforma SouGov.sp.gov.br.

A automação simula interações humanas no navegador utilizando Playwright,
executando tarefas repetitivas de forma **controlada, previsível e reproduzível**,
sem armazenar credenciais sensíveis.

Inicialmente desenvolvido como automação local, o projeto está sendo evoluído
para um **serviço centralizado**, com foco em escalabilidade e operação contínua.

---

## Problema resolvido

Antes da automação:
- Login manual recorrente
- Navegação repetitiva entre páginas
- Download manual de múltiplos documentos
- Alto risco de erro humano
- Baixa previsibilidade operacional

Com a automação:
- Processo padronizado
- Execução assistida e automatizada
- Redução de falhas humanas
- Ganho de eficiência operacional
- Base preparada para execução em larga escala

---

## Como funciona atualmente

1. Inicializa um navegador controlado via Playwright  
2. Solicita login manual inicial (por segurança)  
3. Mantém sessão ativa por perfil persistente  
4. Permite seleção de vínculo via terminal  
5. Seleciona dinamicamente anos e meses disponíveis  
6. Realiza downloads automatizados de PDFs  
7. Trata falhas e gera evidências (screenshots) em caso de erro  

A execução ocorre via terminal, com instruções claras para o operador.

---

## Funcionalidades atuais

- Automação completa de download de holerites em PDF  
- Seleção dinâmica de vínculos  
- Download de múltiplos anos e meses  
- Navegação inteligente por dropdowns  
- Persistência de sessão via perfil de navegador  
- Tratamento básico de erros e exceções  
- Captura automática de screenshots em falhas  
- Configuração flexível de diretórios e tempos de espera  

---

## Stack utilizada

- Python  
- Playwright  
- Chrome DevTools Protocol  
- PyFiglet  

---

## Estrutura do projeto (atual)

```text
ghost_of_gov/
├── ghost_of_gov.py
├── hybrid_chrome_profile/
├── holerites_baixados/
├── requirements.txt
└── README.md
```

---

## Execução local

-pip install playwright pyfiglet colorama
-playwright install
-python ghost_of_gov.py

---

## Durante a execução:

-O login é realizado manualmente
-O controle retorna automaticamente ao script
-Os arquivos são salvos localmente

---

## Considerações de segurança:

-Credenciais não são armazenadas em código
-Login manual por decisão de segurança
-Sessões persistem apenas localmente
-O projeto respeita os limites e comportamento da plataforma automatizada

---

## Evolução planejada (roadmap):

O projeto está sendo evoluído para deixar de ser uma automação local e se tornar
um sistema de automação centralizado, com foco em operações e escalabilidade.

---

## Próximas etapas planejadas:

-Interface web para gerenciamento de execuções
-Suporte a múltiplos usuários
-Sistema de filas para downloads
-Execução em background
-Padronização de logs
-Métricas de execução
-Containerização do serviço
-Integração com CI/CD

---

## Contexto

Projeto pessoal desenvolvido como parte da minha transição para a área de
DevOps / SRE, com foco em automação de processos, confiabilidade e operação
de sistemas.

---

© 2025 Paulo Eduardo Moretti. Todos os direitos reservados.
