# 🗓️ Agenda de Consultório em Google Planilhas

Agenda compartilhada para um consultório de psicologia com várias profissionais e **uma sala só**.
Funciona dentro do Google Planilhas, com **Google Apps Script** (JavaScript), sem sistema pago e sem Google Agenda.

> Todos os dados da demonstração são **fictícios**.

📄 **[Ver a demonstração em PDF](agenda-consultorio-demo.pdf)**: calendário de outubro e novembro, pacientes, faltas e cancelamentos, psicólogas, feriados e configuração.

## O problema

Quatro psicólogas dividem a mesma sala e às vezes atendem em equipe. A agenda era controlada no papel e no WhatsApp, e ninguém conseguia ver de forma rápida:
- quem atende em cada horário,
- quais pacientes faltaram ou cancelaram, e por quê,
- quantos atendimentos houve no mês.

## A solução

Uma planilha em formato de **calendário (segunda a sábado)**. Basta digitar o paciente na célula do dia e do horário, e o script cuida do resto.

| Recurso | Como funciona |
|---|---|
| Cores por status | Agendada, Confirmada, Compareceu, Faltou e Cancelou têm cada uma sua cor |
| Cor por psicóloga | Cada profissional tem uma cor e uma sigla, por exemplo `Ana Souza (HP)` |
| Atendimento em equipe | `Ana Souza (HP+BT)` |
| Horários bloqueados | Ficam em cinza fora do horário de cada psicóloga |
| Feriados | Os nacionais (inclusive Carnaval, Sexta-feira Santa e Corpus Christi) são calculados sozinhos |
| Filtro | Cada psicóloga vê só os próprios horários |
| Faltas e cancelamentos | Registrados automaticamente, com o motivo |
| Resumo do mês | Totais, taxa de faltas, números por psicóloga e motivos mais comuns |
| Paciente fixo | Repete a consulta toda semana ou a cada 15 dias |
| Horários configuráveis | Início, fim e duração da sessão são editáveis |

## Arquivos

| Arquivo | O que é |
|---|---|
| `agenda-calendario-consultorio.gs` | Script para uso real |
| `agenda-calendario-PORTFOLIO.gs` | O mesmo script, com a opção de preencher dados fictícios |
| `agenda-consultorio-demo.pdf` | Planilha de demonstração exportada em PDF |

## Tecnologias

- Google Apps Script (JavaScript V8)
- Google Planilhas: validação de dados, formatação e gatilho `onEdit`

## Como usar

1. Crie uma planilha nova no Google Planilhas.
2. Vá em **Extensões > Apps Script** e cole o código de `agenda-calendario-consultorio.gs`.
3. Salve, escolha a função `configurarPlanilha` e clique em **Executar** (autorize o acesso).
4. Recarregue a planilha. O menu **🗓️ Consultório** vai aparecer.
5. Para ver um exemplo pronto, use o arquivo `agenda-calendario-PORTFOLIO.gs` e o menu **Preencher com dados de exemplo**.

## Privacidade

A planilha guarda só a agenda (nome, telefone e horário). Prontuário e informações clínicas ficam fora dela.
Em uso real, ela deve ser compartilhada apenas com a equipe, como **Editor**, nunca por link público.

## Autor

**Luis Alfredo Olivier Sudbrack**
Técnico em Informática (em formação) · SENAC · Brasília – DF
