
Scraper TRT-2 – Pautas de Audiência + Enriquecimento PJe
Automação de Extração de Dados Jurídicos para Prospecção Comercial

Este projeto automatiza a coleta de audiências futuras nas varas trabalhistas da capital de São Paulo (TRT-2), identificando oportunidades em processos onde a empresa reclamada ainda não possui advogado constituído.

Objetivo
Gerar uma base qualificada de empresas reclamadas em processos trabalhistas, priorizando:

Empresas médias/pequenas (excluindo grandes como bancos).
Processos sem advogado constituído na reclamada.
Inclusão de CNPJ, valor da causa e dados de contato potenciais.
Isso reduz o tempo de coleta manual (de ~1 semana para horas/dias) e melhora a precisão dos contatos comerciais.

Status Atual
Fase 1 – Concluída ✅
Extração completa das pautas de audiências futuras via JTe (sistema público de pautas).

Fase 2 – Em desenvolvimento ⏳
Enriquecimento dos processos via PJe (consulta pública por número do processo) para extrair:

Advogados da reclamada
Valor da causa
CNPJ/CPF da reclamada
Outros dados relevantes
Arquitetura Atual (Fase 1)
text
JTe Público[](https://jte.csjt.jus.br)
├── Acesso anônimo → Seleção TRT-2
├── Entrada no módulo "Pauta"
├── Listagem dinâmica de varas (filtro: Audiências 1º grau + São Paulo Zonas Central/Norte/Oeste → ~92 varas)
├── Loop por vara
│   └── Loop por datas úteis futuras (~2 meses à frente, a partir de +7 dias)
│       ├── Seleção robusta de data no calendário
│       ├── Espera estabilização da pauta
│       └── Extração de processos (número, hora/status, juiz, reclamante, reclamada)
└── Acumula tudo em um único CSV
Saída: ./output/pauta_trt2_2meses_[timestamp].csv
Colunas atuais: geradoEm, vara, data, numeroProcesso, sessao, juiz, reclamante, reclamada

Próxima Fase Recomendada (Fase 2 – Alto Impacto)
A partir do numeroProcesso extraído do JTe:

Acessar a consulta pública do PJe: https://pje.trt2.jus.br/consultaprocessual
Buscar pelo número do processo no formato CNJ unificado (25 dígitos com máscara: NNNNNNN-DD.AAAA.J.TR.OOOO)
Exemplo: 0001234-56.2025.2.02.0001
Na página de detalhes do processo, extrair:
Lista de advogados (seção "Partes" ou "Procuradores")
Valor da causa (seção "Dados do Processo")
CNPJ/CPF da reclamada (na qualificação das partes)
Endereço/e-mail se disponível
Aplicar filtros:
Se tem advogado na reclamada → descartar
Valor da causa abaixo do threshold (configurável) → descartar
Reclamada em blacklist (grandes empresas) → descartar
Exportar CSV enriquecido com colunas adicionais: cnpj, valor_causa, tem_advogado, etc.
Dicas para implementação da Fase 2:

Use Playwright para simular a busca (campo de texto + botão de consulta).
O acesso é público (sem login para consulta básica).
Há possibilidade de CAPTCHA em acessos excessivos → implementar pausas aleatórias (5-15s entre consultas).
Formato do número: padronize com máscara CNJ (função auxiliar pode ser adicionada).
Exemplos de scrapers semelhantes no GitHub (ex: bots para PJe genérico) podem servir de inspiração.
Tecnologias Utilizadas
Node.js
Playwright (com Chromium)
JavaScript puro (modularizado com funções reutilizáveis)
Como Executar (Fase 1 Atual)
Bash
# 1. Instalar dependências
npm install playwright

# 2. Instalar navegadores
npx playwright install chromium

# 3. Executar o scraper
node scraper.js
O script abre o navegador visível (headless: false) para depuração.
Tempo estimado: várias horas (92 varas × ~40 datas).
Saída na pasta ./output
Melhorias Futuras Sugeridas
Tornar configurável (arquivo .env ou CLI): período de datas, lista de varas, threshold de valor da causa, blacklist de empresas.
Modo headless + agendamento (cron job).
Paralelismo controlado (3-5 browsers simultâneos).
Logs mais detalhados + screenshots em erros.
Tratamento de "sem audiências" por vara/data.
Pontos de Atenção
O JTe e PJe são sistemas públicos, mas respeite os termos de uso (evite sobrecarga).
Mudanças no layout Angular Material podem quebrar seletores → o código atual já é robusto com retries e fallbacks.
Processos em segredo de justiça não aparecem ou têm dados restritos.
Qualquer dúvida, o código está bem comentado e modularizado. Boa continuação no desenvolvimento! 🚀

Última atualização: Janeiro 2026
