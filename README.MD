Documentação completa para o próximo desenvolvedor
Nome do projeto: Scraper TRT-2 – Pautas de Audiência + Enriquecimento PJe
Objetivo: Gerar base qualificada de empresas reclamadas em processos trabalhistas sem advogado constituído, com CNPJ e contato potencial.
Arquitetura atual (1ª fase concluída):
JTe Público → 
Lista varas → 
Seleciona vara → 
Loop datas futuras → 
Extrai pauta → 
Acumula em CSV
Próxima fase recomendada (2ª fase – impacto alto):
A partir do número do processo extraído do JTe:
Acessar o PJe público: https://pje.trt2.jus.br/consultaprocessual
Buscar pelo número do processo (com dígito verificador e formatação correta)
Extrair:
Lista de advogados da reclamada (se vazio → oportunidade)
Valor da causa
CNPJ/CPF da reclamada
Endereço/e-mail se disponível
Filtrar no final ou durante a extração:
Tem advogado? → descarta
Valor da causa muito baixo? → descarta
Nome/CNPJ em blacklist de grandes empresas? → descarta
Exportar CSV final com colunas enriquecidas:
geraEm, vara, data_audiencia, numero_processo, reclamante, reclamada, cnpj, valor_causa, tem_advogado, juiz, sessao, etc.


