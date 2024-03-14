# Trabalho de Conclusão de Curso

Esse Web App foi desenvolvido como trabalho de conclusão de curso do curso de Engenharia Elétrica da Universidade Federal de Goiás do estudante Rodrigo Santana Esperidião. Seu propósito é realizar análises sobre o uso da energia elétrica e os custos relacionados a seu faturamento em unidades consumidoras dos grupos A e B.

No aplicativo estão implementados dois métodos de análise: a análise das cargas instaladas e a análise de uma fatura de energia.

O método da análise das cargas considera o uso das cargas inseridas no aplicativo e os horários de sua utilização para estimar o consumo e a demanda da unidade em um dia. Esse método não considera mudanças na utilização das cargas ao longo do mês ou do ano, estimando então que a utilização dos equipamentos será a mesma todos os dias.

O método da análise de uma fatura, por sua vez, leva em conta o histórico de utilização de energia dos últimos 12 meses da unidade consumidora. Esse método consegue então estimar melhor as alterações sazonais do uso da energia, considerando que o perfil analisado será semelhante aos dos anos subsequentes. Esse método, porém, não pode ser empregado para a análise de unidades do grupo B, uma vez que não são informados os dados horários de consumo nas faturas das unidades enquadradas na modalidade convencional.
