# 📊 Enterprise Excel Auto-Refresher

Uma solução de automação robusta escrita em Python para orquestrar a atualização de dados (ETL) em planilhas Excel complexas (Power Query/Pivot Tables).

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/platform-windows-lightgrey)
![Library](https://img.shields.io/badge/lib-pywin32-orange)

## 🎯 O Problema

Scripts de automação Excel comuns sofrem de instabilidade: dependem de tempos de espera fixos (`sleep`), falham silenciosamente ou deixam processos "zumbis" consumindo memória RAM quando ocorrem erros.

## 💡 A Solução

Este projeto implementa um **wrapper orientado a objetos** em torno da API COM do Windows, focando em:

* **Integridade de Recursos:** Utilização do padrão *Context Manager* (`with statement`) para garantir que a instância do Excel seja encerrada corretamente e a memória liberada, mesmo em caso de falhas críticas.
* **Sincronização Inteligente:** Substituição de `time.sleep()` pelo método nativo `CalculateUntilAsyncQueriesDone()`, garantindo que o salvamento ocorra apenas após a conclusão real das consultas de dados.
* **Isolamento:** Uso de `DispatchEx` para criar instâncias separadas do Excel, permitindo que o robô trabalhe sem interferir nas planilhas que o usuário já tenha abertas.
* **Observabilidade:** Sistema de `logging` detalhado para auditoria de execução e fácil depuração.

## 🛠️ Pré-requisitos

* Sistema Operacional Windows (necessário para acesso à API COM).
* Microsoft Excel instalado.
* Python 3.x.

### Instalação das Dependências

```bash
pip install pywin32

