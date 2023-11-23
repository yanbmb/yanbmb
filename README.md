<h1>Otimizador de rotinas contábeis</h1> 

> Status do Projeto: :heavy_check_mark: CONCLUÍDO  :warning: PRECISA DE MELHORIAS! :warning:

### Tópicos 

:small_blue_diamond: [Descrição do projeto](#descrição-do-projeto)

:small_blue_diamond: [Funcionalidades](#funcionalidades)

:small_blue_diamond: [Pré-requisitos](#pré-requisitos)

:small_blue_diamond: [Funcionalidades explicadas](#funcionalidades-explicadas)


## Descrição do projeto 

<p align="justify">
  Olá pessoal, esse é meu primeiro projeto, comecei a estudar sobre Python alguns meses atrás e tive algumas ideias que me ajudariam nas rotinas do escritório de contabilidade onde atualmente trabalho. Foi um desafio e tanto!... Obviamente, o código precisa ser melhorado e otimizado, pois eu fiz com os conhecimentos iniciais que obtive dos estudos, conforme fui avançado nos estudos, consegui ir melhorando algumas coisas, mas...ainda tem muito para melhorar! Essa é a meta, certo!?... 
</p>

## Funcionalidades

:heavy_check_mark: Extrator de NF3E  

:heavy_check_mark: Emissor de recibos  

:heavy_check_mark: Extrator de pagamentos  


## Pré-requisitos

:warning: [PyCharm](https://www.jetbrains.com/pt-br/pycharm/download/?section=windows)


## Funcionalidades explicadas
<p align="center"> 1️⃣  A primeira das três funções do sistema é um extrator de NF3E (notas fiscais de energia elétrica), esse posso dizer que foi o maior desafio para mim...Vamos lá! Eu faço a contabilidade de uma empresa provedora de internet, então ela tem muitos locais de consomem energia elétrica, logo, tem muitas faturas de energia em seu CNPJ, precisamos contabilizar todas, em média são 120 faturas por mês, o sistema contábil não importa esse tipo de nota, então tive essa idéia de fazer um sistema onde vai ler o pdf dessas notas e extrair as informações para um txt de acordo com um layout específico que o sistema fiscal aceite a importação (que nesse caso o layout é do SPED). Essas 120 faturas mensais não são de apenas uma emrpesa fornecedora de energia! Aqui existem três empresas: CEMIG, EDP e ENERGISA. E os layouts dessas faturas são totalmente diferentes, tive que criar para cada uma das três empresas o jeito correto de buscar essas informações nas faturas.</p>

<p align="center"> 2️⃣  A segunda função do sistema é um emissor de recibos em pdf, onde, utilizo para emitir recibos de distribuição de lucros de duas empresas específicas para seus sócios. No código que estou disponibilizando os nomes e informações pessoais estarão alterados como forma de ser apenas um exemplo.</p>

<p align="center"> 3️⃣  A ultima função do sistema é um extrator de pagamentos, onde, utilizo para extratir informações de um pdf que contém os valores pagos para cada funcionário de uma determinada empresa, normalmente são em torno de 90 a 100 funcionários, imagina ter que lançar esses pagamentos no sistema contábil, um por um...então essa função lê o pdf e extrai as informações para um txt num layout específico que o sistema contábil que utilizo aceita.</p>



## Linguagens e libs utilizadas :books:

- Python 3
- tkinter
- customtkinter
- PyMuPDF
- num2words
- reportlab


## Desenvolvedor

[<img src="link da foto" width=115><br><sub>Yan Bernard Molino Barbosa</sub>](https://github.com/yanbmb)

## Licença 

The [MIT License]() (MIT)

Copyright :copyright: 2023 - Otimizador de rotinas contábeis



