## Nome

Webpart Workhub: WHD SimpleBanner

## Versão atual da webpart:

1.0.0

## Node Version:

14.15.0

## Descrição

Webpart com um banner que encaminha para uma URL, a qual é exibida para todos ou apenas para pessoas citadas em audiência.

## TODO

- Adicionar Banner
- Ser possivel ver a alteração do Banner em tempo real
- Ser possivel adicionar uma URL de destino ao clicar no Banner
- Ser possivel adicionar audiencia

# Build

gulp clean
gulp bundle --ship
gulp deploy-azure-storage
gulp package-solution --ship

## Log de alterações

- 20/12/2022: [Gustavo]:
  Titulo: Feature Simple Banner
  Descricao: Versão 1.0.0 do novo SimpleBanner
