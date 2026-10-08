# Processo de Deploy

Existem dois fluxos: o **fluxo correto (padrão)** e o **fluxo atual** (usado hoje na Solicitação de Materiais).

## Regra de aprovação

Todas as Pull Requests abertas por **Rafael** ou **Gustavo** exigem aprovação simultânea de **Júlio** e **Sidney**. Sem as duas aprovações, a PR não é mesclada.

---

## 1. Fluxo correto (padrão)

```
bugfix/<nome> | feature/<nome>
        |
        v   PR
    homolog      -> testes / homologação
        |
        v   PR (após homologado)
   development
        |
        v   PR (após aprovação)
      main
        |
        v
   Servidor     -> cópia linha por linha dos arquivos editados
```

### Passos

1. Criar a branch a partir da base atualizada:
   - Correção: `bugfix/<nome-do-bug>`
   - Nova funcionalidade: `feature/<nome-da-feature>`
2. Desenvolver e commitar na branch.
3. Abrir PR da branch para `homolog`. Exige aprovação de Júlio e Sidney (quando autor for Rafael ou Gustavo).
4. Homologar a alteração em `homolog`.
5. Após homologado, abrir PR de `homolog` para `development`.
6. Após aprovação, abrir PR de `development` para `main`.
7. Com a alteração em `main`, aplicar no servidor: copiar, linha por linha, o que foi alterado em cada arquivo editado.

---

## 2. Fluxo atual (Solicitação de Materiais)

```
bugfix/<nome> | feature/<nome>
        |
        v   PR
    homolog      -> testes / homologação
        |
        v   (após homologado)
   Servidor     -> cópia linha por linha dos arquivos editados
```

### Passos

1. Criar a branch `bugfix/<nome-do-bug>` ou `feature/<nome-da-feature>`.
2. Desenvolver e commitar na branch.
3. Abrir PR da branch para `homolog`. Exige aprovação de Júlio e Sidney (quando autor for Rafael ou Gustavo).
4. Homologar a alteração em `homolog`.
5. Após homologado, aplicar no servidor: copiar, linha por linha, o que foi alterado em cada arquivo editado.

### Diferença para o fluxo correto

| Etapa | Fluxo correto | Fluxo atual |
|---|---|---|
| Branch bugfix/feature | Sim | Sim |
| homolog | Sim | Sim |
| development | Sim | Não |
| main | Sim | Não |
| Aprovação para produção | Sim (PR para main) | Não |
| Aplicação no servidor | Após main | Após homologado |

**Risco do fluxo atual:** o servidor recebe código que não passou por `development` nem `main`, divergindo do repositório.
