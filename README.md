# Horas InCicle

Script para centralizar as planilhas de horas extras da incicle.

## Instalação

Use o [pip](https://pip.pypa.io/en/stable/) para instalar as dependenciar desse script.

```bash
pip install -r requirements.txt
```

## Ultilização

- Você deve adicionar os arquivos das horas na pasta 'excels', as mesmas devem conter o nome da pessoa
separada por um dash('-') do nome do mês. Ex: Joãozinho - Janeiro

- Após concluir você deve rodar o seguinte comando na raiz do projeto:

```bash
python main.py
```

Será gerado um arquivo com o nome de 'relatorio.xlsx'