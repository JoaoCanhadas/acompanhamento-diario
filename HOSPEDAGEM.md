# Hospedagem do dashboard

Este projeto esta pronto para hospedar como um servico web Python.

## Opcao recomendada: Render

1. Crie uma conta em https://render.com.
2. No Render, clique em **New > Web Service**.
3. Conecte o repositorio `JoaoCanhadas/acompanhamento-diario`.
4. Use estas configuracoes:
   - **Runtime:** Python
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python dashboard.py`
5. Escolha o plano **Free**.
6. Publique o servico.

O Render define a porta automaticamente pela variavel `PORT`, e o arquivo `dashboard.py` ja esta preparado para isso.

## Atualizando os dados

O site usa o arquivo `data.json`. Quando os dados mudarem, gere uma nova versao desse arquivo e envie para o GitHub.
