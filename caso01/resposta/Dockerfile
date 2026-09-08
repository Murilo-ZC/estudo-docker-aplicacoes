# Imagem base que será utilizada para criar o container
FROM python:3.12-slim

# Determina o diretório de trabalho dentro do container
WORKDIR /app

# Copia os arquivos do projeto para o diretório de trabalho no container
COPY . .

# Executa primeiro a instalação dos pacotes listados no requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Documenta a porta usada pelo processo. EXPOSE nao publica nada:
# publicar e trabalho do -p no docker run ou do "ports" no compose.
EXPOSE 5050

# Ainda não é um servidor de produção.
CMD ["python", "main.py"]