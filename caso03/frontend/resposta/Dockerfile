# O frontend e HTML estatico. Nao existe processo Node aqui: o processo do
# container e o proprio nginx, que serve os arquivos e repassa /api/ ao
# backend. Container existe para processo em execucao, e arquivo estatico
# nao e processo.
FROM nginx:1.27-alpine

# Substitui a configuracao default do site. O caminho conf.d/default.conf e
# incluido pelo nginx.conf principal da imagem oficial.
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Diretorio onde o "root" do nginx.conf aponta.
COPY index.html /usr/share/nginx/html/index.html

EXPOSE 80

# A imagem oficial do nginx ja define o CMD correto
# (nginx -g "daemon off;"), rodando em primeiro plano para o container nao
# encerrar.