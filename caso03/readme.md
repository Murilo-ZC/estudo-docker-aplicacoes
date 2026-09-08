# Cria uma rede gerenciada pelo docker
docker network create calc-net

# Lança o container do backend
docker run -d --rm --name backend --network calc-net calculadora:1.0

# Lança o container do frontend
docker run -d --rm --name frontend --network calc-net -p 8080:80 calc-front:1.0