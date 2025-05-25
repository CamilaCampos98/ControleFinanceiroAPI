# Imagem base para build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copia o csproj e restaura dependências
COPY *.csproj ./
RUN dotnet restore

# Copia todo o código-fonte
COPY . ./

# Publica a aplicação na pasta 'out'
RUN dotnet publish -c Release -o out

# Copia o arquivo de credenciais para a pasta 'out' onde está a publicação
COPY ControleFinanceiroAPI/wwwroot/credentials.json ./out/credentials.json

# Imagem runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app

# Copia o output da publicação do build, incluindo credentials.json
COPY --from=build /app/out .

# Define o entrypoint para iniciar sua API
ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
