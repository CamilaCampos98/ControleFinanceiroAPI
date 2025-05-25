# Imagem base para build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

COPY *.csproj ./
RUN dotnet restore

COPY . ./
RUN dotnet publish -c Release -o out

# Imagem runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/out .

# Copia o arquivo de credenciais para dentro do container
COPY credentials.json .

# Não precisa fixar a porta aqui, Render passa via variável PORT

ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
