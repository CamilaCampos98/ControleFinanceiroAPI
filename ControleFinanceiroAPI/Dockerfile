# Imagem do SDK para build
FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /app

# Copia os arquivos do projeto para o container
COPY *.csproj ./
RUN dotnet restore

COPY . ./

# Publica a aplicação em modo Release para a pasta /out
RUN dotnet publish -c Release -o out

# Imagem mais leve para rodar a aplicação
FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS runtime
WORKDIR /app

# Copia a publicação da etapa build para esta imagem
COPY --from=build /app/out ./

# Expõe a porta 80 para receber requisições HTTP
EXPOSE 80

# Comando para rodar a aplicação
ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
