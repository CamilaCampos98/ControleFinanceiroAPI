# Imagem base para build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copia arquivos csproj e restaura dependências
COPY *.csproj ./
RUN dotnet restore

# Copia o restante e faz o build
COPY . ./
RUN dotnet publish -c Release -o out

# Imagem runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/out .

# Define a porta (Render detecta automaticamente, mas pode forçar se quiser)
ENV ASPNETCORE_URLS=http://+:10000

# Inicia a aplicação
ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
