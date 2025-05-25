# Imagem base para build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copia csproj e restaura dependências
COPY ControleFinanceiroAPI/*.csproj ./ControleFinanceiroAPI/
WORKDIR /src/ControleFinanceiroAPI
RUN dotnet restore

# Copia tudo e publica
COPY ControleFinanceiroAPI/. .
RUN dotnet publish -c Release -o /app/publish

# Runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:10000

ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
