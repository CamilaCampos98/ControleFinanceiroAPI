# Imagem base para build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copia o csproj e restaura depend�ncias
COPY *.csproj ./
RUN dotnet restore

# Copia todo o c�digo-fonte
COPY . ./

# Publica a aplica��o na pasta 'out'
RUN dotnet publish -c Release -o out


# Imagem runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app

# Copia o output da publica��o do build, incluindo credentials.json
COPY --from=build /app/out .

# Define o entrypoint para iniciar sua API
ENTRYPOINT ["dotnet", "ControleFinanceiroAPI.dll"]
