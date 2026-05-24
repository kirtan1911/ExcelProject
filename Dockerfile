# =========================
# Build Stage
# =========================
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build

WORKDIR /app

# Copy csproj and restore dependencies
COPY *.csproj ./
RUN dotnet restore

# Copy everything else
COPY . ./

# Publish release build
RUN dotnet publish -c Release -o /out

# =========================
# Runtime Stage
# =========================
FROM mcr.microsoft.com/dotnet/aspnet:8.0

WORKDIR /app

# Copy published output from build stage
COPY --from=build /out .

# Expose port (important for Render / Docker)
EXPOSE 8080

# Set environment port (optional but recommended)
ENV ASPNETCORE_URLS=http://+:8080

# Start application
ENTRYPOINT ["dotnet", "ExcelProject.dll"]