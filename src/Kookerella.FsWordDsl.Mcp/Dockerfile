# Build from the repo root as context, since this project references its sibling
# Kookerella.FsWordDsl and Kookerella.CsWordDsl libraries via plain ProjectReferences
# (the latter needed only for generate_csharp_script - see the fsproj's own comment):
#   docker build -f src/Kookerella.FsWordDsl.Mcp/Dockerfile -t fsworddsl-mcp .
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY src/Kookerella.FsWordDsl/Kookerella.FsWordDsl.fsproj src/Kookerella.FsWordDsl/
COPY src/Kookerella.CsWordDsl/Kookerella.CsWordDsl.csproj src/Kookerella.CsWordDsl/
COPY src/Kookerella.FsWordDsl.Mcp/Kookerella.FsWordDsl.Mcp.fsproj src/Kookerella.FsWordDsl.Mcp/
RUN dotnet restore src/Kookerella.FsWordDsl.Mcp/Kookerella.FsWordDsl.Mcp.fsproj
COPY src/Kookerella.FsWordDsl/ src/Kookerella.FsWordDsl/
COPY src/Kookerella.CsWordDsl/ src/Kookerella.CsWordDsl/
COPY src/Kookerella.FsWordDsl.Mcp/ src/Kookerella.FsWordDsl.Mcp/
RUN dotnet publish src/Kookerella.FsWordDsl.Mcp/Kookerella.FsWordDsl.Mcp.fsproj -c Release -o /app --no-restore

FROM mcr.microsoft.com/dotnet/runtime:10.0 AS runtime
WORKDIR /app
COPY --from=build /app .
# MCP over stdio - no exposed port. The client launches this and talks JSON-RPC over
# stdin/stdout (see Program.fs's own comment on why logging is routed to stderr instead).
ENTRYPOINT ["dotnet", "Kookerella.FsWordDsl.Mcp.dll"]
