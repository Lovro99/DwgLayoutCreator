# DwgLayoutCreator

C# konzolni alat za automatsko kreiranje layouta u AutoCAD `.dwg` datotekama (via ODA/Teigha, bez AutoCAD-a).

## Projekti

### DwgLayoutCreator
Standalone alat koji na temelju definicija u kodu kreira layoute u DWG datoteci.

**Ulaz:** putanja do `.dwg` datoteke  
**Izlaz:** modificirana `.dwg` s kreiranim layoutima

### LayoutCrator / LayoutCreator
AutoCAD plugin (netload) verzija — radi unutar AutoCAD-a kao AutoLisp-dostupna komanda.

## Build

```bash
dotnet build DwgLayoutCreator/DwgLayoutCreator.csproj -c Release
```

## Ovisnosti

- [ODA File Converter / Teigha](https://www.opendesign.com/) za čitanje/pisanje DWG bez AutoCAD-a
- .NET 8, x64

## Struktura

```
DwgLayoutCreator/
├── DwgLayoutCreator/        # Standalone konzolna aplikacija
│   ├── Program.cs
│   ├── LayoutFactory.cs
│   ├── LayoutDefinitions.cs
│   ├── ModelSpaceScanner.cs
│   └── TitleBlockInserter.cs
└── LayoutCrator/
    └── LayoutCreator/       # AutoCAD plugin verzija
        ├── CreateLayoutCommand.cs
        ├── LayoutBuilder.cs
        ├── LayoutDefinitions.cs
        ├── ModelSpaceScanner.cs
        ├── PageSetupConfigurator.cs
        └── TitleBlockInserter.cs
```
