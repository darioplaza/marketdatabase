# MarketDataService

Biblioteca de Google Apps Script para consultar precios de mercado de diferentes fuentes públicas.

## Características

- Consulta unificada de fondos, ETFs, acciones y criptomonedas.
- Fuentes soportadas: Investing, Quefondos, Financial Times, Yahoo Finance y Google Finance.
- Caché temporal mediante `CacheService` para reducir llamadas a la red.

## Instalación

1. Abre un documento de Google Sheets.
2. Ve a **Extensiones → Apps Script**.
3. Crea un proyecto y pega el contenido de `codigo.gs`.
4. Guarda el proyecto con un nombre, por ejemplo `MarketDataService`.

## Uso

Las funciones devuelven una única fila con el formato:

`{ Nombre ; Ticker ; Precio ; Divisa ; Fuente ; FechaISO }`

Ejemplos en una celda (configuración regional española):

```
=resolveQuote("FINANCIALTIMES"; "NVDA:NSQ")
=resolveQuoteByIsin("LU2601038735"; ; TRUE)
```

Puedes combinar la salida con `INDEX` u otras funciones de Google Sheets para obtener un campo concreto.

### resolveQuote(source, identifier)
Obtiene la cotización directamente de la fuente indicada.

### resolveQuoteByIsin(isin, [hint], [strictFunds])
Busca por ISIN aplicando un orden de *fallback*: Investing → Quefondos → FinancialTimes → (solo no fondos) Yahoo Finance → Google Finance.

## Licencia

Uso personal y educativo. Respeta los términos de servicio de cada fuente.

