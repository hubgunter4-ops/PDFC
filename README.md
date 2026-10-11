# AUDITPRO · Revisión local de CFDI

Panel responsivo para explorar CFDI XML, complementos de pago, saldos e importes IEPS. Este rediseño se integra sobre la aplicación React/Vite y conserva el parser y las exportaciones Excel que ya existían en el repositorio.

## Ejecutar en local

```bash
npm ci
npm run dev
```

La aplicación abre en `http://localhost:3000`. Para comprobar tipos y generar la compilación de producción:

```bash
npm run lint
npm run build
```

## Funciones

- Importación múltiple de XML y ZIP con parsing en el navegador.
- Resumen, listado de CFDI, complementos de pago, saldos agrupados, desglose IEPS y observaciones.
- Búsqueda por folio, UUID, emisor y otros campos visibles; `/` enfoca el buscador.
- Ficha de lectura con importes, partes del comprobante y navegación a la factura relacionada cuando está en el lote.
- Exportación de registros a un libro Excel; el módulo de exportación se carga bajo demanda.
- Diseño adaptable con navegación inferior en móvil, estados vacíos, foco visible y movimiento reducido.

## Privacidad y alcance

El parser lee los archivos en el navegador y mantiene los registros en memoria durante la sesión. No hay backend ni llamadas al SAT implementadas en este proyecto. La herramienta **no confirma autenticidad, vigencia, estado oficial, deducibilidad ni cumplimiento fiscal**. Los saldos y etiquetas presentados reflejan campos encontrados en los XML y deben revisarse con documentación y criterio profesional.

Los archivos de prueba usados durante la verificación se mantuvieron fuera del repositorio; no se incorporan RFC, UUID, facturas, pagos ni saldos ficticios a la aplicación.

## Archivos principales

- `src/App.tsx`: navegación, importación, vistas y fichas de lectura.
- `src/index.css`: sistema visual y puntos de adaptación.
- `src/services/parser.ts`: parser de XML/ZIP y cálculos existentes.
- `src/services/excelService.ts`: generación del libro de salida.
