export function descripcionFormaPago(clave: string): string {
  const catalog: Record<string, string> = {
    "01": "01 - Efectivo",
    "02": "02 - Cheque nominativo",
    "03": "03 - Transferencia electrónica de fondos",
    "04": "04 - Tarjeta de crédito",
    "05": "05 - Monedero electrónico",
    "06": "06 - Dinero electrónico",
    "08": "08 - Vales de despensa",
    "12": "12 - Dación en pago",
    "13": "13 - Pago por subrogación",
    "14": "14 - Pago por consignación",
    "15": "15 - Condonación",
    "17": "17 - Compensación",
    "23": "23 - Novación",
    "24": "24 - Confusión",
    "25": "25 - Remisión de deuda",
    "26": "26 - Prescripción o caducidad",
    "27": "27 - A satisfacción del acreedor",
    "28": "28 - Tarjeta de débito",
    "29": "29 - Tarjeta de servicios",
    "30": "30 - Aplicación de anticipos",
    "31": "31 - Intermediario pagos",
    "99": "99 - Por definir"
  };
  return catalog[clave] || clave;
}
