# Directorio de Menciones - Reportes SSFF WhatsApp

## Propósito
Archivo para mapear los números de WhatsApp y @menciones de personas que deben ser notificadas en el grupo "Canal SSFF 2026 Gestión Auren I".

---

## Supervisores
Completa con los datos de cada supervisor que deba recibir menciones:

| Nombre | Número WhatsApp | Teléfono Real | Código Ruta | Estado |
|--------|-----------------|---------------|------------|--------|
| [Supervisor 1] | 51XXXXXXXXX@c.us |  |  | Activo ☐ |
| [Supervisor 2] | 51XXXXXXXXX@c.us |  |  | Activo ☐ |
| [Supervisor N] | 51XXXXXXXXX@c.us |  |  | Activo ☐ |

---

## Jefe de Proyecto
Datos del responsable de proyecto SSFF:

| Rol | Nombre | Número WhatsApp | Teléfono Real | Email |
|-----|--------|-----------------|---------------|-------|
| Jefe de Proyecto | [Nombre] | 51XXXXXXXXX@c.us |  | |

---

## Gerente Comercial
Datos del responsable comercial:

| Rol | Nombre | Número WhatsApp | Teléfono Real | Email |
|-----|--------|-----------------|---------------|-------|
| Gerente Comercial | [Nombre] | 51XXXXXXXXX@c.us |  | |

---

## Notas
- Los números deben estar en formato WhatsApp: `51XXXXXXXXX@c.us` (sin espacios)
- Para obtener el **ID exacto del grupo**, ejecuta:
  ```bash
  python wa_client.py --list-groups
  ```
  Luego copia el ID completo en `config.json` bajo `groups.Canal_SSFF_2026_Gestion.group_id`

- Para verificar que un contacto puede ser mencionado:
  ```bash
  python wa_client.py --list-contacts "[nombre]"
  ```

- Datos sensibles: NUNCA guardes números reales sin encriptar. Este archivo debe ser `.gitignore`d una vez completado.
