# PORTAFOLIO AUTÓNOMO — Varios negocios que Claude arma y opera de punta a punta

> Este documento reemplaza la idea de un solo negocio del plan anterior.
> El criterio para elegir cada negocio fue: **¿puede Claude hacer casi todo el ciclo (idea → construcción → lanzamiento → operación → mejora) sin que vos tengas que estar?**

---

## 0. Qué puedo hacer y qué no (sin vueltas)

**Puedo hacer de punta a punta:**
- Investigar el mercado y las palabras clave, elegir el nicho y validarlo con datos públicos.
- Construir sitios web, calculadoras, bots y herramientas (el código entero) y publicarlos (GitHub Pages o Cloudflare Pages).
- Crear productos digitales completos: plantillas de Excel/Sheets, eBooks en PDF/EPUB, prompt packs, guías, imágenes de presentación hechas con código.
- Escribir todo el contenido: artículos SEO, newsletters, publicaciones de venta y emails, en español e inglés.
- **Operar solo, todas las semanas**, con una **tarea programada**: una sesión mía que se ejecuta sola cada semana, escribe contenido nuevo, actualiza los sitios, revisa los números y te manda un resumen por Gmail.
- Publicar mediante API en las plataformas que lo permiten, una vez que cargás tus claves de API (se guardan cifradas y yo nunca las veo).

**No puedo (y nadie puede hacerlo por vos legalmente):**
- Crear cuentas a tu nombre, verificar tu identidad (KYC), aceptar los términos y condiciones de las plataformas ni vincular tu cuenta bancaria.
- Cobrar o mover plata.
- Subir productos donde la plataforma no tiene API (Amazon KDP, Etsy sin una app aprobada): ahí son unos **5 minutos tuyos por producto** (arrastrar el archivo y pegar el texto que yo te dejo listo).
- Publicar automáticamente en redes donde la automatización va contra las reglas (Reddit, por ejemplo) o donde la API es cara (X/Twitter). Esos canales quedan afuera.

**Tu parte total:** una configuración inicial de unas **3–4 h** (crear cuentas y cargar claves de API) + **15 min por semana** (leer el resumen y aprobar lo que haga falta) + **5 min por producto** en las plataformas sin API.

---

## 1. Ideas evaluadas (12) y cuáles quedan

Autonomía de Claude = qué porcentaje del ciclo completo puedo hacer yo una vez armadas las cuentas.

| # | Negocio | Autonomía | Ingresos a 12 meses (base) | Veredicto |
|---|---|---|---|---|
| 1 | **Red de calculadoras y herramientas financieras web** (SEO + afiliados + anuncios) | 95 % | USD 300–1.500/mes | ✅ **ENTRA** |
| 2 | **Tienda de productos digitales** (plantillas, prompt packs, kits de IA) | 85 % | USD 500–2.000/mes | ✅ **ENTRA** |
| 3 | **Newsletter automatizada semanal** (finanzas LATAM + IA) | 90 % | USD 200–1.000/mes | ✅ **ENTRA** |
| 4 | **Bot de Telegram con alertas financieras premium** (dólar, tasas, bonos, plazo fijo) | 95 % | USD 300–1.500/mes | ✅ **ENTRA** |
| 5 | **eBooks / guías en Amazon KDP** (ES + EN) | 80 % | USD 100–600/mes | ✅ Entra como complemento |
| 6 | **Sitio de nicho con reseñas y comparativas** (brokers, herramientas de IA, billeteras) + afiliados | 90 % | USD 200–1.500/mes | ✅ Se fusiona con el #1 (mismo sitio) |
| 7 | **Micro‑SaaS** (por ejemplo, analizador de deals o generador de memos con IA, por suscripción) | 85 % | USD 0–3.000/mes (muy variable) | ⏳ Fase 2, si los demás traen tráfico |
| 8 | Licencias B2B de las plantillas | 60 % | USD 0–1.000/mes | ⏳ Fase 2 (requiere vender a empresas) |
| 9 | Canal de YouTube sin cara con voz de IA | 40 % | — | ❌ Baja calidad automatizado, mucho riesgo de desmonetización |
| 10 | Print on demand (remeras, tazas) | 30 % | — | ❌ No genero diseños de imagen de calidad, mercado saturado |
| 11 | Bots que publican en Instagram/TikTok/Reddit | — | — | ❌ Van contra las reglas de las plataformas: te pueden cerrar la cuenta |
| 12 | Dropshipping / Amazon FBA | 20 % | — | ❌ Requiere capital, stock y atención al cliente |

---

## 2. El portafolio elegido: 5 negocios que se alimentan entre sí

```
   ┌──────────────────────────┐        tráfico de Google
   │ 1. CALCULADORAS + SEO    │◄────────────────────────
   │ (sitio web, ES + EN)     │
   └──────┬──────────┬────────┘
          │          │ suscripción
    venta │          ▼
          │   ┌──────────────────────┐   upsell   ┌────────────────────────┐
          │   │ 3. NEWSLETTER        │──────────► │ 4. BOT TELEGRAM PREMIUM│
          │   │ semanal automática   │            │ alertas (suscripción)  │
          │   └──────────┬───────────┘            └────────────────────────┘
          ▼              │ venta
   ┌──────────────────────────┐        ┌───────────────────────┐
   │ 2. TIENDA DIGITAL        │◄──────►│ 5. eBOOKS EN KDP      │
   │ plantillas + kits de IA  │ mismo  │ (Amazon trae tráfico) │
   └──────────────────────────┘ conten.└───────────────────────┘
```

Cada pieza de contenido se reutiliza 4 veces: **un artículo → una edición de la newsletter → un capítulo de eBook → la página de venta de una plantilla.**

---

### Negocio 1 — Red de calculadoras financieras (el motor de tráfico)

- **Qué es:** un sitio con 20–50 herramientas gratuitas que la gente busca en Google todos los días: plazo fijo vs. dólar, interés compuesto, inflación acumulada, crédito hipotecario UVA, categoría de monotributo, rendimiento de bonos y ONs, cuota de un préstamo, TIR/VAN de un proyecto inmobiliario, calculadora de FIRE. Las mismas en inglés para el público global (mortgage, DSCR, cap rate, IRR).
- **Cómo gana plata:** links de afiliados (brokers, billeteras, bancos digitales) + anuncios (AdSense/Ezoic) + venta de plantillas "pro" del Negocio 2.
- **Qué hago yo:** investigo palabras clave, programo cada calculadora, escribo el artículo de cada una, lo publico, enlazo todo entre sí, agrego 2–4 herramientas nuevas por semana con la tarea programada y actualizo los datos (tasas, inflación) desde APIs públicas.
- **Qué hacés vos (una vez):** comprar el dominio, darte de alta en Google Search Console y AdSense y anotarte en 3–5 programas de afiliados.
- **Ojo:** el SEO tarda **4–8 meses** en traer tráfico. Por eso va combinado con los negocios que venden desde el día 1 (2, 4 y 5).

### Negocio 2 — Tienda de productos digitales

- **Qué es:** un catálogo creciente de plantillas de Excel/Sheets (finanzas personales, análisis inmobiliario, memo de crédito, presupuesto familiar, cartera de inversión), prompt packs y kits de IA para profesionales, en ES y EN.
- **Dónde:** Gumroad o Lemon Squeezy (cobran por vos como vendedor registrado y te pagan a tu cuenta del exterior) + Etsy (trae compradores propios).
- **Qué hago yo:** el producto entero, las imágenes de presentación, los textos con SEO, el precio, los emails posteriores a la compra y **1 producto nuevo por semana**.
- **Qué hacés vos:** 5 minutos por producto en Etsy si no hay API (en Gumroad lo intento por API).
- **Precios:** USD 9–79 por unidad; bundles de USD 99–149.

### Negocio 3 — Newsletter semanal automatizada

- **Qué es:** "Lo que tenés que saber esta semana sobre plata, inversiones e IA", en español (LATAM) y en una edición en inglés más adelante.
- **Plataforma:** **Buttondown** (USD 0–9/mes, tiene una API completa para que yo envíe sin que vos intervengas) o Beehiiv.
- **Qué hago yo:** todos los viernes la tarea programada investiga las noticias de la semana, escribe la edición, incluye 1 herramienta del sitio y 1 producto de la tienda y la deja **programada o enviada**.
- **Qué hacés vos:** al principio aprobar cada edición (2 min). Cuando confíes en el formato, se manda sola.
- **Cómo gana plata:** sponsors desde los 2.000 suscriptores, afiliados y la venta de los Negocios 2 y 4.

### Negocio 4 — Bot de Telegram con alertas premium

- **Qué es:** un bot que avisa cuando algo pasa en el mercado: el dólar MEP/CCL cruza un valor, cambia la tasa de plazo fijo, una ON llega a cierta TIR, salen datos de inflación. Tiene un **plan gratis** (resumen diario) y un **plan premium** (alertas personalizadas y en tiempo real).
- **Cómo gana plata:** suscripción de USD 3–7/mes (Telegram Stars o Lemon Squeezy).
- **Qué hago yo:** el bot completo, la conexión con fuentes de datos públicas, el servidor (GitHub Actions o Cloudflare Workers, gratis) y el mantenimiento.
- **Qué hacés vos:** crear el bot en @BotFather (2 min) y pasarme el token mediante el formulario seguro.
- **Ojo:** es información de mercado, **no son recomendaciones de inversión**. Lleva aviso legal.

### Negocio 5 — eBooks en Amazon KDP (complemento)

- **Qué es:** guías prácticas cortas (60–120 páginas) que salen del mismo contenido: "Cómo invertir en ONs paso a paso", "Finanzas personales en Argentina con inflación", "Real estate underwriting for beginners", "AI for financial analysts".
- **Qué hago yo:** escribo el libro entero, lo maqueto (EPUB + PDF para papel), armo la tapa (con diseño tipográfico; las ilustraciones de tapa pueden necesitar Canva), la descripción y las palabras clave.
- **Qué hacés vos:** 10 min por libro en KDP (subir y marcar que el contenido fue generado con IA, que Amazon exige).
- **Ritmo:** 1 libro por mes, cada uno en ES y EN.

---

## 3. Cómo funciona solo: la operación semanal

Se configura **una tarea programada** que todos los lunes abre una sesión nueva mía y ejecuta este proceso:

| Paso | Qué hace | ¿Te necesita? |
|---|---|---|
| 1 | Revisa los números: ventas, suscriptores, tráfico | No |
| 2 | Agrega 2–4 calculadoras o artículos al sitio y lo publica | No |
| 3 | Crea 1 producto digital nuevo y lo sube a Gumroad por API (o te lo deja listo para Etsy y KDP) | Solo los 5 min de subida en Etsy/KDP |
| 4 | Escribe y programa la newsletter del viernes | Aprobación de 2 min (después, ni eso) |
| 5 | Revisa que el bot y el sitio funcionen; si algo se rompió, lo arregla | No |
| 6 | **Te manda un email de resumen**: qué hizo, cuánto se vendió, qué tenés que aprobar | 10 min de lectura |
| 7 | Cada trimestre propone qué cortar y qué duplicar | 15 min para decidir |

**Tu tiempo real:** ~15 min por semana + 5 min por cada producto en plataformas sin API.

---

## 4. Lo que tenés que hacer vos (una sola vez, ~3–4 h)

| # | Tarea | Para qué negocio | Tiempo |
|---|---|---|---|
| 1 | Elegir el nombre de la marca paraguas (o dejarme elegir) | Todos | 2 min |
| 2 | Comprar 1 dominio | 1, 3 | 10 min |
| 3 | Cuenta de Gumroad o Lemon Squeezy + vincular la cuenta del exterior | 2, 4 | 30 min |
| 4 | Cuenta de Etsy (opcional) | 2 | 30 min |
| 5 | Cuenta de Buttondown/Beehiiv | 3 | 10 min |
| 6 | Crear el bot en @BotFather | 4 | 2 min |
| 7 | Cuenta de Amazon KDP (pide los datos fiscales) | 5 | 30 min |
| 8 | Google Search Console + AdSense | 1 | 20 min |
| 9 | Anotarte en 3–5 programas de afiliados (te paso la lista) | 1, 3 | 40 min |
| 10 | Cargar las claves de API en el formulario seguro (yo no las veo) | Todos | 15 min |
| 11 | Aprobar la tarea programada semanal | Todos | 1 min |

Podés hacerlo de a poco: **cada negocio arranca apenas está su cuenta**, sin esperar a los demás.

---

## 5. Orden de lanzamiento

| Semana | Lanzo | Por qué primero |
|---|---|---|
| 1 | **Tienda digital** (5 productos) + **Bot de Telegram** (plan gratis) | Pueden vender desde el día 1 |
| 2 | **Sitio de calculadoras** (10 herramientas) | El SEO tarda; cuanto antes arranque, mejor |
| 3 | **Newsletter** (con suscripción desde el sitio y el bot) | Ya hay lugares desde donde captar suscriptores |
| 4 | **Primer eBook** en KDP + **plan premium del bot** | Ya hay contenido y usuarios del plan gratis |
| 5 en adelante | La tarea programada semanal se ocupa de todo | Vos solo leés el resumen |

---

## 6. Presupuesto (dentro de tus USD 200–1.000)

| Ítem | Costo |
|---|---|
| Dominio | USD 12/año |
| Hosting (GitHub Pages / Cloudflare Pages / GitHub Actions) | Gratis |
| Gumroad / Lemon Squeezy / Telegram | Gratis (comisión solo cuando vendés) |
| Buttondown | USD 0–9/mes |
| Etsy | USD 0,20 por publicación |
| KDP | Gratis |
| Canva Pro para tapas (opcional) | USD 13/mes |
| **Total año 1** | **~USD 50–300** |

---

## 7. Números honestos

| Mes | Conservador | Base | Optimista |
|---|---|---|---|
| 3 | USD 50 | USD 200 | USD 600 |
| 6 | USD 250 | USD 800 | USD 2.000 |
| 12 | USD 800 | USD 2.500 | USD 6.000 |
| 18 | USD 1.500 | USD 4.000 | USD 10.000+ |

- **Diversificar conviene**: si un negocio no funciona, los demás siguen. En la revisión trimestral se corta lo que no rinde.
- El escenario base supone **unos 50 productos en la tienda, más de 100 calculadoras indexadas, 2.000+ suscriptores y 100+ usuarios premium del bot hacia el mes 12**. Es exigente pero se puede, porque el trabajo lo hago yo en cada sesión semanal.

---

## 8. Riesgos y cómo cubrirlos

| Riesgo | Cobertura |
|---|---|
| Google penaliza contenido de IA de baja calidad | Herramientas útiles de verdad (calculadoras con datos reales) + artículos revisados, no relleno masivo |
| Una plataforma cambia sus reglas o cierra la cuenta | 5 negocios en plataformas distintas + lista de email propia |
| Errores en datos financieros | Fuentes oficiales (BCRA, INDEC, APIs públicas), fecha visible en cada dato, chequeos automáticos |
| Que parezca asesoramiento financiero | Todo es educativo/informativo, con aviso legal; nunca se recomienda comprar o vender un instrumento |
| Políticas sobre contenido de IA (KDP, Etsy) | Se declara siempre el uso de IA cuando la plataforma lo pide |

---

## 9. Para arrancar ya

Decime **qué negocios activás** (todos o algunos) y el **nombre de la marca** (o "elegí vos"). Con eso:
1. Construyo el primero entero en esta misma sesión.
2. Te dejo la lista de cuentas para crear, en orden y con los links.
3. Configuro la tarea programada semanal cuando esté la primera cuenta lista.
