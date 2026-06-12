/**
 * Planning & Analytics — Formulario "Aplicaciones IA"
 *
 * Crea un segundo formulario y lo enlaza a la MISMA hoja de cálculo
 * que ya tienes (las respuestas van a una pestaña nueva dentro del mismo Excel).
 *
 * USO (1 sola vez):
 *  1. https://script.google.com → "Nuevo proyecto"
 *  2. Borra el código de ejemplo y pega TODO este archivo
 *  3. Click ▶ Ejecutar
 *  4. Acepta los permisos con tu cuenta de Google
 *  5. Abre "Ver → Registros" — verás 2 URLs:
 *       • Form para compartir (la que pasas a los 10)
 *       • Form editor (por si quieres retocar)
 *     Las respuestas aparecen en una pestaña nueva del mismo spreadsheet.
 */

// ID de la hoja existente (el mismo Excel del primer form)
const SPREADSHEET_ID = '1EtrDQwfxlRcdGGo_eU4zTHYE4Ih6LDxpzJ5sr-c_bHI';

function crearFormularioAplicacionesIA() {
  const form = FormApp.create('Planning & Analytics — Aplicaciones IA');
  form.setDescription(
    'Ideas y casos de uso reales donde la IA puede ayudarnos en nuestro día a día. ' +
    'Si eres nueva incorporación, selecciona "Nueva incorporación" y escribe tu nombre.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);

  // ── ¿Quién eres? ─────────────────────────────────────────────
  form.addListItem()
    .setTitle('¿Quién eres?')
    .setRequired(true)
    .setChoiceValues([
      'Luis',
      'Alejandra',
      'Blanca',
      'Claudia',
      'Manu',
      'Adrian',
      'Jesús',
      'Alba',
      'Pablo',
      'Alvaro',
      '— Nueva incorporación —'
    ]);

  form.addTextItem()
    .setTitle('Si eres nueva incorporación, escribe tu nombre')
    .setHelpText('Déjalo vacío si ya estabas en la lista anterior.');

  // ── Áreas o procesos ────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Áreas o procesos donde aplicar IA')
    .setHelpText('Hasta 3 áreas de tu día a día donde la IA podría ayudarte.');

  for (let i = 1; i <= 3; i++) {
    form.addTextItem()
      .setTitle('Área / proceso ' + i)
      .setHelpText(i === 1 ? 'Ej: "Preparación de propuestas comerciales"' : 'Opcional');
  }

  // ── Casos de uso concretos ──────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Casos de uso concretos')
    .setHelpText('Tareas específicas donde la IA te ahorraría tiempo o mejoraría el resultado.');

  for (let i = 1; i <= 3; i++) {
    form.addParagraphTextItem()
      .setTitle('Caso de uso ' + i)
      .setHelpText(i === 1 ? 'Ej: "Resumir reuniones largas con cliente y extraer próximos pasos"' : 'Opcional');
  }

  // ── Tareas repetitivas a automatizar ────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Tareas repetitivas que podríamos automatizar');

  form.addParagraphTextItem()
    .setTitle('¿Qué tarea repetitiva te quita más tiempo cada semana?')
    .setHelpText('Cuanto más concreta, mejor.');

  // ── Herramientas ────────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Herramientas');

  form.addTextItem()
    .setTitle('Herramientas de IA que YA usas')
    .setHelpText('Ej: ChatGPT, Claude, Copilot, Gemini, Notion AI…');

  form.addTextItem()
    .setTitle('Herramientas de IA que te gustaría PROBAR')
    .setHelpText('Lo que has visto que usan otros y quieres explorar.');

  // ── Comentarios ─────────────────────────────────────────────
  form.addParagraphTextItem()
    .setTitle('Comentarios o ideas adicionales')
    .setHelpText('Opcional. Cualquier cosa relacionada que quieras añadir.');

  // ── Enlazar al spreadsheet existente (nueva pestaña auto) ──
  form.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);

  // ── Mostrar URLs ────────────────────────────────────────────
  Logger.log('================================================');
  Logger.log('FORM para COMPARTIR (pásalo a los 10):');
  Logger.log(form.getPublishedUrl());
  Logger.log('');
  Logger.log('FORM editor (si quieres retocar):');
  Logger.log(form.getEditUrl());
  Logger.log('');
  Logger.log('Las respuestas aparecen en una NUEVA pestaña del mismo Excel.');
  Logger.log('================================================');
}
