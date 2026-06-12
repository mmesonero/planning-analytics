/**
 * Planning & Analytics — Formulario "Aplicaciones IA" (v2)
 *
 * Crea el form y enlaza las respuestas a la MISMA hoja de cálculo
 * existente (nueva pestaña dentro del mismo Excel).
 *
 * USO:
 *  1. https://script.google.com → "Nuevo proyecto"
 *  2. Borra el código de ejemplo y pega TODO este archivo
 *  3. Click ▶ Ejecutar
 *  4. Acepta los permisos
 *  5. Ver → Registros → copia el primer URL (/viewform)
 *
 *  Si ya ejecutaste una versión anterior, ese form viejo queda huérfano
 *  en tu Drive: bórralo a mano si quieres limpiar.
 */

// ID de la hoja existente (el mismo Excel del primer form)
const SPREADSHEET_ID = '1EtrDQwfxlRcdGGo_eU4zTHYE4Ih6LDxpzJ5sr-c_bHI';

function crearFormularioAplicacionesIA() {
  const form = FormApp.create('Planning & Analytics — Aplicaciones IA');
  form.setDescription(
    'Casos de uso reales, herramientas que ya usamos y hacia dónde queremos ir en los próximos 6 meses. ' +
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

  // ── Casos de uso ────────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Casos de uso')
    .setHelpText('Hasta 3 casos concretos donde la IA puede aplicarse en tu día a día.');

  for (let i = 1; i <= 3; i++) {
    form.addParagraphTextItem()
      .setTitle('Caso de uso ' + i)
      .setHelpText(
        i === 1
          ? 'Ej: "Resumir reuniones largas con cliente y extraer próximos pasos"'
          : 'Opcional'
      );
  }

  // ── Herramientas que usamos AHORA ───────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Herramientas que usamos ahora');

  form.addParagraphTextItem()
    .setTitle('¿Qué herramientas de IA usas actualmente y para qué?')
    .setHelpText('Ej: ChatGPT para redactar correos, Copilot en Excel para fórmulas…');

  // ── Próximos 6 meses ────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Próximos 6 meses')
    .setHelpText('Hacia dónde queremos ir.');

  form.addParagraphTextItem()
    .setTitle('¿Qué queremos APRENDER a hacer con IA?')
    .setHelpText('Capacidades, técnicas o flujos que te gustaría dominar.');

  form.addParagraphTextItem()
    .setTitle('¿Qué HERRAMIENTAS queremos empezar a usar?')
    .setHelpText('Herramientas concretas que te gustaría incorporar.');

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
