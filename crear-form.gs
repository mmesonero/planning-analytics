/**
 * Planning & Analytics — Creador de formulario de objetivos y canciones
 *
 * USO (1 sola vez):
 *  1. Ve a https://script.google.com → "Nuevo proyecto"
 *  2. Borra el código de ejemplo y pega TODO este archivo
 *  3. Click ▶ Ejecutar
 *  4. Acepta los permisos con tu cuenta de Google
 *  5. Abre "Ver → Registros" — verás 3 URLs:
 *       • Form para compartir (la que pasas a los 10)
 *       • Form editor (por si quieres retocar)
 *       • Hoja de respuestas
 */

function crearFormulario() {
  const form = FormApp.create('Planning & Analytics — Objetivos y Canción');
  form.setDescription(
    'Comparte tus retos, habilidades a desarrollar y la canción que define el momento. ' +
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

  // ── Retos ────────────────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Retos a alcanzar')
    .setHelpText('Hasta 5 retos. Indica la fecha objetivo entre paréntesis.');

  for (let i = 1; i <= 5; i++) {
    form.addTextItem()
      .setTitle('Reto ' + i)
      .setHelpText(i <= 2 ? 'Ej: "Presentar Activos a internacional (Febrero)"' : 'Opcional');
  }

  // ── Habilidades ──────────────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Habilidades a desarrollar')
    .setHelpText('Hasta 5 habilidades. Indica cuándo entre paréntesis.');

  for (let i = 1; i <= 5; i++) {
    form.addTextItem()
      .setTitle('Habilidad ' + i)
      .setHelpText(i <= 2 ? 'Ej: "Story Telling (Junio)"' : 'Opcional');
  }

  // ── Canción del momento ─────────────────────────────────────
  form.addSectionHeaderItem()
    .setTitle('Canción del momento');

  form.addTextItem()
    .setTitle('Canción que te define HOY')
    .setHelpText('Nombre + artista si lo recuerdas.');

  form.addTextItem()
    .setTitle('Canción que te define en 6 MESES')
    .setHelpText('Nombre + artista si lo recuerdas.');

  // ── Crear hoja de respuestas y enlazar ──────────────────────
  const ss = SpreadsheetApp.create('Planning & Analytics — Respuestas');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // ── Mostrar URLs ────────────────────────────────────────────
  Logger.log('================================================');
  Logger.log('FORM para COMPARTIR (pásalo a los 10):');
  Logger.log(form.getPublishedUrl());
  Logger.log('');
  Logger.log('FORM editor (si quieres retocar):');
  Logger.log(form.getEditUrl());
  Logger.log('');
  Logger.log('HOJA de respuestas:');
  Logger.log(ss.getUrl());
  Logger.log('================================================');
}
