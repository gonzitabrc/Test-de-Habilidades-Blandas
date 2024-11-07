function crearEncuestaHabilidadesBlandas() {
  const form = FormApp.create("Autoevaluación de Habilidades Blandas");
  form.setTitle("Autoevaluación de Habilidades Blandas")
      .setDescription("Completa esta autoevaluación para identificar tus fortalezas y áreas de mejora en habilidades blandas.");

  form.addTextItem()
      .setTitle("Nombre")
      .setRequired(true);

  form.addTextItem()
      .setTitle("Correo electrónico")
      .setRequired(true)
      .setValidation(FormApp.createTextValidation()
          .requireTextIsEmail()
          .build());

  const escala = ["1 - Nunca/Muy bajo", "2 - Rara vez/Bajo", "3 - A veces/Medio", "4 - Frecuentemente/Alto", "5 - Siempre/Muy alto"];

  agregarSeccion(form, "1. Comunicación Efectiva", [
    "Me comunico de forma clara y concisa cuando expreso mis ideas.",
    "Escucho activamente a los demás sin interrumpir y mostrando interés genuino.",
    "Evito distracciones cuando estoy en una conversación importante.",
    "Expreso mis opiniones de manera asertiva sin ser agresivo.",
    "Ajusto mi comunicación de acuerdo con el contexto y la audiencia."
  ], escala);

  agregarSeccion(form, "2. Inteligencia Emocional", [
    "Reconozco y gestiono mis emociones de forma consciente.",
    "Soy capaz de mantener la calma en situaciones de estrés o conflicto.",
    "Empatizo con los sentimientos y perspectivas de otras personas.",
    "Me esfuerzo en comprender por qué reacciono de cierta manera ante ciertas situaciones.",
    "Acepto las críticas constructivas sin tomarlo como algo personal."
  ], escala);

  agregarSeccion(form, "3. Resolución de Conflictos", [
    "Trato de identificar el problema subyacente en los conflictos antes de responder.",
    "Escucho todas las perspectivas antes de ofrecer una solución en un conflicto.",
    "Me esfuerzo en encontrar soluciones que beneficien a ambas partes involucradas.",
    "Evito los comportamientos defensivos y trato de mantener una actitud neutral.",
    "Soy capaz de negociar de manera efectiva para resolver diferencias."
  ], escala);

  agregarSeccion(form, "4. Trabajo en Equipo y Colaboración", [
    "Contribuyo activamente a los objetivos comunes del equipo.",
    "Me esfuerzo por construir relaciones de confianza y respeto con mis compañeros.",
    "Estoy dispuesto a ayudar a los demás cuando lo necesitan.",
    "Valoro las ideas y opiniones de los demás, incluso si son diferentes a las mías.",
    "Colaboro efectivamente con personas de diferentes áreas y habilidades."
  ], escala);

  agregarSeccion(form, "5. Adaptabilidad y Gestión del Cambio", [
    "Acepto los cambios en el trabajo con una actitud positiva.",
    "Soy capaz de adaptarme rápidamente a nuevas tareas o entornos.",
    "Busco activamente maneras de mejorar cuando hay un cambio en los procesos.",
    "Soy flexible cuando surgen imprevistos o cambios en los planes.",
    "Me esfuerzo por ver los cambios como oportunidades de aprendizaje."
  ], escala);

  agregarSeccion(form, "6. Gestión del Tiempo y Productividad Personal", [
    "Priorizo mis tareas de manera efectiva según su urgencia e importancia.",
    "Evito las distracciones y mantengo mi enfoque en el trabajo.",
    "Planifico mis actividades y me ajusto al tiempo disponible para cumplir con mis compromisos.",
    "Mantengo un equilibrio saludable entre el trabajo y mi vida personal.",
    "Delego tareas cuando es posible para gestionar mejor mi tiempo."
  ], escala);

  agregarSeccion(form, "7. Liderazgo y Motivación Personal", [
    "Tomo la iniciativa cuando veo una oportunidad para mejorar algo.",
    "Motivo a los demás y los apoyo para alcanzar sus objetivos.",
    "Soy responsable de mis acciones y decisiones.",
    "Me esfuerzo por ser un modelo a seguir en términos de actitud y ética laboral.",
    "Busco formas de inspirar a los demás a través de mi trabajo y conducta."
  ], escala);

  form.setConfirmationMessage("Gracias por completar la autoevaluación de habilidades blandas. Comunicate con quien te envió este test si deseas conocer el resultado");

  form.setProgressBar(true);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  Logger.log("Encuesta creada y vinculada a la hoja de cálculo: " + form.getEditUrl());
}

/**
 * Función auxiliar para agregar una sección con varias preguntas de escala al formulario
 * @param {GoogleAppsScript.Forms.Form} form - El formulario al que se agregarán las preguntas
 * @param {string} tituloSeccion - Título de la sección
 * @param {Array<string>} preguntas - Array de preguntas para la sección
 * @param {Array<string>} opciones - Opciones de escala para las preguntas
 */
function agregarSeccion(form, tituloSeccion, preguntas, opciones) {
  form.addPageBreakItem().setTitle(tituloSeccion);
  preguntas.forEach((pregunta) => {
    form.addMultipleChoiceItem()
      .setTitle(pregunta)
      .setChoiceValues(opciones)
      .setRequired(true);
  });
}


function evaluarRespuestasHabilidadesBlandas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = spreadsheet.getSheetByName("Respuestas de formulario 1");
  let resultsSheet = spreadsheet.getSheetByName("Resultados");

  if (!resultsSheet) {
    resultsSheet = spreadsheet.insertSheet("Resultados");
    resultsSheet.appendRow([
      "Marca de tiempo", "Nombre", "Correo electrónico", "Promedio General", 
      "Comunicación", "Inteligencia Emocional", "Resolución de Conflictos", 
      "Trabajo en Equipo", "Adaptabilidad", "Productividad", "Liderazgo"
    ]);
  } else if (resultsSheet.getLastRow() === 0) {
    resultsSheet.appendRow([
      "Marca de tiempo", "Nombre", "Correo electrónico", "Promedio General", 
      "Comunicación", "Inteligencia Emocional", "Resolución de Conflictos", 
      "Trabajo en Equipo", "Adaptabilidad", "Productividad", "Liderazgo"
    ]);
  }

  const respuestas = formResponsesSheet.getDataRange().getValues();

  let correosRegistrados = [];
  if (resultsSheet.getLastRow() > 1) {
    correosRegistrados = resultsSheet.getRange(2, 3, resultsSheet.getLastRow() - 1, 1)
                                     .getValues()
                                     .flat();
  }

  for (let i = 1; i < respuestas.length; i++) {
    const fila = respuestas[i];
    const timestamp = fila[0];
    const nombreParticipante = fila[1];
    const correoParticipante = fila[2];

    if (correosRegistrados.includes(correoParticipante)) {
      continue; // Salta esta entrada para evitar duplicados
    }

    const respuestasHabilidades = fila.slice(3); 
    const promedioGeneral = calcularPromedio(respuestasHabilidades);
    const promedioComunicacion = calcularPromedio(respuestasHabilidades.slice(0, 5));
    const promedioInteligencia = calcularPromedio(respuestasHabilidades.slice(5, 10));
    const promedioConflictos = calcularPromedio(respuestasHabilidades.slice(10, 15));
    const promedioEquipo = calcularPromedio(respuestasHabilidades.slice(15, 20));
    const promedioAdaptabilidad = calcularPromedio(respuestasHabilidades.slice(20, 25));
    const promedioProductividad = calcularPromedio(respuestasHabilidades.slice(25, 30));
    const promedioLiderazgo = calcularPromedio(respuestasHabilidades.slice(30, 35));
    
    resultsSheet.appendRow([
      timestamp,
      nombreParticipante,
      correoParticipante,
      promedioGeneral,
      promedioComunicacion,
      promedioInteligencia,
      promedioConflictos,
      promedioEquipo,
      promedioAdaptabilidad,
      promedioProductividad,
      promedioLiderazgo
    ]);
  }
}

function calcularPromedio(respuestas) {
  const valores = respuestas.map(respuesta => parseInt(respuesta) || 0);
  const suma = valores.reduce((acc, val) => acc + val, 0);
  return (suma / valores.length).toFixed(2);
}
