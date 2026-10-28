const pptxgen = require('pptxgenjs');

async function createPresentation() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'Manuel Ayala';
  pptx.title = 'Mi Plan de Vida';

  const BG_COLOR = "0A0E1A";
  const CARD_BG = "131A2E";
  const TEXT_WHITE = "FFFFFF";
  const TEXT_MUTED = "94A3B8";
  const ACCENT_INDIGO = "6366F1";
  const ACCENT_CYAN = "06B6D4";
  const ACCENT_AMBER = "F59E0B";
  const ACCENT_EMERALD = "10B981";

  // ==========================================================
  // DIAPOSITIVA 1: PORTADA
  // ==========================================================
  const s1 = pptx.addSlide();
  s1.background = { color: BG_COLOR };

  s1.addShape(pptx.ShapeType.roundRect, {
    x: 0.8, y: 0.8, w: 3.5, h: 0.38,
    fill: { color: "1E293B" }, line: { color: ACCENT_CYAN, width: 1 }, rectRadius: 0.1
  });
  s1.addText("PROYECTO DE VIDA Y VISIÓN", {
    x: 0.8, y: 0.8, w: 3.5, h: 0.38,
    fontSize: 10, bold: true, color: ACCENT_CYAN, align: "center", fontFace: "Calibri"
  });

  s1.addText("Mi Plan de Vida", {
    x: 0.8, y: 1.5, w: 10.0, h: 1.4,
    fontSize: 50, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });

  s1.addText("Un mapa estratégico forjado con convicción, metas claras y la determinación constante de construir un futuro de excelencia y trascendencia.", {
    x: 0.8, y: 3.0, w: 9.5, h: 1.0,
    fontSize: 16, color: TEXT_MUTED, fontFace: "Calibri"
  });

  const s1Boxes = [
    { title: "PROPÓSITO", desc: "Superación Académica y Profesional", color: ACCENT_INDIGO },
    { title: "MOTOR PRINCIPAL", desc: "Carisma, Disciplina y Perseverancia", color: ACCENT_AMBER },
    { title: "HORIZONTE", desc: "Bachillerato ➔ Universidad ➔ Maestría", color: ACCENT_CYAN }
  ];

  s1Boxes.forEach((b, i) => {
    const xPos = 0.8 + (i * 3.75);
    s1.addShape(pptx.ShapeType.roundRect, {
      x: xPos, y: 4.5, w: 3.5, h: 1.8,
      fill: { color: CARD_BG }, line: { color: b.color, width: 1.5 }, rectRadius: 0.1
    });
    s1.addText(b.title, {
      x: xPos + 0.3, y: 4.75, w: 3.0, h: 0.3,
      fontSize: 11, bold: true, color: b.color, fontFace: "Calibri"
    });
    s1.addText(b.desc, {
      x: xPos + 0.3, y: 5.2, w: 3.0, h: 0.8,
      fontSize: 13, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
    });
  });

  // ==========================================================
  // DIAPOSITIVA 2: MIS CUALIDADES Y PERSONALIDAD
  // ==========================================================
  const s2 = pptx.addSlide();
  s2.background = { color: BG_COLOR };

  s2.addText("IDENTIDAD & FORTALEZA", {
    x: 0.8, y: 0.6, w: 6.0, h: 0.3,
    fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s2.addText("Mis Cualidades y Personalidad", {
    x: 0.8, y: 0.9, w: 10.0, h: 0.8,
    fontSize: 34, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });

  // Card 1: Carismático
  s2.addShape(pptx.ShapeType.roundRect, {
    x: 0.8, y: 1.85, w: 5.4, h: 4.6,
    fill: { color: CARD_BG }, line: { color: ACCENT_AMBER, width: 1.5 }, rectRadius: 0.1
  });
  s2.addText("Soy Carismático", {
    x: 1.2, y: 2.1, w: 4.6, h: 0.6,
    fontSize: 24, bold: true, color: ACCENT_AMBER, fontFace: "Arial"
  });
  s2.addText("Habilidad natural para conectar positivamente con las personas mediante empatía, energía constructiva y comunicación asertiva.", {
    x: 1.2, y: 2.8, w: 4.6, h: 0.8,
    fontSize: 13, color: TEXT_WHITE, fontFace: "Calibri"
  });
  s2.addText([
    { text: "• Liderazgo empático: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Inspirar entusiasmo en proyectos de equipo.\n\n", options: { color: TEXT_MUTED } },
    { text: "• Comunicación asertiva: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Construcción de relaciones sólidas y enriquecedoras.\n\n", options: { color: TEXT_MUTED } },
    { text: "• Actitud positiva: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Apertura que crea ambientes de confianza y dinamismo.", options: { color: TEXT_MUTED } }
  ], { x: 1.2, y: 3.7, w: 4.6, h: 2.3, fontSize: 12, fontFace: "Calibri" });

  // Card 2: Muy Perseverante
  s2.addShape(pptx.ShapeType.roundRect, {
    x: 6.6, y: 1.85, w: 5.4, h: 4.6,
    fill: { color: CARD_BG }, line: { color: ACCENT_INDIGO, width: 1.5 }, rectRadius: 0.1
  });
  s2.addText("Muy Perseverante", {
    x: 7.0, y: 2.1, w: 4.6, h: 0.6,
    fontSize: 24, bold: true, color: ACCENT_INDIGO, fontFace: "Arial"
  });
  s2.addText("Determinación inquebrantable que no se rinde ante los retos, convirtiendo cualquier obstáculo en superación y aprendizaje.", {
    x: 7.0, y: 2.8, w: 4.6, h: 0.8,
    fontSize: 13, color: TEXT_WHITE, fontFace: "Calibri"
  });
  s2.addText([
    { text: "• Constancia diaria: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Trabajo continuo y disciplina aún bajo exigencia.\n\n", options: { color: TEXT_MUTED } },
    { text: "• Resiliencia activa: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Firmeza para levantarse con más fuerza ante tropiezos.\n\n", options: { color: TEXT_MUTED } },
    { text: "• Foco incondicional: ", options: { bold: true, color: TEXT_WHITE } },
    { text: "Compromiso absoluto hasta culminar el objetivo propuesto.", options: { color: TEXT_MUTED } }
  ], { x: 7.0, y: 3.7, w: 4.6, h: 2.3, fontSize: 12, fontFace: "Calibri" });

  // ==========================================================
  // DIAPOSITIVA 3: MIS METAS Y SUEÑOS
  // ==========================================================
  const s3 = pptx.addSlide();
  s3.background = { color: BG_COLOR };

  s3.addText("ASPIRACIÓN & VOCACIÓN", {
    x: 0.8, y: 0.6, w: 6.0, h: 0.3,
    fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s3.addText("Mis Metas y Sueños", {
    x: 0.8, y: 0.9, w: 10.0, h: 0.8,
    fontSize: 34, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });

  // Goal 1: Ingresar a la Universidad
  s3.addShape(pptx.ShapeType.roundRect, {
    x: 0.8, y: 1.9, w: 5.3, h: 4.5,
    fill: { color: CARD_BG }, line: { color: ACCENT_CYAN, width: 1.5 }, rectRadius: 0.1
  });
  s3.addText("ETAPA 1: EDUCACIÓN SUPERIOR", {
    x: 1.2, y: 2.2, w: 4.5, h: 0.3,
    fontSize: 10, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s3.addText("Ingresar a la Universidad", {
    x: 1.2, y: 2.5, w: 4.5, h: 0.7,
    fontSize: 22, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });
  s3.addText("Acceder a la educación superior como la puerta principal hacia el pensamiento crítico, la ciencia y la formación tecnológica de excelencia.", {
    x: 1.2, y: 3.3, w: 4.5, h: 1.4,
    fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
  });
  s3.addText("✓ Dedicación académica al 100%\n✓ Vocación técnica y pasión por la ingeniería", {
    x: 1.2, y: 4.8, w: 4.5, h: 1.0,
    fontSize: 12, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
  });

  // Goal 2: Sacar una Maestría
  s3.addShape(pptx.ShapeType.roundRect, {
    x: 6.7, y: 1.9, w: 5.3, h: 4.5,
    fill: { color: CARD_BG }, line: { color: "A855F7", width: 1.5 }, rectRadius: 0.1
  });
  s3.addText("ETAPA 2: ALTA ESPECIALIZACIÓN", {
    x: 7.1, y: 2.2, w: 4.5, h: 0.3,
    fontSize: 10, bold: true, color: "A855F7", fontFace: "Calibri"
  });
  s3.addText("Poder Sacar una Maestría", {
    x: 7.1, y: 2.5, w: 4.5, h: 0.7,
    fontSize: 22, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });
  s3.addText("Consolidar mis estudios superiores mediante un posgrado de maestría que me posicione como un especialista de referencia, capaz de liderar soluciones de alto impacto.", {
    x: 7.1, y: 3.3, w: 4.5, h: 1.4,
    fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
  });
  s3.addText("✓ Liderazgo e investigación aplicada\n✓ Proyección y competitividad internacional", {
    x: 7.1, y: 4.8, w: 4.5, h: 1.0,
    fontSize: 12, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
  });

  // ==========================================================
  // DIAPOSITIVA 4: MIS OBJETIVOS A FUTURO
  // ==========================================================
  const s4 = pptx.addSlide();
  s4.background = { color: BG_COLOR };

  s4.addText("HOJA DE RUTA EN 3 FASES", {
    x: 0.8, y: 0.6, w: 6.0, h: 0.3,
    fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s4.addText("Mis Objetivos a Futuro", {
    x: 0.8, y: 0.9, w: 10.0, h: 0.8,
    fontSize: 34, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });

  const timelineCards = [
    {
      horizon: "CORTO PLAZO",
      goal: "Graduarme de bachillerato",
      desc: "Culminar la educación media superior con solidez académica, disciplina y la preparación necesaria para afrontar el examen de admisión.",
      color: ACCENT_CYAN,
      actions: "• Cierre académico exitoso\n• Preparación de ingreso"
    },
    {
      horizon: "MEDIANO PLAZO",
      goal: "Ingresar a la universidad y elegir una ingeniería o un técnico en mantenimiento",
      desc: "Obtener mi lugar en la educación superior y consolidar mi camino vocacional entre una ingeniería o un técnico en mantenimiento.",
      color: ACCENT_INDIGO,
      actions: "• Admisión universitaria lograda\n• Elección vocacional técnica"
    },
    {
      horizon: "LARGO PLAZO",
      goal: "Graduarme como ingeniero",
      desc: "Completar la totalidad del plan de estudios y titularme formalmente como Ingeniero, listo para liderar proyectos en la industria.",
      color: ACCENT_EMERALD,
      actions: "• Título profesional universitario\n• Integración al campo laboral"
    }
  ];

  timelineCards.forEach((t, i) => {
    const xPos = 0.8 + (i * 3.75);
    s4.addShape(pptx.ShapeType.roundRect, {
      x: xPos, y: 1.9, w: 3.5, h: 4.6,
      fill: { color: CARD_BG }, line: { color: t.color, width: 1.5 }, rectRadius: 0.1
    });
    s4.addText(t.horizon, {
      x: xPos + 0.3, y: 2.2, w: 2.9, h: 0.3,
      fontSize: 10, bold: true, color: t.color, fontFace: "Calibri"
    });
    s4.addText(t.goal, {
      x: xPos + 0.3, y: 2.6, w: 2.9, h: 1.2,
      fontSize: 15, bold: true, color: TEXT_WHITE, fontFace: "Arial"
    });
    s4.addText(t.desc, {
      x: xPos + 0.3, y: 3.9, w: 2.9, h: 1.5,
      fontSize: 12, color: TEXT_MUTED, fontFace: "Calibri"
    });
    s4.addText(t.actions, {
      x: xPos + 0.3, y: 5.4, w: 2.9, h: 0.8,
      fontSize: 11, bold: true, color: t.color, fontFace: "Calibri"
    });
  });

  // ==========================================================
  // DIAPOSITIVA 5: MIS COMPROMISOS PERSONALES
  // ==========================================================
  const s5 = pptx.addSlide();
  s5.background = { color: BG_COLOR };

  s5.addText("PACTO Y DISCIPLINA DIARIA", {
    x: 0.8, y: 0.6, w: 6.0, h: 0.3,
    fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s5.addText("Mis Compromisos Personales", {
    x: 0.8, y: 0.9, w: 10.0, h: 0.8,
    fontSize: 34, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });

  // Commitment 1
  s5.addShape(pptx.ShapeType.roundRect, {
    x: 0.8, y: 1.9, w: 5.4, h: 4.5,
    fill: { color: CARD_BG }, line: { color: ACCENT_INDIGO, width: 1.5 }, rectRadius: 0.1
  });
  s5.addText("COMPROMISO #1", {
    x: 1.2, y: 2.2, w: 4.6, h: 0.3,
    fontSize: 10, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });
  s5.addText('"Esforzarme cada día para alcanzar mis metas"', {
    x: 1.2, y: 2.6, w: 4.6, h: 0.9,
    fontSize: 18, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });
  s5.addText("La perseverancia se demuestra en el día a día. Me comprometo a dedicar esfuerzo sincero, enfoque y hábitos disciplinados de estudio en cada jornada, sin rendirme ante la pereza ni las dificultades momentáneas.", {
    x: 1.2, y: 3.6, w: 4.6, h: 1.5,
    fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
  });
  s5.addText("✓ Hábitos de estudio diarios y constantes\n✓ Autodisciplina y enfoque en resultados", {
    x: 1.2, y: 5.2, w: 4.6, h: 0.8,
    fontSize: 12, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
  });

  // Commitment 2
  s5.addShape(pptx.ShapeType.roundRect, {
    x: 6.6, y: 1.9, w: 5.4, h: 4.5,
    fill: { color: CARD_BG }, line: { color: ACCENT_AMBER, width: 1.5 }, rectRadius: 0.1
  });
  s5.addText("COMPROMISO #2", {
    x: 7.0, y: 2.2, w: 4.6, h: 0.3,
    fontSize: 10, bold: true, color: ACCENT_AMBER, fontFace: "Calibri"
  });
  s5.addText('"Aprovechar las oportunidades que se me presenten"', {
    x: 7.0, y: 2.6, w: 4.6, h: 0.9,
    fontSize: 18, bold: true, color: TEXT_WHITE, fontFace: "Arial"
  });
  s5.addText("Mantener una mentalidad proactiva y despierta para identificar cursos, becas, proyectos, prácticas y relaciones constructivas que impulsen mi avance vocacional y profesional.", {
    x: 7.0, y: 3.6, w: 4.6, h: 1.5,
    fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
  });
  s5.addText("✓ Mentalidad abierta al aprendizaje continuo\n✓ Iniciativa y audacia constructiva", {
    x: 7.0, y: 5.2, w: 4.6, h: 0.8,
    fontSize: 12, bold: true, color: ACCENT_AMBER, fontFace: "Calibri"
  });

  // ==========================================================
  // DIAPOSITIVA 6: CIERRE
  // ==========================================================
  const s6 = pptx.addSlide();
  s6.background = { color: BG_COLOR };

  s6.addText("CONCLUSIÓN Y HORIZONTE", {
    x: 0.8, y: 1.0, w: 11.2, h: 0.4,
    fontSize: 12, bold: true, color: ACCENT_AMBER, align: "center", fontFace: "Calibri"
  });
  s6.addText("Convicción y Futuro", {
    x: 0.8, y: 1.45, w: 11.2, h: 1.0,
    fontSize: 44, bold: true, color: TEXT_WHITE, align: "center", fontFace: "Arial"
  });

  s6.addShape(pptx.ShapeType.rect, {
    x: 1.5, y: 2.65, w: 9.8, h: 1.8,
    fill: { color: CARD_BG }, line: { color: ACCENT_INDIGO, width: 1.5 }
  });
  s6.addText('"El futuro no se espera, se construye. Con la certeza de mis cualidades, el anhelo de mis metas y la firmeza de mis compromisos diarios, avanzo con paso seguro hacia mi meta: ser un profesional de excelencia que trascienda y deje huella."', {
    x: 1.8, y: 2.85, w: 9.2, h: 1.4,
    fontSize: 15, italic: true, color: TEXT_WHITE, align: "center", fontFace: "Calibri"
  });

  const finalPillars = [
    { title: "CONVICCIÓN", desc: "Seguridad y fe en mis capacidades", color: ACCENT_CYAN },
    { title: "PERSEVERANCIA", desc: "Constancia inquebrantable diaria", color: ACCENT_AMBER },
    { title: "FUTURO", desc: "Ingeniero y Maestro de excelencia", color: ACCENT_EMERALD }
  ];

  finalPillars.forEach((p, i) => {
    const xPos = 1.5 + (i * 3.4);
    s6.addShape(pptx.ShapeType.roundRect, {
      x: xPos, y: 4.8, w: 3.0, h: 1.4,
      fill: { color: "111827" }, line: { color: p.color, width: 1 }, rectRadius: 0.1
    });
    s6.addText(p.title, {
      x: xPos + 0.2, y: 5.0, w: 2.6, h: 0.35,
      fontSize: 12, bold: true, color: p.color, align: "center", fontFace: "Calibri"
    });
    s6.addText(p.desc, {
      x: xPos + 0.2, y: 5.4, w: 2.6, h: 0.5,
      fontSize: 11, color: TEXT_MUTED, align: "center", fontFace: "Calibri"
    });
  });

  const outputPath = 'Mi_Plan_de_Vida.pptx';
  await pptx.writeFile({ fileName: outputPath });
  console.log(`Presentación PowerPoint creada con éxito en: ${outputPath}`);
}

createPresentation().catch(err => {
  console.error("Error al generar la presentación:", err);
  process.exit(1);
});
