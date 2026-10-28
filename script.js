/**
 * MI PLAN DE VIDA - SCRIPT DE CONTROL DE PRESENTACIÓN
 * Exactamente 6 diapositivas únicas, navegación precisa sin duplicados.
 */

document.addEventListener('DOMContentLoaded', () => {
  // Constants & Elements
  const TOTAL_SLIDES = 6;
  let currentIndex = 0;

  const slides = document.querySelectorAll('.slide-card');
  const stepBtns = document.querySelectorAll('.step-btn');
  const dotBtns = document.querySelectorAll('.dot-btn');
  const gridItems = document.querySelectorAll('.grid-item');

  const btnPrev = document.getElementById('btnPrev');
  const btnNext = document.getElementById('btnNext');
  const progressBar = document.getElementById('progressBar');
  const currentSlideNum = document.getElementById('currentSlideNum');
  const slideTitleIndicator = document.getElementById('slideTitleIndicator');

  const btnFullscreen = document.getElementById('btnFullscreen');
  const btnPrint = document.getElementById('btnPrint');
  const btnGrid = document.getElementById('btnGrid');
  const btnCloseGrid = document.getElementById('btnCloseGrid');
  const gridModal = document.getElementById('gridModal');
  const gridOverlay = document.getElementById('gridOverlay');
  const btnRestartPres = document.getElementById('btnRestartPres');
  const btnDownloadPptx = document.getElementById('btnDownloadPptx');

  // Slide Titles Mapping
  const slideTitles = [
    "Portada: Mi Plan de Vida",
    "Mis Cualidades y Personalidad: Carismático y Muy Perseverante",
    "Mis Metas y Sueños: Universidad y Maestría",
    "Mis Objetivos a Futuro: Corto, Mediano y Largo Plazo",
    "Mis Compromisos Personales: Esfuerzo Diario y Oportunidades",
    "Cierre: Convicción y Futuro"
  ];

  /**
   * Main function to navigate to a specific slide (0-5)
   * Strictly bounded to guarantee exactly 6 unique slides without loops or repeats
   */
  function goToSlide(targetIndex) {
    if (targetIndex < 0 || targetIndex >= TOTAL_SLIDES) return;

    // Previous slide animation state
    slides.forEach((slide, idx) => {
      slide.classList.remove('active', 'prev-slide');
      if (idx < targetIndex) {
        slide.classList.add('prev-slide');
      }
    });

    currentIndex = targetIndex;
    slides[currentIndex].classList.add('active');

    // Update Counter & Title
    currentSlideNum.textContent = (currentIndex + 1);
    slideTitleIndicator.textContent = slideTitles[currentIndex];

    // Update Progress Bar
    const progressPercent = ((currentIndex + 1) / TOTAL_SLIDES) * 100;
    progressBar.style.width = `${progressPercent}%`;

    // Update Navigation Buttons States (Strict bounds: Prev disabled on 0, Next disabled on 5)
    btnPrev.disabled = (currentIndex === 0);
    btnNext.disabled = (currentIndex === TOTAL_SLIDES - 1);

    if (currentIndex === TOTAL_SLIDES - 1) {
      btnNext.classList.remove('btn-primary');
    } else {
      btnNext.classList.add('btn-primary');
    }

    // Update Step buttons in header
    stepBtns.forEach((btn, idx) => {
      btn.classList.remove('active', 'completed');
      if (idx === currentIndex) {
        btn.classList.add('active');
      } else if (idx < currentIndex) {
        btn.classList.add('completed');
      }
    });

    // Update Footer Dots
    dotBtns.forEach((dot, idx) => {
      dot.classList.toggle('active', idx === currentIndex);
    });

    // Update Grid modal items
    gridItems.forEach((item, idx) => {
      item.classList.toggle('active', idx === currentIndex);
    });
  }

  // Previous & Next Handlers
  function prevSlide() {
    if (currentIndex > 0) {
      goToSlide(currentIndex - 1);
    }
  }

  function nextSlide() {
    if (currentIndex < TOTAL_SLIDES - 1) {
      goToSlide(currentIndex + 1);
    }
  }

  // Button Event Listeners
  btnPrev.addEventListener('click', prevSlide);
  btnNext.addEventListener('click', nextSlide);

  if (btnRestartPres) {
    btnRestartPres.addEventListener('click', () => goToSlide(0));
  }

  // Step Header Clicks
  stepBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.getAttribute('data-index'), 10);
      goToSlide(idx);
    });
  });

  // Dot Clicks
  dotBtns.forEach(dot => {
    dot.addEventListener('click', () => {
      const idx = parseInt(dot.getAttribute('data-index'), 10);
      goToSlide(idx);
    });
  });

  // Grid items click
  gridItems.forEach(item => {
    item.addEventListener('click', () => {
      const idx = parseInt(item.getAttribute('data-index'), 10);
      goToSlide(idx);
      closeGridModal();
    });
  });

  // Keyboard Navigation
  document.addEventListener('keydown', (e) => {
    // If modal is open, Escape closes it
    if (gridModal.classList.contains('open')) {
      if (e.key === 'Escape') closeGridModal();
      return;
    }

    switch (e.key) {
      case 'ArrowRight':
      case 'PageDown':
      case ' ': // Spacebar
        e.preventDefault();
        nextSlide();
        break;
      case 'ArrowLeft':
      case 'PageUp':
        e.preventDefault();
        prevSlide();
        break;
      case 'Home':
        e.preventDefault();
        goToSlide(0);
        break;
      case 'End':
        e.preventDefault();
        goToSlide(TOTAL_SLIDES - 1);
        break;
      case 'f':
      case 'F':
        toggleFullscreen();
        break;
      case 'Escape':
        openGridModal();
        break;
    }
  });

  // Fullscreen Handler
  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => {
        console.warn(`Fullscreen error: ${err.message}`);
      });
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      }
    }
  }

  btnFullscreen.addEventListener('click', toggleFullscreen);

  // Print / Save to PDF
  btnPrint.addEventListener('click', () => {
    window.print();
  });

  // Grid Modal Controls
  function openGridModal() {
    gridModal.classList.add('open');
  }

  function closeGridModal() {
    gridModal.classList.remove('open');
  }

  btnGrid.addEventListener('click', openGridModal);
  btnCloseGrid.addEventListener('click', closeGridModal);
  gridOverlay.addEventListener('click', closeGridModal);

  // PowerPoint (.pptx) Export Handler using PptxGenJS
  btnDownloadPptx.addEventListener('click', exportToPowerPoint);

  function exportToPowerPoint() {
    if (typeof PptxGenJS === 'undefined') {
      alert("Generador PPTX no disponible en este momento. Puedes usar la versión descargada en la carpeta.");
      return;
    }

    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';
      pptx.author = 'Manuel Ayala';
      pptx.title = 'Mi Plan de Vida';

      // Common theme colors
      const BG_COLOR = "0B0F19";
      const CARD_BG = "141B2D";
      const TEXT_WHITE = "FFFFFF";
      const TEXT_MUTED = "94A3B8";
      const ACCENT_INDIGO = "6366F1";
      const ACCENT_CYAN = "06B6D4";
      const ACCENT_AMBER = "F59E0B";
      const ACCENT_EMERALD = "10B981";

      // -------------------------------------------------------------
      // SLIDE 1: PORTADA
      // -------------------------------------------------------------
      const s1 = pptx.addSlide();
      s1.background = { color: BG_COLOR };

      s1.addShape(pptx.ShapeType.rect, {
        x: 0.8, y: 0.8, w: 3.5, h: 0.4,
        fill: { color: "1E293B" }, line: { color: ACCENT_CYAN, width: 1 }
      });
      s1.addText("PROYECTO DE VIDA Y VISIÓN", {
        x: 0.8, y: 0.8, w: 3.5, h: 0.4,
        fontSize: 10, bold: true, color: ACCENT_CYAN, align: "center", fontFace: "Calibri"
      });

      s1.addText("Mi Plan de Vida", {
        x: 0.8, y: 1.6, w: 8.4, h: 1.5,
        fontSize: 48, bold: true, color: TEXT_WHITE, fontFace: "Arial"
      });

      s1.addText("Un mapa estratégico forjado con convicción, metas claras y la determinación constante de construir un futuro de excelencia y trascendencia.", {
        x: 0.8, y: 3.2, w: 8.4, h: 1.0,
        fontSize: 16, color: TEXT_MUTED, fontFace: "Calibri"
      });

      // 3 boxes
      const boxes = [
        { title: "PROPÓSITO", desc: "Superación Académica y Profesional", icon: "🎯", color: ACCENT_INDIGO },
        { title: "MOTOR PRINCIPAL", desc: "Carisma, Disciplina y Perseverancia", icon: "⚡", color: ACCENT_AMBER },
        { title: "HORIZONTE", desc: "Bachillerato ➔ Universidad ➔ Maestría", icon: "🎓", color: ACCENT_CYAN }
      ];

      boxes.forEach((b, i) => {
        const xPos = 0.8 + (i * 3.7);
        s1.addShape(pptx.ShapeType.roundRect, {
          x: xPos, y: 4.6, w: 3.5, h: 1.6,
          fill: { color: CARD_BG }, line: { color: b.color, width: 1 }, rectRadius: 0.1
        });
        s1.addText(b.title, {
          x: xPos + 0.3, y: 4.8, w: 3.0, h: 0.3,
          fontSize: 11, bold: true, color: b.color, fontFace: "Calibri"
        });
        s1.addText(b.desc, {
          x: xPos + 0.3, y: 5.2, w: 3.0, h: 0.8,
          fontSize: 13, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
        });
      });

      // -------------------------------------------------------------
      // SLIDE 2: CUALIDADES Y PERSONALIDAD
      // -------------------------------------------------------------
      const s2 = pptx.addSlide();
      s2.background = { color: BG_COLOR };

      s2.addText("IDENTIDAD & FORTALEZA", {
        x: 0.8, y: 0.6, w: 6.0, h: 0.3,
        fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
      });
      s2.addText("Mis Cualidades y Personalidad", {
        x: 0.8, y: 0.9, w: 10.0, h: 0.8,
        fontSize: 32, bold: true, color: TEXT_WHITE, fontFace: "Arial"
      });

      // Card 1: Carismático
      s2.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.4, h: 4.6,
        fill: { color: CARD_BG }, line: { color: ACCENT_AMBER, width: 1.5 }, rectRadius: 0.1
      });
      s2.addText("🌟 Soy Carismático", {
        x: 1.2, y: 2.1, w: 4.6, h: 0.6,
        fontSize: 22, bold: true, color: ACCENT_AMBER, fontFace: "Arial"
      });
      s2.addText("Habilidad natural para conectar positivamente con las personas mediante empatía y comunicación asertiva.", {
        x: 1.2, y: 2.8, w: 4.6, h: 0.8,
        fontSize: 13, color: TEXT_WHITE, fontFace: "Calibri"
      });
      s2.addText([
        { text: "• Liderazgo empático: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Inspirar entusiasmo en equipos.\n\n", options: { color: TEXT_MUTED } },
        { text: "• Comunicación asertiva: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Relaciones sólidas y constructivas.\n\n", options: { color: TEXT_MUTED } },
        { text: "• Actitud positiva: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Generación de confianza y dinamismo.", options: { color: TEXT_MUTED } }
      ], { x: 1.2, y: 3.7, w: 4.6, h: 2.3, fontSize: 12, fontFace: "Calibri" });

      // Card 2: Muy Perseverante
      s2.addShape(pptx.ShapeType.roundRect, {
        x: 6.6, y: 1.8, w: 5.4, h: 4.6,
        fill: { color: CARD_BG }, line: { color: ACCENT_INDIGO, width: 1.5 }, rectRadius: 0.1
      });
      s2.addText("🛡️ Muy Perseverante", {
        x: 7.0, y: 2.1, w: 4.6, h: 0.6,
        fontSize: 22, bold: true, color: ACCENT_INDIGO, fontFace: "Arial"
      });
      s2.addText("Determinación inquebrantable que no se rinde ante los retos, transformando obstáculos en superación constante.", {
        x: 7.0, y: 2.8, w: 4.6, h: 0.8,
        fontSize: 13, color: TEXT_WHITE, fontFace: "Calibri"
      });
      s2.addText([
        { text: "• Constancia diaria: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Disciplina activa día con día.\n\n", options: { color: TEXT_MUTED } },
        { text: "• Resiliencia activa: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Firmeza para levantarse ante tropiezos.\n\n", options: { color: TEXT_MUTED } },
        { text: "• Foco incondicional: ", options: { bold: true, color: TEXT_WHITE } },
        { text: "Compromiso total hasta culminar el objetivo.", options: { color: TEXT_MUTED } }
      ], { x: 7.0, y: 3.7, w: 4.6, h: 2.3, fontSize: 12, fontFace: "Calibri" });

      // -------------------------------------------------------------
      // SLIDE 3: METAS Y SUEÑOS
      // -------------------------------------------------------------
      const s3 = pptx.addSlide();
      s3.background = { color: BG_COLOR };

      s3.addText("ASPIRACIÓN & VOCACIÓN", {
        x: 0.8, y: 0.6, w: 6.0, h: 0.3,
        fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
      });
      s3.addText("Mis Metas y Sueños", {
        x: 0.8, y: 0.9, w: 10.0, h: 0.8,
        fontSize: 32, bold: true, color: TEXT_WHITE, fontFace: "Arial"
      });

      // Goal 1
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
      s3.addText("Alcanzar el cupo y comenzar formalmente mi formación académica profesional en el ámbito universitario, adquiriendo competencias avanzadas.", {
        x: 1.2, y: 3.3, w: 4.5, h: 1.5,
        fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
      });
      s3.addText("• Dedicación académica al 100%\n• Vocación técnica y pasión por la ingeniería", {
        x: 1.2, y: 4.8, w: 4.5, h: 1.0,
        fontSize: 12, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
      });

      // Goal 2
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
      s3.addText("Dar el salto hacia la excelencia posgradual para profundizar mis capacidades investigativas, de gestión y liderazgo técnico a nivel nacional e internacional.", {
        x: 7.1, y: 3.3, w: 4.5, h: 1.5,
        fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
      });
      s3.addText("• Liderazgo e innovación técnica\n• Alta competitividad profesional", {
        x: 7.1, y: 4.8, w: 4.5, h: 1.0,
        fontSize: 12, bold: true, color: TEXT_WHITE, fontFace: "Calibri"
      });

      // -------------------------------------------------------------
      // SLIDE 4: OBJETIVOS A FUTURO
      // -------------------------------------------------------------
      const s4 = pptx.addSlide();
      s4.background = { color: BG_COLOR };

      s4.addText("TRAYECTORIA ESTRATÉGICA", {
        x: 0.8, y: 0.6, w: 6.0, h: 0.3,
        fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
      });
      s4.addText("Mis Objetivos a Futuro", {
        x: 0.8, y: 0.9, w: 10.0, h: 0.8,
        fontSize: 32, bold: true, color: TEXT_WHITE, fontFace: "Arial"
      });

      const timelines = [
        {
          horizon: "CORTO PLAZO",
          goal: "Graduarme de bachillerato",
          desc: "Culminar la etapa de educación media superior con solidez académica y el certificado que abra las puertas universitarias.",
          color: ACCENT_CYAN
        },
        {
          horizon: "MEDIANO PLAZO",
          goal: "Ingresar a la universidad y elegir una ingeniería o un técnico en mantenimiento",
          desc: "Superar el proceso de admisión y dar el paso hacia la carrera técnica o ingeniería elegida.",
          color: ACCENT_INDIGO
        },
        {
          horizon: "LARGO PLAZO",
          goal: "Graduarme como ingeniero",
          desc: "Obtener el título universitario de ingeniería, consolidando mis competencias para liderar en la industria.",
          color: ACCENT_EMERALD
        }
      ];

      timelines.forEach((t, i) => {
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
          x: xPos + 0.3, y: 3.9, w: 2.9, h: 1.8,
          fontSize: 12, color: TEXT_MUTED, fontFace: "Calibri"
        });
      });

      // -------------------------------------------------------------
      // SLIDE 5: COMPROMISOS PERSONALES
      // -------------------------------------------------------------
      const s5 = pptx.addSlide();
      s5.background = { color: BG_COLOR };

      s5.addText("PACTO Y DISCIPLINA", {
        x: 0.8, y: 0.6, w: 6.0, h: 0.3,
        fontSize: 11, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
      });
      s5.addText("Mis Compromisos Personales", {
        x: 0.8, y: 0.9, w: 10.0, h: 0.8,
        fontSize: 32, bold: true, color: TEXT_WHITE, fontFace: "Arial"
      });

      // Commit 1
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
      s5.addText("La perseverancia es un compromiso diario. Me comprometo a dedicar esfuerzo continuo, constancia y disciplina personal en cada jornada formativa, sin ceder ante las distracciones ni el desánimo.", {
        x: 1.2, y: 3.6, w: 4.6, h: 1.6,
        fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
      });
      s5.addText("✓ Hábitos de estudio constantes\n✓ Disciplina y autogestión", {
        x: 1.2, y: 5.2, w: 4.6, h: 0.8,
        fontSize: 12, bold: true, color: ACCENT_CYAN, fontFace: "Calibri"
      });

      // Commit 2
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
      s5.addText("Mantener siempre una mente proactiva y alerta para tomar cualquier iniciativa académica, laboral o de desarrollo personal que sume a mi crecimiento. Estar preparado cuando llegue la oportunidad.", {
        x: 7.0, y: 3.6, w: 4.6, h: 1.6,
        fontSize: 13, color: TEXT_MUTED, fontFace: "Calibri"
      });
      s5.addText("✓ Mentalidad proactiva y abierta\n✓ Iniciativa y aprendizaje continuo", {
        x: 7.0, y: 5.2, w: 4.6, h: 0.8,
        fontSize: 12, bold: true, color: ACCENT_AMBER, fontFace: "Calibri"
      });

      // -------------------------------------------------------------
      // SLIDE 6: CIERRE
      // -------------------------------------------------------------
      const s6 = pptx.addSlide();
      s6.background = { color: BG_COLOR };

      s6.addText("CONCLUSIÓN Y HORIZONTE", {
        x: 0.8, y: 1.0, w: 11.2, h: 0.4,
        fontSize: 12, bold: true, color: ACCENT_AMBER, align: "center", fontFace: "Calibri"
      });
      s6.addText("Convicción y Futuro", {
        x: 0.8, y: 1.5, w: 11.2, h: 1.0,
        fontSize: 44, bold: true, color: TEXT_WHITE, align: "center", fontFace: "Arial"
      });

      s6.addShape(pptx.ShapeType.rect, {
        x: 1.5, y: 2.7, w: 9.8, h: 1.8,
        fill: { color: CARD_BG }, line: { color: ACCENT_INDIGO, width: 1.5 }
      });
      s6.addText('"El futuro no se espera, se construye. Con la certeza de mis cualidades, el anhelo de mis metas y la firmeza de mis compromisos diarios, avanzo con paso seguro hacia mi meta: ser un profesional de excelencia que trascienda y deje huella."', {
        x: 1.8, y: 2.9, w: 9.2, h: 1.4,
        fontSize: 15, italic: true, color: TEXT_WHITE, align: "center", fontFace: "Calibri"
      });

      const pillars = [
        { title: "CONVICCIÓN", desc: "Seguridad en mis capacidades", icon: "✦" },
        { title: "PERSEVERANCIA", desc: "Constancia inquebrantable", icon: "◈" },
        { title: "FUTURO", desc: "Ingeniero y Maestro", icon: "▲" }
      ];

      pillars.forEach((p, i) => {
        const xPos = 1.5 + (i * 3.4);
        s6.addShape(pptx.ShapeType.roundRect, {
          x: xPos, y: 4.8, w: 3.0, h: 1.4,
          fill: { color: "111827" }, line: { color: "374151", width: 1 }, rectRadius: 0.1
        });
        s6.addText(`${p.icon}  ${p.title}`, {
          x: xPos + 0.2, y: 5.0, w: 2.6, h: 0.35,
          fontSize: 12, bold: true, color: ACCENT_CYAN, align: "center", fontFace: "Calibri"
        });
        s6.addText(p.desc, {
          x: xPos + 0.2, y: 5.4, w: 2.6, h: 0.5,
          fontSize: 11, color: TEXT_MUTED, align: "center", fontFace: "Calibri"
        });
      });

      // Save presentation
      pptx.writeFile({ fileName: "Mi_Plan_de_Vida.pptx" });
    } catch (err) {
      console.error("Error al exportar PPTX:", err);
      alert("Error al generar PPTX: " + err.message);
    }
  }

  // Initialize on first slide
  goToSlide(0);
});
