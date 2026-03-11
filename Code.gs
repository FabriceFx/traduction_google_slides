// ============================================================
//  SLIDES TRANSLATOR — Google Apps Script  |  Code.gs
// ============================================================

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SlidesApp.getUi()
    .createMenu('🌐 Traducteur')
    .addItem('Ouvrir le traducteur', 'openSidebar')
    .addSeparator()
    .addItem('À propos', 'showAbout')
    .addToUi();
}

function openSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('🌐 Traducteur de slides')
      .setWidth(320)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SlidesApp.getUi().showSidebar(html);
  } catch (e) {
    console.error("Erreur lors de l'ouverture de la sidebar : " + e.toString());
  }
}

function showAbout() {
  SlidesApp.getUi().alert(
    'Traduction Google Slides',
    'Traduit tous les textes de votre présentation.\n\nUtilise le service de traduction natif de Google.\n\n© Fabrice Faucheux',
    SlidesApp.getUi().ButtonSet.OK
  );
}

// ─── TRADUCTION PRINCIPALE ───────────────────────────────

function translatePresentation(options) {
  let step = "Démarrage";
  
  try {
    step = "Récupération de la présentation active";
    const pres = SlidesApp.getActivePresentation();
    let targetPres = pres;

    if (options.createCopy) {
      step = "Traduction du titre de la présentation";
      const originalTitle = pres.getName();
      const src = options.sourceLang === 'auto' ? '' : options.sourceLang;
      let translatedTitle = originalTitle;
      
      try {
        // On traduit le titre de la présentation
        translatedTitle = LanguageApp.translate(originalTitle, src, options.targetLang);
      } catch (err) {
        console.warn("Impossible de traduire le titre : " + err.message);
      }

      step = "Création de la copie (API Drive Avancée)";
      const copyName = translatedTitle + ' [' + getLanguageLabel(options.targetLang) + ']';
      
      const copiedFile = Drive.Files.copy(
        { name: copyName }, 
        pres.getId(), 
        { supportsAllDrives: true }
      );
      
      step = "Pause de synchronisation (Google Drive)";
      Utilities.sleep(2500); 
      
      step = "Ouverture de la copie (SlidesApp)";
      targetPres = SlidesApp.openById(copiedFile.id);
    }

    step = "Récupération des slides";
    const allSlides = targetPres.getSlides();
    const slideIndices = parseSlideRange(options.slideRange, allSlides.length);

    for (let i = 0; i < slideIndices.length; i++) {
      const idx = slideIndices[i];
      step = "Traduction de la slide " + (idx + 1);
      translateSlide(allSlides[idx], options.sourceLang, options.targetLang, options.translateNotes);
    }

    if (options.createCopy) {
      step = "Sauvegarde du nouveau fichier";
      targetPres.saveAndClose();
    }

    step = "Génération du lien URL";
    const url = options.createCopy ? DriveApp.getFileById(targetPres.getId()).getUrl() : null;

    return {
      success: true,
      message: slideIndices.length + ' slide(s) traduite(s) avec succès !',
      url: url,
    };
  } catch (e) {
    return { success: false, message: 'Crash à l\'étape [' + step + '] : ' + e.message };
  }
}

function translateSlide(slide, sourceLang, targetLang, translateNotes) {
  for (const shape of slide.getShapes()) {
    if (shape.getText) translateTextRange(shape.getText(), sourceLang, targetLang);
  }
  for (const table of slide.getTables()) {
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getNumColumns(); c++) {
        translateTextRange(table.getCell(r, c).getText(), sourceLang, targetLang);
      }
    }
  }
  if (translateNotes) {
    const notes = slide.getSpeakerNotesShape();
    if (notes) translateTextRange(notes.getText(), sourceLang, targetLang);
  }
}

function translateTextRange(textRange, sourceLang, targetLang) {
  // 1. Sécurité : si la zone de texte n'existe pas, on s'arrête
  if (!textRange) return; 
  
  // 2. On récupère directement les segments de texte (on ignore les paragraphes !)
  const runs = textRange.getRuns();
  if (!runs) return;

  for (const run of runs) {
    const original = run.asString();
    
    // 3. On ignore les segments vides ou qui ne sont que des espaces/sauts de ligne
    if (!original.trim()) continue;
    
    try {
      const src = sourceLang === 'auto' ? '' : sourceLang;
      // Traduction
      const translatedText = LanguageApp.translate(original, src, targetLang);
      // Remplacement du texte
      run.setText(translatedText);
    } catch (e) {
      console.warn('Impossible de traduire : ' + original + ' | Erreur: ' + e.message);
    }
  }
}

// ─── UTILITAIRES ─────────────────────────────────────────

function getPresentationInfo() {
  try {
    const pres = SlidesApp.getActivePresentation();
    return { name: pres.getName(), slideCount: pres.getSlides().length };
  } catch (e) {
    return { name: 'Inconnu', slideCount: 0 };
  }
}

function parseSlideRange(range, total) {
  if (!range || range === 'all') return Array.from({ length: total }, (_, i) => i);
  const indices = new Set();
  for (const part of range.split(',')) {
    const t = part.trim();
    if (t.includes('-')) {
      const [s, e] = t.split('-').map(Number);
      for (let i = s; i <= e; i++) if (i >= 1 && i <= total) indices.add(i - 1);
    } else {
      const n = parseInt(t);
      if (!isNaN(n) && n >= 1 && n <= total) indices.add(n - 1);
    }
  }
  return Array.from(indices).sort((a, b) => a - b);
}

function getLanguageLabel(code) {
  const lang = getLanguages().find(l => l.code === code);
  return lang ? lang.label : code;
}

// ─── LISTE DES LANGUES ────────────────────────────────────

function getLanguages() {
  return [
    { code: 'auto', label: '🔍 Détection automatique' },
    { code: 'af', label: 'Afrikaans' },
    { code: 'sq', label: 'Albanais' },
    { code: 'de', label: 'Allemand' },
    { code: 'en', label: 'Anglais' },
    { code: 'ar', label: 'Arabe' },
    { code: 'hy', label: 'Arménien' },
    { code: 'az', label: 'Azerbaïdjanais' },
    { code: 'eu', label: 'Basque' },
    { code: 'bn', label: 'Bengali' },
    { code: 'bg', label: 'Bulgare' },
    { code: 'ca', label: 'Catalan' },
    { code: 'zh', label: 'Chinois (simplifié)' },
    { code: 'zh-TW', label: 'Chinois (traditionnel)' },
    { code: 'ko', label: 'Coréen' },
    { code: 'hr', label: 'Croate' },
    { code: 'da', label: 'Danois' },
    { code: 'es', label: 'Espagnol' },
    { code: 'eo', label: 'Espéranto' },
    { code: 'et', label: 'Estonien' },
    { code: 'fi', label: 'Finnois' },
    { code: 'fr', label: 'Français' },
    { code: 'gl', label: 'Galicien' },
    { code: 'ka', label: 'Géorgien' },
    { code: 'el', label: 'Grec' },
    { code: 'gu', label: 'Gujarati' },
    { code: 'he', label: 'Hébreu' },
    { code: 'hi', label: 'Hindi' },
    { code: 'hu', label: 'Hongrois' },
    { code: 'id', label: 'Indonésien' },
    { code: 'ga', label: 'Irlandais' },
    { code: 'is', label: 'Islandais' },
    { code: 'it', label: 'Italien' },
    { code: 'ja', label: 'Japonais' },
    { code: 'kn', label: 'Kannada' },
    { code: 'lv', label: 'Letton' },
    { code: 'lt', label: 'Lituanien' },
    { code: 'mk', label: 'Macédonien' },
    { code: 'ms', label: 'Malais' },
    { code: 'ml', label: 'Malayalam' },
    { code: 'mt', label: 'Maltais' },
    { code: 'mr', label: 'Marathi' },
    { code: 'nl', label: 'Néerlandais' },
    { code: 'ne', label: 'Népalais' },
    { code: 'nb', label: 'Norvégien' },
    { code: 'fa', label: 'Persan' },
    { code: 'pl', label: 'Polonais' },
    { code: 'pt', label: 'Portugais' },
    { code: 'pa', label: 'Punjabi' },
    { code: 'ro', label: 'Roumain' },
    { code: 'ru', label: 'Russe' },
    { code: 'sr', label: 'Serbe' },
    { code: 'sk', label: 'Slovaque' },
    { code: 'sl', label: 'Slovène' },
    { code: 'sv', label: 'Suédois' },
    { code: 'sw', label: 'Swahili' },
    { code: 'ta', label: 'Tamoul' },
    { code: 'te', label: 'Télougou' },
    { code: 'th', label: 'Thaï' },
    { code: 'cs', label: 'Tchèque' },
    { code: 'tr', label: 'Turc' },
    { code: 'uk', label: 'Ukrainien' },
    { code: 'ur', label: 'Ourdou' },
    { code: 'vi', label: 'Vietnamien' }
  ].sort((a, b) => {
    if (a.code === 'auto') return -1;
    return a.label.localeCompare(b.label, 'fr');
  });
}
