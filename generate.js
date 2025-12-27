const { Document, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, Packer } = require('docx');

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  // Handle preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  // Only POST allowed
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  
  try {
    const data = req.body;
    console.log('📥 Generating document for:', data.nom_entreprise);
    
    // Generate document
    const doc = await generateDocument(data);
    const buffer = await Packer.toBuffer(doc);
    
    // Send file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="PAP_DUE_${sanitizeFilename(data.nom_entreprise)}.docx"`);
    res.send(buffer);
    
    console.log('✅ Document generated successfully');
    
  } catch (error) {
    console.error('❌ Error:', error);
    res.status(500).json({ error: error.message, stack: error.stack });
  }
};

// ========================================
// DOCUMENT GENERATION
// ========================================

async function generateDocument(data) {
  const paragraphs = [];
  
  // Header
  paragraphs.push(
    new Paragraph({
      text: 'PROTOCOLE D\'ACCORD PRÉÉLECTORAL',
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 }
    }),
    new Paragraph({
      text: 'RELATIF À LA MISE EN PLACE',
      alignment: AlignmentType.CENTER,
    }),
    new Paragraph({
      text: 'DU COMITÉ SOCIAL ET ÉCONOMIQUE (CSE)',
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 }
    })
  );
  
  // Entre
  paragraphs.push(
    new Paragraph({
      text: 'ENTRE :',
      bold: true,
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'La société ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: `, domiciliée au ${data.adresse || ''}, immatriculée au Registre du Commerce et des Sociétés de ${data.ville || ''} sous le numéro ${data.siret || ''} au capital de ${data.capital_social || ''} Euros, est représentée par ${data.dirigeant_nom || ''}, agissant en qualité de ${data.dirigeant_poste || ''} et ci-après dénommée la « Société ».` })
      ],
      spacing: { after: 200 }
    })
  );
  
  // Syndicats
  if (data.syndicats && data.syndicats.length > 0) {
    paragraphs.push(
      new Paragraph({
        text: 'ET :',
        bold: true,
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        text: 'Les organisations syndicales ayant répondu à l\'invitation de négocier le présent Protocole d\'Accord Préélectoral :',
        spacing: { after: 100 }
      })
    );
    
    data.syndicats.forEach(s => {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: '• ' }),
            new TextRun({ text: s.nom, bold: true }),
            new TextRun({ text: ` - ${s.complet}` })
          ]
        })
      );
    });
    
    paragraphs.push(new Paragraph({ text: '', spacing: { after: 200 } }));
  }
  
  // PRÉAMBULE
  paragraphs.push(
    new Paragraph({
      text: 'PRÉAMBULE',
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le ' }),
        new TextRun({ text: data.date_annonce_elections || '', bold: true }),
        new TextRun({ text: ', le personnel a été informé de l\'organisation des élections professionnelles par affichage dans les locaux de la société ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 150 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'La Direction a invité les Organisations Syndicales visées à l\'article L. 2314-5 du code du travail à négocier le Protocole d\'Accord Préélectoral, par lettre recommandée avec accusé de réception le ' }),
        new TextRun({ text: data.date_invitation_os || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 150 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le ' }),
        new TextRun({ text: data.date_reunion_negociation || '', bold: true }),
        new TextRun({ text: ', les organisations syndicales se sont présentées à la table des négociations.' })
      ],
      spacing: { after: 150 }
    })
  );
  
  // Recours au vote électronique
  if (data.date_signature_due) {
    paragraphs.push(
      new Paragraph({
        text: 'Recours au vote électronique :',
        bold: true,
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'La décision unilatérale signée le ' }),
          new TextRun({ text: data.date_signature_due, bold: true }),
          new TextRun({ text: ' a autorisé l\'utilisation du vote électronique pour l\'élection des membres de la délégation du personnel du CSE. L\'accompagnement ainsi que la mise à disposition de la plateforme de vote en ligne ont été confiés à la société ELECTIS.' })
        ],
        spacing: { after: 200 }
      })
    );
  }
  
  // ARTICLE 2 - DATES
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 2 - DATE DES ÉLECTIONS ET HORAIRES DU SCRUTIN',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'PREMIER TOUR',
      bold: true,
      spacing: { before: 200, after: 100 }
    })
  );
  
  if (data.date_limite_depot_candidatures_1er) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: 'Date limite de dépôt des candidatures : ' }),
          new TextRun({ text: `${data.date_limite_depot_candidatures_1er} à ${data.heure_limite_depot_candidatures_1er || ''}`, bold: true })
        ]
      })
    );
  }
  
  if (data.date_ceremonie_cles_1er) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: 'Date de scellement / Cérémonie des clés : ' }),
          new TextRun({ text: `${data.date_ceremonie_cles_1er} à ${data.heure_ceremonie_cles_1er || ''}`, bold: true })
        ]
      })
    );
  }
  
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'Date d\'ouverture du scrutin : ' }),
        new TextRun({ text: `${data.date_premier_tour || ''} à ${data.heure_ouverture_premier_tour || ''}`, bold: true })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Date de clôture du scrutin : ' }),
        new TextRun({ text: `${data.date_cloture_premier_tour || data.date_premier_tour || ''} à ${data.heure_cloture_premier_tour || ''}`, bold: true })
      ],
      spacing: { after: 200 }
    })
  );
  
  // SECOND TOUR
  paragraphs.push(
    new Paragraph({
      text: 'SECOND TOUR',
      bold: true,
      spacing: { before: 200, after: 100 }
    })
  );
  
  if (data.date_limite_depot_candidatures_2nd) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: 'Date limite de dépôt des candidatures : ' }),
          new TextRun({ text: `${data.date_limite_depot_candidatures_2nd} à ${data.heure_limite_depot_candidatures_2nd || ''}`, bold: true })
        ]
      })
    );
  }
  
  if (data.date_ceremonie_cles_2nd) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: 'Date de scellement / Cérémonie des clés : ' }),
          new TextRun({ text: `${data.date_ceremonie_cles_2nd} à ${data.heure_ceremonie_cles_2nd || ''}`, bold: true })
        ]
      })
    );
  }
  
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'Date d\'ouverture du scrutin : ' }),
        new TextRun({ text: `${data.date_second_tour || ''} à ${data.heure_ouverture_second_tour || ''}`, bold: true })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Date de clôture du scrutin : ' }),
        new TextRun({ text: `${data.date_cloture_second_tour || data.date_second_tour || ''} à ${data.heure_cloture_second_tour || ''}`, bold: true })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 3 - EFFECTIFS
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 3 - EFFECTIFS ET NOMBRE DE SIÈGES',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'effectif à la date du premier jour du 1er tour de scrutin est de ' }),
        new TextRun({ text: `${data.nombre_effectif || ''} salariés`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // Tables
  if (data.colleges && Object.keys(data.colleges).length > 0) {
    paragraphs.push(
      new Paragraph({
        text: 'Répartition par collège :',
        spacing: { before: 200, after: 100 }
      })
    );
    
    paragraphs.push(...createGenreTable(data));
    paragraphs.push(...createSiegesTable(data));
  }
  
  // ARTICLE 6
  if (data.responsable_nom_complet) {
    paragraphs.push(
      new Paragraph({
        text: 'ARTICLE 6 - DÉPÔT DES CANDIDATURES',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400, after: 200 }
      }),
      new Paragraph({
        text: 'Les listes de candidats doivent être déposées obligatoirement :',
        spacing: { after: 100 }
      })
    );
    
    if (data.moyens_depot_candidatures) {
      const moyens = data.moyens_depot_candidatures.split(',');
      moyens.forEach(moyen => {
        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({ text: '• ' }),
              new TextRun({ text: moyen.trim() })
            ]
          })
        );
      });
    }
    
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: 'Auprès de : ' }),
          new TextRun({ text: data.responsable_nom_complet, bold: true }),
          new TextRun({ text: `, ${data.responsable_poste || ''}` })
        ],
        spacing: { before: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Email : ' }),
          new TextRun({ text: data.mail_de_la_personne_en_charge_de_lorganisation_des_elections || '', bold: true })
        ]
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Bureau : ' }),
          new TextRun({ text: data.adresse_bureau || '' })
        ]
      })
    );
  }
  
  // Signatures
  paragraphs.push(
    new Paragraph({
      text: '',
      spacing: { before: 400 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Fait à ' }),
        new TextRun({ text: data.ville || '', bold: true }),
        new TextRun({ text: ', le ' }),
        new TextRun({ text: data.date_signature_pap || '', bold: true })
      ],
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Pour la Société :',
      bold: true
    }),
    new Paragraph({
      text: data.personne_nom_complet || ''
    }),
    new Paragraph({
      text: data.poste || data.dirigeant_poste || ''
    })
  );
  
  return new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: paragraphs
    }]
  });
}

// ========================================
// TABLES
// ========================================

function createGenreTable(data) {
  const rows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: 'Collège', bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: 'Hommes', bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: 'Femmes', bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: 'Total', bold: true })] }),
      ],
    })
  ];

  const collegeNames = {
    college_n0_unique: 'Collège unique',
    college_n1ouvriers_amp_employes: 'Collège n°1 : Ouvriers & Employés',
    college_n2_techniciens_amp_agents_de_matrise: 'Collège n°2 : Techniciens & Agents de maîtrise',
    college_n3__cadres_amp_assimiles: 'Collège n°3 : Cadres & Assimilés'
  };

  for (const [name, college] of Object.entries(data.colleges)) {
    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(collegeNames[name] || name)] }),
        new TableCell({ children: [new Paragraph(String(college.hommes || 0))] }),
        new TableCell({ children: [new Paragraph(String(college.femmes || 0))] }),
        new TableCell({ children: [new Paragraph(String(college.total || 0))] }),
      ],
    }));
  }

  return [
    new Paragraph({ text: '', spacing: { before: 100 } }),
    new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }),
    new Paragraph({ text: '', spacing: { after: 200 } })
  ];
}

function createSiegesTable(data) {
  const rows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: 'Collège', bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: 'Titulaires', bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: 'Suppléants', bold: true })] }),
      ],
    })
  ];

  const collegeNames = {
    college_n0_unique: 'Collège unique',
    college_n1ouvriers_amp_employes: 'Collège n°1 : Ouvriers & Employés',
    college_n2_techniciens_amp_agents_de_matrise: 'Collège n°2 : Techniciens & Agents de maîtrise',
    college_n3__cadres_amp_assimiles: 'Collège n°3 : Cadres & Assimilés'
  };

  for (const [name, college] of Object.entries(data.colleges)) {
    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(collegeNames[name] || name)] }),
        new TableCell({ children: [new Paragraph(String(college.titulaires || 0))] }),
        new TableCell({ children: [new Paragraph(String(college.suppleants || 0))] }),
      ],
    }));
  }

  return [
    new Paragraph({ text: '', spacing: { before: 100 } }),
    new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }),
    new Paragraph({ text: '', spacing: { after: 200 } })
  ];
}

// ========================================
// HELPERS
// ========================================

function sanitizeFilename(filename) {
  return (filename || 'document').replace(/[^a-z0-9]/gi, '_').toLowerCase();
}
