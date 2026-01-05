const { Document, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, Packer, BorderStyle } = require('docx');

module.exports = async (req, res) => {
  // CORS headers - CRÍTICO para HubSpot
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');
   
  if (req.method === 'OPTIONS') {
    console.log('✅ CORS preflight request handled');
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    console.log('❌ Invalid method:', req.method);
    return res.status(405).json({ error: 'Method not allowed' });
  }
  
  try {
    const data = req.body;
    console.log('📥 Generating COMPLETE document for:', data.nom_entreprise);
    
    const doc = await generateCompleteDocument(data);
    const buffer = await Packer.toBuffer(doc);
    
    console.log('✅ Complete document generated successfully');
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="PAP_DUE_${sanitizeFilename(data.nom_entreprise)}.docx"`);
    res.send(buffer);
    
  } catch (error) {
    console.error('❌ Error:', error);
    res.status(500).json({ error: error.message, stack: error.stack });
  }
};

// ========================================
// COMPLETE DOCUMENT GENERATION
// ========================================

async function generateCompleteDocument(data) {
  const paragraphs = [];
  
  // HEADER
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '–﻿Toutes les mentions entre tirets, comme dans cet exemple : –exemple–, sont des indications d\'aide et doivent être supprimées.–',
                    italics: true
        })
      ],
      spacing: { before: 0, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'PROTOCOLE D\'ACCORD PRÉÉLECTORAL', font: 'Roboto',
          bold: true
        })
      ],
      heading: HeadingLevel.HEADING_1,
      bold: true,
      spacing: { before: 0, after: 120 }
    }),
    new Paragraph({ children: [new TextRun({ text: 'RELATIF À LA MISE EN PLACE', font: 'Roboto' })], spacing: { after: 60 }
    }),
    new Paragraph({ children: [new TextRun({ text: 'DU COMITÉ SOCIAL ET ÉCONOMIQUE (CSE)', font: 'Roboto' })], spacing: { after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '–OU (garder la mention qui correspond à votre situation)–', font: 'Roboto',
          italics: true
        })
      ],
      spacing: { after: 120 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'DÉCISION UNILATÉRALE DE L\'EMPLOYEUR (DUE)', font: 'Roboto',
          bold: true
        })
      ],
      heading: HeadingLevel.HEADING_1,
      bold: true,
      spacing: { after: 120 }
    }),
    new Paragraph({ children: [new TextRun({ text: 'RELATIVE À LA MISE EN PLACE', font: 'Roboto' })], spacing: { after: 60 }
    }),
    new Paragraph({ children: [new TextRun({ text: 'DU COMITÉ SOCIAL ET ÉCONOMIQUE (CSE)', font: 'Roboto' })], spacing: { after: 400 }
    })
  );
  
  // ENTRE
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: '–(supprimer si c\'est une DUE garder que le paragraphe ci-dessous)–', font: 'Roboto',
          italics: true
        })
      ],
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'ENTRE :', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'La société ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: `, domiciliée au ${data.adresse || ''}, immatriculée au Registre du Commerce et des Sociétés de ${data.ville || ''} sous le numéro ${data.siret || ''} au capital de ${data.capital_social || ''} Euros, est représentée par ${data.personne_nom_complet || data.dirigeant_nom || ''}, agissant en qualité de ${data.dirigeant_poste || ''} et ci-après dénommée la « Société ».` })
      ],
      spacing: { after: 200 }
    })
  );
  
  // SYNDICATS
  if (data.syndicats && data.syndicats.length > 0) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: '–(supprimer si c\'est une DUE)–', font: 'Roboto',
            italics: true
          })
        ],
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'ET :', font: 'Roboto',
            bold: true
          })
        ],
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        children: [
          new TextRun({ font: 'Roboto', text: 'Les organisations syndicales ayant répondu à l\'invitation de négocier le présent Protocole d\'Accord Préélectoral, ayant dûment mandaté :',
            font: 'Roboto'
          })
        ],
        spacing: { after: 100 }
      })
    );
    
    data.syndicats.forEach(s => {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ font: 'Roboto', text: '• ' }),
            new TextRun({ font: 'Roboto', text: s.nom, bold: true }),
            new TextRun({ font: 'Roboto', text: ` - ${s.complet}` })
          ]
        })
      );
      
      // CORRECCIÓN 3: Agregar información del representante sindical
      if (s.representante && s.representante.trim()) {
        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({ font: 'Roboto', text: `   ${s.representante} en sa qualité de représentant(e) syndical(e)` })
            ],
            spacing: { left: 360 }
          })
        );
      }
    });
    
    paragraphs.push(new Paragraph({ text: '', spacing: { after: 200 } }));
  }
  
  // PRÉAMBULE
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'PRÉAMBULE', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_1,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le ' }),
        new TextRun({ font: 'Roboto', text: data.date_annonce_elections || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ', le personnel a été informé de l\'organisation des élections professionnelles par affichage dans les locaux de la société ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Dans le cadre de la mise en place du Comité Social et Economique (CSE), la Direction a invité les Organisations Syndicales visées à l\'article L. 2314-5 du code du travail à négocier le Protocole d\'Accord Préélectoral, par lettre recommandée avec accusé de réception ' }),
        new TextRun({ font: 'Roboto', text: data.date_invitation_os || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les autres Organisations syndicales intéressées ont été informées de l\'organisation des élections et invitées à négocier le protocole d\'accord préélectoral par voie d\'affichage dans les locaux de la société ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ' le ' }),
        new TextRun({ font: 'Roboto', text: data.date_invitation_os || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
        new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '—Si DUE, garder le paragraphe ci-dessous et rajouter la négation.—',
                    italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le ' }),
        new TextRun({ font: 'Roboto', text: data.date_reunion_negociation || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ', les organisations syndicales précitées se sont présentées à la table des négociations en vue de négocier, notamment la répartition du personnel et des sièges dans les différents collèges électoraux ainsi que les modalités d\'organisation et de déroulement des élections professionnelles.' })
      ],
      spacing: { after: 150 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Recours au vote électronique :', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'article 54 de la loi n° 2004-575 du 21 juin 2004 pour la confiance dans l\'économie numérique autorise le recours au vote électronique pour les élections professionnelles.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les articles R. 2314-5 à R. 2314-18 du code du travail définissent les modalités de vote par voie électronique pour l\'élection des membres de la délégation du personnel du CSE.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'La décision unilatérale signée le ' }),
        new TextRun({ font: 'Roboto', text: data.date_signature_due || data.date_signature_pap || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ' a autorisé l\'utilisation du vote électronique pour l\'élection des membres de la délégation du personnel du CSE. L\'accompagnement ainsi que la mise à disposition de la plateforme de vote en ligne ont été confiés à la société ELECTIS, pour les élections régies par le présent accord.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 1
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 1 - OBJET ET CHAMP D\'APPLICATION', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le présent accord a pour objet de définir les modalités d\'organisation de l\'élection des représentants du personnel au CSE de la société ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '–Si DUE changer « le présent accord » par la décision–', font: 'Roboto',
          italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément aux dispositions légales, le présent accord a pour champ d\'application le CSE de l\'entreprise ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 2
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 2 - DATE DES ÉLECTIONS ET HORAIRES DU SCRUTIN', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Pour le premier tour, la date d\'ouverture du scrutin est fixée pour l\'ensemble des collèges le ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_premier_tour || ''} à ${data.heure_ouverture_premier_tour || ''}`, bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Au premier tour, sont habilitées à présenter leur liste de candidats, les organisations syndicales représentatives dans l\'entreprise, les syndicats affiliés à une organisation reconnue représentative aux niveaux national et interprofessionnel, et tout syndicat qui satisfait aux critères de respect des valeurs républicaines, d\'indépendance, légalement constitué depuis au moins 2 ans, et dont le champ professionnel et géographique couvre l\'entreprise concernée.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Il y aura lieu de procéder à un second tour, ouvert aux candidats libres et aux candidatures syndicales dans les conditions définies pour le premier tour, dans l\'un des cas suivants :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Quorum non atteint au premier tour : moins de la moitié des électeurs inscrits ont émis un vote valable.', font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Carence : absence de candidature syndicale au premier tour.', font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Non attribution de siège : les sièges n\'ont pas tous été pourvus dès le premier tour.', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les cas ci-dessus s\'apprécient par collège et par scrutin.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le cas échéant, la date d\'ouverture du scrutin du deuxième tour est fixée le ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_second_tour || ''} à ${data.heure_ouverture_second_tour || ''}`, bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les collèges concernés par cette élection sont composés de :', font: 'Roboto'
        })
      ],
      spacing: { before: 100, after: 50 }
    })
  );
  
  // Collèges
  const collegeLabels = {
    'college_n0_unique': 'Collège unique',
    'college_n1ouvriers_amp_employes': 'Collège n°1 : Ouvriers & Employés',
    'college_n2_techniciens_amp_agents_de_matrise': 'Collège n°2 : Techniciens & Agents de maîtrise',
    'college_n3__cadres_amp_assimiles': 'Collège n°3 : Cadres & Assimilés'
  };
  
  if (data.colleges) {
    for (const [key, college] of Object.entries(data.colleges)) {
      if (collegeLabels[key]) {
        paragraphs.push(
          new Paragraph({
            text: `• ${collegeLabels[key]}`,
            spacing: { left: 360 }
          })
        );
      }
    }
  }
  
  paragraphs.push(new Paragraph({ text: '', spacing: { after: 200 } }));
  
  // ARTICLE 3
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 3 - EFFECTIFS ET NOMBRE DE SIÈGES', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'effectif à la date du premier jour du 1er tour de scrutin est calculé selon les règles prévues aux articles L. 1111-2 et L. 1111-3 du code du travail.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Plus précisément, y sont inclus :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les CDI à temps plein,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les CDD (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les salariés mis à la disposition de l\'entreprise par une entreprise extérieure, qui sont présents dans les locaux de l\'entreprise utilisatrice et y travaillent depuis au moins un an (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les salariés temporaires sont pris en compte dans l\'effectif de l\'entreprise proportionnellement à leur temps de présence au cours des douze mois précédents (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• les salariés à temps partiel en fonction de leur durée du travail.', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'En sont exclus :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les apprentis,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les stagiaires,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les titulaires d\'un contrat de professionnalisation, jusqu\'au terme prévu par le contrat lorsque celui-ci est à durée déterminée, ou, jusqu\'à la fin de l\'action de professionnalisation lorsque le contrat est à durée indéterminée.',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'L\'effectif (ETP : équivalent temps plein) de référence servant à déterminer le nombre de sièges est de ' }),
        new TextRun({ font: 'Roboto', text: `${data.nombre_effectif || ''}`, bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Ci-après le détail :', font: 'Roboto'
        })
      ],
      spacing: { before: 100, after: 100 }
    })
  );
  
  // Tables
  if (data.colleges && Object.keys(data.colleges).length > 0) {
    paragraphs.push(...createGenreTable(data));
    
    // Texte entre les tables
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ font: 'Roboto', text: 'Conformément aux dispositions légales en vigueur, le nombre de Représentant du personnel à élire au CSE est de :',
            font: 'Roboto'
          })
        ],
        spacing: { before: 100, after: 100 }
      })
    );
    
    paragraphs.push(...createSiegesTable(data));
  }
  
  // ARTICLE 4
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 4 - CONDITIONS POUR ÊTRE ÉLECTEUR', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément aux dispositions de l\'article L. 2314-18 du code du travail, sont électeurs les salariés de l\'entreprise qui remplissent les conditions suivantes à la date du premier tour du scrutin :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• être titulaire d\'un contrat de travail,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• être âgé de 16 ans révolus,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• n\'avoir fait l\'objet d\'aucune interdiction, déchéance ou incapacité relatives à leurs droits civiques,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• travailler depuis 3 mois au moins dans l\'entreprise à la date du premier tour des élections.', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Par application de l\'article L. 2314-23 du code du travail, pour les salariés mis à disposition qui remplissent les conditions mentionnées au 2° de l\'article L.1111-2 du code du travail, la condition de présence dans l\'entreprise utilisatrice est de douze mois continus pour y être électeur.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '—Dans le paragraphe ci-dessous, enlever la négation si c\'est le cas et préciser quelle(s) entreprise(s) et combien de personnes.—',
                    italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'A cet effet, la Direction constate qu\'il n\'y a pas d\'entreprises prestataires ayant du personnel mis à disposition de la société. Il n\'y a donc aucune entreprise prestataire à contacter afin qu\'elle fournisse la liste des salariés mis à disposition répondant aux critères de présence dans les locaux et d\'ancienneté et qui souhaiteraient voter au sein de la société.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 5
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 5 - CONDITIONS POUR ÊTRE ÉLIGIBLE', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément aux dispositions de l\'article L.2314-19 du code du travail, pour être éligible, un salarié doit répondre aux conditions suivantes pour chaque tour de scrutin :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• être électeur,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• être âgé de 18 ans révolus à la date du scrutin,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• travailler depuis 1 an au moins dans l\'entreprise,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• ne pas avoir de lien proche avec l\'employeur (conjoint, partenaire de PACS, concubin, ascendant, descendant, frère, sœur ou allié au même degré) ou disposer d\'une délégation écrite particulière d\'autorité leur permettant d\'être assimilé au chef d\'entreprise ou de le représenter effectivement devant le CSE.',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les personnes mises à disposition, même enregistrées comme électeurs, ne sont pas éligibles au CSE.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 6
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 6 - LISTE DES CANDIDATS', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Pour rappel, le personnel a été informé le ' }),
        new TextRun({ font: 'Roboto', text: data.date_annonce_elections || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ', par ' }),
        new TextRun({ font: 'Roboto', text: data.moyens_information_salaries || 'affichage', bold: true }),
        new TextRun({ font: 'Roboto', text: ', du déroulement des élections.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Pour des raisons d\'ordre matériel tenant à l\'organisation du vote, les dates limites de dépôt des listes sont fixées :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• pour le premier tour : le ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_limite_depot_candidatures_1er || ''} à ${data.heure_limite_depot_candidatures_1er || '12:00'}`, bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• dans l\'éventualité d\'un second tour : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_limite_depot_candidatures_2nd || ''} à ${data.heure_limite_depot_candidatures_2nd || '12:00'}`, bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes de candidats doivent être déposées obligatoirement :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    })
  );
  
  if (data.moyens_depot_candidatures) {
    const moyens = data.moyens_depot_candidatures.split(',');
    moyens.forEach(moyen => {
      paragraphs.push(
        new Paragraph({
          text: `• ${moyen.trim()}`,
          spacing: { left: 360 }
        })
      );
    });
  }
  
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Auprès de : ' }),
        new TextRun({ font: 'Roboto', text: data.responsable_nom_complet || '', bold: true }),
        new TextRun({ font: 'Roboto', text: `, ${data.responsable_poste || ''}` })
      ],
      spacing: { before: 100, left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Email : ' }),
        new TextRun({ font: 'Roboto', text: data.mail_de_la_personne_en_charge_de_lorganisation_des_elections || '', bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Bureau : ' }),
        new TextRun({ font: 'Roboto', text: data.adresse_bureau || data.adresse || '' })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Dans l\'éventualité d\'un second tour, les candidatures présentées au premier tour seront considérées comme maintenues au second tour, sauf si les organisations syndicales déposent de nouvelles listes avant la date limite.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes de candidats seront également transmises au prestataire retenu pour le vote électronique selon le même calendrier.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le prestataire assurera la mise en œuvre des pages du site internet et notamment de la présentation des listes de candidats et des bulletins de vote. Afin de ne pas favoriser une liste ou un vote plutôt qu\'un autre, le prestataire veillera à ce que les dimensions des bulletins et la typographie utilisée soient identiques pour toutes les listes ou choix proposés.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les listes de candidats sont établies distinctement pour chacun des scrutins (Titulaires et Suppléants). Elles peuvent être incomplètes, mais ne doivent pas comporter plus de candidats que le nombre de sièges à pourvoir.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les candidatures doubles (Titulaires et Suppléants) sont autorisées, mais un candidat élu Titulaire ne peut être élu Suppléant.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes sont affichées par la Direction le jour de la clôture des candidatures dans l\'après-midi.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Au premier tour de scrutin, les organisations syndicales suivantes peuvent présenter des candidats :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les syndicats représentatifs dans l\'entreprise et/ou ayant constitué une section syndicale dans l\'entreprise,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• les syndicats affiliés à une organisation reconnue représentative au niveau national et interprofessionnel,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• tout syndicat qui satisfait aux critères de respect des valeurs républicaines et d\'indépendance, légalement constitué depuis au moins deux ans, et dont le champ professionnel et géographique couvre l\'entreprise.',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Au second tour, les candidatures sont libres.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les candidatures individuelles constituent chacune une liste.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Pour rappel, les candidatures présentées au premier tour seront considérées comme maintenues au deuxième tour sauf si les organisations syndicales déposent de nouvelles listes avant la date limite.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les listes des candidats seront communiquées par ' }),
        new TextRun({ font: 'Roboto', text: data.preciser_voie_daffichage || 'affichage', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 7
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 7 - REPRÉSENTATIVITÉ ÉQUILIBRÉE DES FEMMES ET DES HOMMES', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément à l\'article L. 2314-30 du code du travail, pour le premier tour, les listes comportent plusieurs candidats et sont composées d\'un nombre de femmes et d\'hommes correspondant à la part de femmes et d\'hommes inscrits sur la liste électorale.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les listes sont composées alternativement d\'un candidat de chaque sexe jusqu\'à épuisement des candidats d\'un des deux sexes. Ces règles s\'appliquent à la liste de candidature pour le scrutin des élus titulaires et le scrutin des élus suppléants pour le premier tour et le second tour des élections professionnelles à l\'exception des candidatures libres au second tour.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    // CORRECCIÓN 6: Mover la frase después del párrafo anterior
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '—En cas de collège unique, ou un siège par collège supprimer la phrase ci-dessous.—',
                    italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    // CORRECCIÓN 6: Agregar nuevo párrafo con información de composición
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément aux effectifs de ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ', chaque liste devra présenter alternativement ' }),
        new TextRun({ font: 'Roboto', text: calculateFemmesTotal(data), bold: true }),
        new TextRun({ font: 'Roboto', text: ' femmes et ' }),
        new TextRun({ font: 'Roboto', text: calculateHommesTotal(data), bold: true }),
        new TextRun({ font: 'Roboto', text: ' hommes.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 8
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 8 - MODALITÉS DE VOTE', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '8.1 Recours au vote électronique', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les élections professionnelles auront lieu par voie dématérialisée, conformément à l\'accord ou la DUE sur le vote électronique du ' }),
        new TextRun({ font: 'Roboto', text: data.date_signature_due || data.date_signature_pap || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '. La solution technique utilisée pour le vote électronique est celle mise au point et commercialisée par :' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'ELECTIS SOLUTION', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 0 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '83 rue de l\'université', font: 'Roboto'
        })
      ],
      spacing: { after: 0 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '75007 PARIS', font: 'Roboto'
        })
      ],
      spacing: { after: 0 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Immatriculée au registre du commerce et des sociétés de PARIS', font: 'Roboto'
        })
      ],
      spacing: { after: 0 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'sous le numéro 918 956 178', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'L\'URL retenue pour le site de vote est : ' }),
        new TextRun({ font: 'Roboto', text: data.plateforme_vote_url || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.cse.electis.app' })
      ],
      spacing: { after: 100 }
    }),
    // CORRECCIÓN 7: Mover la frase antes del párrafo siguiente
    new Paragraph({
      children: [
        new TextRun({ text: '—Si DUE changer au présent PAP par à la présente décision—', font: 'Roboto',
          italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Conformément à l\'article R. 2314-13 du code du travail, le cahier des charges du système retenu et du déroulement des opérations électorales est annexé au présent Protocole d\'Accord Préélectoral.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '8.2 Envoi du lien de connexion au site de vote électronique', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'La société Electis solution, retenue pour le vote électronique, adressera par email sur l\'adresse de chacun des électeurs un lien OTP (One time password) permettant aux votants de se connecter sur le site de vote.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Si le site de vote est fermé par l\'utilisateur, alors il devra regénérer un lien unique de connexion. Le site le guidera pour le faire.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '8.3 Déroulement des élections professionnelles', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'L\'électeur se connecte grâce à son lien unique d\'authentification, et valide son vote grâce à une troisième information personnelle connue de l\'électeur et définie dans le cahier des charges. L\'électeur reçoit une notice d\'information détaillée sur le déroulement des opérations électorales.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Par ailleurs :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• l\'électeur pourra accéder, durant toute la période de vote, au site de vote gratuitement à partir de tout terminal connecté à internet,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• l\'électeur choisit un scrutin (titulaires ou suppléants). Les scrutins pour lesquels il a déjà voté ne sont plus sélectionnables,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• le service affiche les listes des candidats pour le scrutin choisi,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• l\'électeur peut : choisir une liste complète, raturer des candidats, voter blanc,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• le choix de l\'électeur lui est rappelé et il peut le modifier,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• l\'électeur confirme son vote,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• un accusé de vote est affiché à l\'écran.', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '8.4 Validation et test du système de vote', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Une procédure de validation, de scellement précède l\'ouverture du scrutin. Cette procédure a pour objectif de vérifier les données du système. Elle est menée les membres du bureau de vote assisté ou non par le prestataire.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'À la suite de la procédure, le bureau de vote procède au scellement des urnes électroniques. Le système est alors sécurisé jusqu\'à la fin du scrutin.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le scellement (fermeture des urnes) donne lieu à la création de clés pour chaque membre du bureau de vote appelées clés de chiffrement. Ces clés sont confiées aux membres du bureau de vote et devront être renseignées à l\'issue du scrutin pour accéder aux urnes, lancer le dépouillement et générer les documents de résultats.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '8.5 Assistance aux utilisateurs', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le support niveau 1 et 2 est assuré par les membres du bureau et par toutes personnes ayant suivi au préalable une formation avec les membres du bureau.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'La cellule assistance technique du prestataire sera chargée de veiller au bon fonctionnement et à la supervision technique de ce système de vote. Elle assure le support niveau 3.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Durant la période d\'ouverture, la cellule d\'assistance pourra être contactée par les électeurs à l\'adresse aide@electis.io',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 9
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 9 - BUREAU DE VOTE', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '—Si vous avez 2 collèges, et que les OS sont présents à la réunion vous pouvez négocier un bureau de vote unique et par conséquent changer le paragraphe ci-après.—',
                    italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Compte tenu de l\'organisation par vote électronique, un bureau de vote par collège sera mis en place pour les deux tours de scrutin.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'A défaut, le bureau de vote est composé d\'un président et deux assesseurs.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le bureau de vote sera constitué par appel à volontaires.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le bureau de vote est composé des deux salariés électeurs les plus âgés (dont le président) et du plus jeune des salariés électeurs, sous réserve que tous acceptent cette fonction.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Dans la mesure du possible, le bureau constitué pour le premier tour est conservé à l\'identique pour l\'éventuel second tour.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Afin de veiller au principe de neutralité, le président ne pourra pas être candidat aux élections.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le prestataire formera le bureau de vote à l\'utilisation des outils du site de vote qui lui permettront d\'assurer ses missions.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    // CORRECCIÓN 8: Eliminada la línea sobre formación ficticia
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Durant la période de vote, l\'ensemble des suffrages exprimés sont chiffrés dès leur expression et conservés dans le système de vote. Seuls les détenteurs des clés de déchiffrement pourront, après clôture, déchiffrer les suffrages pour accéder aux résultats.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Au moins 2 clés de déchiffrement sont nécessaires pour générer les opérations de dépouillement des urnes.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le bureau de vote :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Contrôle le déroulement des opérations électorales,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• S\'assure de la régularité du scrutin et du secret du vote,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Procède au dépouillement,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Signe les procès-verbaux des élections.', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'employeur ou son représentant pourra assister au déroulement des opérations électorales à condition de respecter une stricte neutralité.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Un représentant désigné par chaque liste pourra également assister aux opérations électorales dans ces mêmes conditions.', font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 10
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 10 - DATE ET HORAIRES DU VOTE ÉLECTRONIQUE', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'L\'élection peut avoir lieu pendant le temps de travail. Le cas échéant, la participation aux scrutins n\'entraîne aucune perte de salaire pour le salarié.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les plages horaires de vote électronique pour le premier tour sont fixées ainsi :', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Premier tour :', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Scellement de la plateforme de vote : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_ceremonie_cles_1er || ''} à ${data.heure_ceremonie_cles_1er || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Période de vote : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_premier_tour || ''} à ${data.heure_ouverture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Descellement et proclamation des résultats : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_cloture_premier_tour || data.date_premier_tour || ''} à ${data.heure_cloture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'S\'il y a lieu de procéder à un second tour, les plages horaires de vote électronique sont fixées ainsi :',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Second tour :', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Scellement de la plateforme de vote : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_ceremonie_cles_2nd || ''} à ${data.heure_ceremonie_cles_2nd || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Période de vote : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_second_tour || ''} à ${data.heure_ouverture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• Descellement et proclamation des résultats : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_cloture_second_tour || data.date_second_tour || ''} à ${data.heure_cloture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 200 }
    })
  );
  
  // ARTICLE 11
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 11 - CAMPAGNE ÉLECTORALE ET PROPAGANDE DES CANDIDATS', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '11.1 Profession de foi', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes en présence (organisations syndicales et candidats libres au second tour) remettent au plus tard à la Direction leur profession de foi aux dates limites indiquées ci-dessous :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• pour le premier tour : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_limite_depot_candidatures_1er || ''} à ${data.heure_limite_depot_candidatures_1er || '12:00'}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• pour le second tour : ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_limite_depot_candidatures_2nd || ''} à ${data.heure_limite_depot_candidatures_2nd || '12:00'}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les professions de foi seront transmises, à l\'initiative de la Direction des Ressources Humaines, au prestataire retenu pour le vote électronique le même jour que celui de la date limite de dépôt des listes de candidats.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les professions de foi de chaque liste présentée seront affichées sur le site de vote sécurisé Electis solution (prestataire retenu pour la fourniture d\'un site de vote en ligne).', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Celles-ci devront respecter les prérequis suivants :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• format PDF de 500 Ko au maximum', font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• transmission du logo en JPG', font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• format de l\'ensemble du fichier : A4', font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 100 }
    }),
    // CORRECCIÓN 9: Agregar nuevo párrafo sobre logos y fotos
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Par ailleurs, les listes déposées peuvent être accompagnées d\'un logo et / ou d\'une photo, qui seront – le cas échéant – affichés sur le site de vote sécurisé.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Dans ce cas, le logo devra respecter les prérequis suivants : format jpg de 500 KO au maximum, taille de 100 pixels x 100 pixels. La photo devra respecter les prérequis suivants : format jpg de 500 KO au maximum, taille de 50 pixels x 50 pixels.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Les professions de foi et logos doivent être envoyés selon les mêmes modalités que celles prévues pour le dépôt des listes de candidats, c\'est-à-dire par voie dématérialisée à l\'adresse ' }),
        new TextRun({ font: 'Roboto', text: data.mail_de_la_personne_en_charge_de_lorganisation_des_elections || '', bold: true })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes en présence assurent leur propagande électorale dans le cadre des dispositions légales en vigueur et notamment de l\'article L. 2142-4 du code du travail.', font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '11.2 Fin de la campagne électorale', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'En référence aux articles 49 et suivants du code électoral, il est convenu qu\'aucune diffusion de tract ou autre support de propagande électorale ne se fera durant la période de vote. En conséquence, la propagande électorale prendra fin 72H avant chaque scrutin.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 12
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 12 - DURÉE DES MANDATS', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'La durée des mandats est légalement fixée à 4 années. ' }),
        new TextRun({ font: 'Roboto', text: '(Modifier si ce n\'est pas votre cas, ne peut être inférieur à 2 ans)' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le nombre de mandats successifs est limité à 3, conformément aux dispositions légales en vigueur.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 13
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 13 - DÉPOUILLEMENT', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'La date de dépouillement des élections est fixée :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• pour le premier tour : le ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_cloture_premier_tour || data.date_premier_tour || ''} à ${data.heure_cloture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• pour le second tour, le cas échéant : le ' }),
        new TextRun({ font: 'Roboto', text: `${data.date_cloture_second_tour || data.date_second_tour || ''} à ${data.heure_cloture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le processus de dépouillement est le suivant :', font: 'Roboto'
        })
      ],
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• clôture du site internet de vote,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• déchiffrement des suffrages à l\'aide des clés des membres du bureau de vote,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• calcul automatique des résultats et attribution des sièges,',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: '• téléchargement des listes d\'émargement, des procès-verbaux, impression et envoi des procès-verbaux aux membres du bureau pour signature.',
          font: 'Roboto'
        })
      ],
      spacing: { left: 360, after: 200 }
    })
  );
  
  // ARTICLE 14
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 14 - PROCLAMATION DES RÉSULTATS', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les résultats sont proclamés par le Service Ressources Humaines. Un procès-verbal par urne original est établi permettant de faire état des résultats de chaque scrutin. Ces procès-verbaux sont signés par les membres du bureau de vote.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Chaque liste ayant présenté des candidats peut se faire remettre une copie de ce procès-verbal sur simple demande.', font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Dès le lendemain des élections professionnelles, les résultats seront communiqués à l\'ensemble des salariés par tous moyens.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Une fois cette action réalisée, ELECTIS SOLUTIONS est invité à télétransmettre les élections auprès du CTEP dans les quinze jours suivant la tenue des élections professionnelles. Puis l\'entreprise pourra téléverser chaque PV sur la plateforme du CTEP.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 15
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: 'ARTICLE 15 - DÉPÔT ET PUBLICITÉ DU PROTOCOLE (OU DE LA DÉCISION)', font: 'Roboto'
        })
      ],
      heading: HeadingLevel.HEADING_2,
      bold: true,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '—Si DUE supprimer les 3 paragraphes ci-dessous puis changer le présent accord en La décision—', font: 'Roboto',
          italics: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le présent protocole d\'accord sera notifié, par lettre recommandée avec avis de réception ou par courrier remis en mains propres contre récépissé, à l\'ensemble des organisations syndicales ayant participé à la négociation du présent accord, signataires ou non.',
          font: 'Roboto'
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'En application du décret n° 2018-362 du 15 mai 2018 relatif à la procédure de dépôt des accords collectifs, les formalités de dépôt seront effectuées par le représentant de ' }),
        new TextRun({ font: 'Roboto', text: data.nom_entreprise || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '. Ce dernier déposera le présent accord sur la plateforme nationale « Télé Accords » à l\'adresse suivante : https://accords-depot.travail.gouv.fr' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Le déposant déposera un exemplaire de l\'accord au secrétariat greffe du Conseil de Prud\'hommes de ' }),
        new TextRun({ font: 'Roboto', text: data.ville || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '. Les parties rappellent que, dans un acte distinct du présent accord, elles pourront convenir qu\'une partie du présent accord ne fera pas l\'objet de la publication prévue à l\'article L. 2231-5-1 du Code du Travail. A défaut d\'un tel acte, le présent accord sera publié dans une version intégrale, ne comportant pas les noms et prénoms des négociateurs et des signataires.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le présent accord sera porté à la connaissance des salariés par tous moyen.', font: 'Roboto'
        })
      ],
      spacing: { after: 200 }
    })
  );
  
  // SIGNATURES
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: '', font: 'Roboto'
        })
      ],
      spacing: { before: 400 }
    }),
    new Paragraph({
      children: [
        new TextRun({ font: 'Roboto', text: 'Fait à ' }),
        new TextRun({ font: 'Roboto', text: data.ville || '', bold: true }),
        new TextRun({ font: 'Roboto', text: ', le ' }),
        new TextRun({ font: 'Roboto', text: data.date_signature_pap || '', bold: true }),
        new TextRun({ font: 'Roboto', text: '.' })
      ],
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Pour la Société :', font: 'Roboto',
          bold: true
        })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: data.personne_nom_complet || data.dirigeant_nom || ''
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
        new TableCell({ 
          children: [new Paragraph({ text: 'Collège', bold: true })],
          width: { size: 40, type: WidthType.PERCENTAGE }
        }),
        new TableCell({ 
          children: [new Paragraph({ text: 'Hommes', bold: true })],
          width: { size: 20, type: WidthType.PERCENTAGE }
        }),
        new TableCell({ 
          children: [new Paragraph({ text: 'Femmes', bold: true })],
          width: { size: 20, type: WidthType.PERCENTAGE }
        }),
        new TableCell({ 
          children: [new Paragraph({ text: 'Total', bold: true })],
          width: { size: 20, type: WidthType.PERCENTAGE }
        }),
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
    new Table({ 
      rows, 
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      }
    }),
    new Paragraph({ text: '', spacing: { after: 200 } })
  ];
}

function createSiegesTable(data) {
  const rows = [
    new TableRow({
      children: [
        new TableCell({ 
          children: [new Paragraph({ text: 'Collège', bold: true })],
          width: { size: 50, type: WidthType.PERCENTAGE }
        }),
        new TableCell({ 
          children: [new Paragraph({ text: 'Titulaires', bold: true })],
          width: { size: 25, type: WidthType.PERCENTAGE }
        }),
        new TableCell({ 
          children: [new Paragraph({ text: 'Suppléants', bold: true })],
          width: { size: 25, type: WidthType.PERCENTAGE }
        }),
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
    new Table({ 
      rows, 
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      }
    }),
    new Paragraph({ text: '', spacing: { after: 200 } })
  ];
}

// ========================================
// HELPERS
// ========================================

function calculateFemmesTotal(data) {
  if (!data.colleges) return '0';
  let total = 0;
  for (const college of Object.values(data.colleges)) {
    total += parseInt(college.femmes || 0);
  }
  return String(total);
}

function calculateHommesTotal(data) {
  if (!data.colleges) return '0';
  let total = 0;
  for (const college of Object.values(data.colleges)) {
    total += parseInt(college.hommes || 0);
  }
  return String(total);
}

function sanitizeFilename(filename) {
  return (filename || 'document').replace(/[^a-z0-9]/gi, '_').toLowerCase();
}
