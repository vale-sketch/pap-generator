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
  
  
    new Paragraph({
      text: '–﻿Toutes les mentions entre tirets, comme dans cet exemple : –exemple–, sont des indications d’aide et doivent être supprimées.--',
      spacing: { after: 100 }
    }),

// HEADER
  paragraphs.push(
    new Paragraph({
      text: 'PROTOCOLE D\'ACCORD PRÉÉLECTORAL RELATIF À LA MISE EN PLACE DU COMITÉ SOCIAL ET ÉCONOMIQUE (CSE)',
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 }
    }),
    new Paragraph({
      text: '–OU (garder la mention qui correspond à votre situation) –',
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 }
    }),
    new Paragraph({
      text: 'DÉCISION UNILATÉRALE DE L’EMPLOYEUR (DUE) RELATIVE A LA MISE EN PLACE DU COMITE SOCIAL ET ÉCONOMIQUE (CSE)',
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 }
    })
  );
  
  // ENTRE
  paragraphs.push(
    new Paragraph({
      text: 'ENTRE :',
      bold: true,
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      text: ` –(supprimer si c’est une DUE garder que le paragraphe ci-dessous)--`,
      spacing: { after: 200 }
    })
  );
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
  
  // SYNDICATS
  if (data.syndicats && data.syndicats.length > 0) {
    paragraphs.push(
      new Paragraph({
        text: 'ET :',
        bold: true,
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        text: 'Les organisations syndicales ayant répondu à l\'invitation de négocier le présent Protocole d\'Accord Préélectoral, ayant dûment mandaté :',
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
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Dans le cadre de la mise en place du Comité Social et Economique (CSE), la Direction a invité les Organisations Syndicales visées à l\'article L. 2314-5 du code du travail à négocier le Protocole d\'Accord Préélectoral, par lettre recommandée avec accusé de réception ' }),
        new TextRun({ text: data.date_invitation_os || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les autres Organisations syndicales intéressées ont été informées de l\'organisation des élections et invitées à négocier le protocole d\'accord préélectoral par voie d\'affichage dans les locaux de la société ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: ' le ' }),
        new TextRun({ text: data.date_invitation_os || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le ' }),
        new TextRun({ text: data.date_reunion_negociation || '', bold: true }),
        new TextRun({ text: ', les organisations syndicales précitées se sont présentées à la table des négociations en vue de négocier, notamment la répartition du personnel et des sièges dans les différents collèges électoraux ainsi que les modalités d\'organisation et de déroulement des élections professionnelles.' })
      ],
      spacing: { after: 150 }
    }),
    new Paragraph({
      text: 'Recours au vote électronique :',
      bold: true,
      spacing: { before: 200, after: 100 }
    }),
    new Paragraph({
      text: 'L\'article 54 de la loi n° 2004-575 du 21 juin 2004 pour la confiance dans l\'économie numérique autorise le recours au vote électronique pour les élections professionnelles.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les articles R. 2314-5 à R. 2314-18 du code du travail définissent les modalités de vote par voie électronique pour l\'élection des membres de la délégation du personnel du CSE.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'La décision unilatérale signée le ' }),
        new TextRun({ text: data.date_signature_due || data.date_signature_pap || '', bold: true }),
        new TextRun({ text: ' a autorisé l\'utilisation du vote électronique pour l\'élection des membres de la délégation du personnel du CSE. L\'accompagnement ainsi que la mise à disposition de la plateforme de vote en ligne ont été confiés à la société ELECTIS, pour les élections régies par le présent accord.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 1
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 1 - OBJET ET CHAMP D\'APPLICATION',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le présent accord a pour objet de définir les modalités d\'organisation de l\'élection des représentants du personnel au CSE de la société ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Conformément aux dispositions légales, le présent accord a pour champ d\'application le CSE de l\'entreprise ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 2
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 2 - DATE DES ÉLECTIONS ET HORAIRES DU SCRUTIN',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Pour le premier tour, la date d\'ouverture du scrutin est fixée pour l\'ensemble des collèges le ' }),
        new TextRun({ text: `${data.date_premier_tour || ''} à ${data.heure_ouverture_premier_tour || ''}`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Au premier tour, sont habilitées à présenter leur liste de candidats, les organisations syndicales représentatives dans l\'entreprise, les syndicats affiliés à une organisation reconnue représentative aux niveaux national et interprofessionnel, et tout syndicat qui satisfait aux critères de respect des valeurs républicaines, d\'indépendance, légalement constitué depuis au moins 2 ans, et dont le champ professionnel et géographique couvre l\'entreprise concernée.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Il y aura lieu de procéder à un second tour, ouvert aux candidats libres et aux candidatures syndicales dans les conditions définies pour le premier tour, dans l\'un des cas suivants :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• Quorum non atteint au premier tour : moins de la moitié des électeurs inscrits ont émis un vote valable.',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• Carence : absence de candidature syndicale au premier tour.',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• Non attribution de siège : les sièges n\'ont pas tous été pourvus dès le premier tour.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Les cas ci-dessus s\'apprécient par collège et par scrutin.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le cas échéant, la date d\'ouverture du scrutin du deuxième tour est fixée le ' }),
        new TextRun({ text: `${data.date_second_tour || ''} à ${data.heure_ouverture_second_tour || ''}`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les collèges concernés par cette élection sont composés de :',
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
      text: 'ARTICLE 3 - EFFECTIFS ET NOMBRE DE SIÈGES',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'L\'effectif à la date du premier jour du 1er tour de scrutin est calculé selon les règles prévues aux articles L. 1111-2 et L. 1111-3 du code du travail.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Plus précisément, y sont inclus :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• les CDI à temps plein,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les CDD (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les salariés mis à la disposition de l\'entreprise par une entreprise extérieure, qui sont présents dans les locaux de l\'entreprise utilisatrice et y travaillent depuis au moins un an (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les salariés temporaires sont pris en compte dans l\'effectif de l\'entreprise proportionnellement à leur temps de présence au cours des douze mois précédents (sauf s\'ils remplacent un salarié absent ou dont le contrat est suspendu),',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les salariés à temps partiel en fonction de leur durée du travail.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'En sont exclus :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• les apprentis,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les stagiaires,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les titulaires d\'un contrat de professionnalisation, jusqu\'au terme prévu par le contrat lorsque celui-ci est à durée déterminée, ou, jusqu\'à la fin de l\'action de professionnalisation lorsque le contrat est à durée indéterminée.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'effectif (ETP : équivalent temps plein) de référence servant à déterminer le nombre de sièges est de ' }),
        new TextRun({ text: `${data.nombre_effectif || ''}`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Ci-après le détail :',
      spacing: { before: 100, after: 100 }
    })
  );
  
  // Tables
  if (data.colleges && Object.keys(data.colleges).length > 0) {
    paragraphs.push(...createGenreTable(data));
    paragraphs.push(...createSiegesTable(data));
  }
  
  // ARTICLE 4
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 4 - CONDITIONS POUR ÊTRE ÉLECTEUR',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Conformément aux dispositions de l\'article L. 2314-18 du code du travail, sont électeurs les salariés de l\'entreprise qui remplissent les conditions suivantes à la date du premier tour du scrutin :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• être titulaire d\'un contrat de travail,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• être âgé de 16 ans révolus,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• n\'avoir fait l\'objet d\'aucune interdiction, déchéance ou incapacité relatives à leurs droits civiques,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• travailler depuis 3 mois au moins dans l\'entreprise à la date du premier tour des élections.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Par application de l\'article L. 2314-23 du code du travail, pour les salariés mis à disposition qui remplissent les conditions mentionnées au 2° de l\'article L.1111-2 du code du travail, la condition de présence dans l\'entreprise utilisatrice est de douze mois continus pour y être électeur.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'A cet effet, la Direction constate qu\'il n\'y a pas d\'entreprises prestataires ayant du personnel mis à disposition de la société. Il n\'y a donc aucune entreprise prestataire à contacter afin qu\'elle fournisse la liste des salariés mis à disposition répondant aux critères de présence dans les locaux et d\'ancienneté et qui souhaiteraient voter au sein de la société.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 5
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 5 - CONDITIONS POUR ÊTRE ÉLIGIBLE',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Conformément aux dispositions de l\'article L.2314-19 du code du travail, pour être éligible, un salarié doit répondre aux conditions suivantes pour chaque tour de scrutin :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• être électeur,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• être âgé de 18 ans révolus à la date du scrutin,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• travailler depuis 1 an au moins dans l\'entreprise,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• ne pas avoir de lien proche avec l\'employeur (conjoint, partenaire de PACS, concubin, ascendant, descendant, frère, sœur ou allié au même degré) ou disposer d\'une délégation écrite particulière d\'autorité leur permettant d\'être assimilé au chef d\'entreprise ou de le représenter effectivement devant le CSE.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Les personnes mises à disposition, même enregistrées comme électeurs, ne sont pas éligibles au CSE.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 6
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 6 - LISTE DES CANDIDATS',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Pour rappel, le personnel a été informé le ' }),
        new TextRun({ text: data.date_annonce_elections || '', bold: true }),
        new TextRun({ text: ', par ' }),
        new TextRun({ text: data.moyens_information_salaries || 'affichage', bold: true }),
        new TextRun({ text: ', du déroulement des élections.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Pour des raisons d\'ordre matériel tenant à l\'organisation du vote, les dates limites de dépôt des listes sont fixées :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• pour le premier tour : le ' }),
        new TextRun({ text: `${data.date_limite_depot_candidatures_1er || ''} à ${data.heure_limite_depot_candidatures_1er || '12:00'}`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• dans l\'éventualité d\'un second tour : ' }),
        new TextRun({ text: `${data.date_limite_depot_candidatures_2nd || ''} à ${data.heure_limite_depot_candidatures_2nd || '12:00'}`, bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Les listes de candidats doivent être déposées obligatoirement :',
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
        new TextRun({ text: 'Auprès de : ' }),
        new TextRun({ text: data.responsable_nom_complet || '', bold: true }),
        new TextRun({ text: `, ${data.responsable_poste || ''}` })
      ],
      spacing: { before: 100, left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Email : ' }),
        new TextRun({ text: data.mail_de_la_personne_en_charge_de_lorganisation_des_elections || '', bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Bureau : ' }),
        new TextRun({ text: data.adresse_bureau || data.adresse || '' })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Dans l\'éventualité d\'un second tour, les candidatures présentées au premier tour seront considérées comme maintenues au second tour, sauf si les organisations syndicales déposent de nouvelles listes avant la date limite.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes de candidats seront également transmises au prestataire retenu pour le vote électronique selon le même calendrier.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le prestataire assurera la mise en œuvre des pages du site internet et notamment de la présentation des listes de candidats et des bulletins de vote. Afin de ne pas favoriser une liste ou un vote plutôt qu\'un autre, le prestataire veillera à ce que les dimensions des bulletins et la typographie utilisée soient identiques pour toutes les listes ou choix proposés.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes de candidats sont établies distinctement pour chacun des scrutins (Titulaires et Suppléants). Elles peuvent être incomplètes, mais ne doivent pas comporter plus de candidats que le nombre de sièges à pourvoir.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les candidatures doubles (Titulaires et Suppléants) sont autorisées, mais un candidat élu Titulaire ne peut être élu Suppléant.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes sont affichées par la Direction le jour de la clôture des candidatures dans l\'après-midi.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Au premier tour de scrutin, les organisations syndicales suivantes peuvent présenter des candidats :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• les syndicats représentatifs dans l\'entreprise et/ou ayant constitué une section syndicale dans l\'entreprise,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• les syndicats affiliés à une organisation reconnue représentative au niveau national et interprofessionnel,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• tout syndicat qui satisfait aux critères de respect des valeurs républicaines et d\'indépendance, légalement constitué depuis au moins deux ans, et dont le champ professionnel et géographique couvre l\'entreprise.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Au second tour, les candidatures sont libres.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les candidatures individuelles constituent chacune une liste.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Pour rappel, les candidatures présentées au premier tour seront considérées comme maintenues au deuxième tour sauf si les organisations syndicales déposent de nouvelles listes avant la date limite.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les listes des candidats seront communiquées par ' }),
        new TextRun({ text: data.preciser_voie_daffichage || 'affichage', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 7
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 7 - REPRÉSENTATIVITÉ ÉQUILIBRÉE DES FEMMES ET DES HOMMES',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Conformément à l\'article L. 2314-30 du code du travail, pour le premier tour, les listes comportent plusieurs candidats et sont composées d\'un nombre de femmes et d\'hommes correspondant à la part de femmes et d\'hommes inscrits sur la liste électorale.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes sont composées alternativement d\'un candidat de chaque sexe jusqu\'à épuisement des candidats d\'un des deux sexes. Ces règles s\'appliquent à la liste de candidature pour le scrutin des élus titulaires et le scrutin des élus suppléants pour le premier tour et le second tour des élections professionnelles à l\'exception des candidatures libres au second tour.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 8
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 8 - MODALITÉS DE VOTE',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: '8.1 Recours au vote électronique',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les élections professionnelles auront lieu par voie dématérialisée, conformément à l\'accord ou la DUE sur le vote électronique du ' }),
        new TextRun({ text: data.date_signature_due || data.date_signature_pap || '', bold: true }),
        new TextRun({ text: '. La solution technique utilisée pour le vote électronique est celle mise au point et commercialisée par :' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'ELECTIS SOLUTION',
      bold: true,
      spacing: { after: 0 }
    }),
    new Paragraph({
      text: '83 rue de l\'université',
      spacing: { after: 0 }
    }),
    new Paragraph({
      text: '75007 PARIS',
      spacing: { after: 0 }
    }),
    new Paragraph({
      text: 'Immatriculée au registre du commerce et des sociétés de PARIS',
      spacing: { after: 0 }
    }),
    new Paragraph({
      text: 'sous le numéro 918 956 178',
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'L\'URL retenue pour le site de vote est : ' }),
        new TextRun({ text: data.plateforme_vote_url || '', bold: true }),
        new TextRun({ text: '.cse.electis.app' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Conformément à l\'article R. 2314-13 du code du travail, le cahier des charges du système retenu et du déroulement des opérations électorales est annexé au présent Protocole d\'Accord Préélectoral.',
      spacing: { after: 200 }
    }),
    new Paragraph({
      text: '8.2 Envoi du lien de connexion au site de vote électronique',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'La société Electis solution, retenue pour le vote électronique, adressera par email sur l\'adresse de chacun des électeurs un lien OTP (One time password) permettant aux votants de se connecter sur le site de vote.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Si le site de vote est fermé par l\'utilisateur, alors il devra regénérer un lien unique de connexion. Le site le guidera pour le faire.',
      spacing: { after: 200 }
    }),
    new Paragraph({
      text: '8.3 Déroulement des élections professionnelles',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'L\'électeur se connecte grâce à son lien unique d\'authentification, et valide son vote grâce à une troisième information personnelle connue de l\'électeur et définie dans le cahier des charges. L\'électeur reçoit une notice d\'information détaillée sur le déroulement des opérations électorales.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Par ailleurs :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• l\'électeur pourra accéder, durant toute la période de vote, au site de vote gratuitement à partir de tout terminal connecté à internet,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• l\'électeur choisit un scrutin (titulaires ou suppléants). Les scrutins pour lesquels il a déjà voté ne sont plus sélectionnables,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• le service affiche les listes des candidats pour le scrutin choisi,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• l\'électeur peut : choisir une liste complète, raturer des candidats, voter blanc,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• le choix de l\'électeur lui est rappelé et il peut le modifier,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• l\'électeur confirme son vote,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• un accusé de vote est affiché à l\'écran.',
      spacing: { left: 360, after: 200 }
    }),
    new Paragraph({
      text: '8.4 Validation et test du système de vote',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Une procédure de validation, de scellement précède l\'ouverture du scrutin. Cette procédure a pour objectif de vérifier les données du système. Elle est menée les membres du bureau de vote assisté ou non par le prestataire.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'À la suite de la procédure, le bureau de vote procède au scellement des urnes électroniques. Le système est alors sécurisé jusqu\'à la fin du scrutin.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le scellement (fermeture des urnes) donne lieu à la création de clés pour chaque membre du bureau de vote appelées clés de chiffrement. Ces clés sont confiées aux membres du bureau de vote et devront être renseignées à l\'issue du scrutin pour accéder aux urnes, lancer le dépouillement et générer les documents de résultats.',
      spacing: { after: 200 }
    }),
    new Paragraph({
      text: '8.5 Assistance aux utilisateurs',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le support niveau 1 et 2 est assuré par les membres du bureau et par toutes personnes ayant suivi au préalable une formation avec les membres du bureau.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'La cellule assistance technique du prestataire sera chargée de veiller au bon fonctionnement et à la supervision technique de ce système de vote. Elle assure le support niveau 3.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Durant la période d\'ouverture, la cellule d\'assistance pourra être contactée par les électeurs à l\'adresse aide@electis.io',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 9
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 9 - BUREAU DE VOTE',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Compte tenu de l\'organisation par vote électronique, un bureau de vote par collège sera mis en place pour les deux tours de scrutin.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'A défaut, le bureau de vote est composé d\'un président et deux assesseurs.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le bureau de vote sera constitué par appel à volontaires.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le bureau de vote est composé des deux salariés électeurs les plus âgés (dont le président) et du plus jeune des salariés électeurs, sous réserve que tous acceptent cette fonction.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Dans la mesure du possible, le bureau constitué pour le premier tour est conservé à l\'identique pour l\'éventuel second tour.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Afin de veiller au principe de neutralité, le président ne pourra pas être candidat aux élections.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le prestataire formera le bureau de vote à l\'utilisation des outils du site de vote qui lui permettront d\'assurer ses missions. Lors de cette formation, chaque membre du bureau réalisera une élection fictive.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Durant la période de vote, l\'ensemble des suffrages exprimés sont chiffrés dès leur expression et conservés dans le système de vote. Seuls les détenteurs des clés de déchiffrement pourront, après clôture, déchiffrer les suffrages pour accéder aux résultats.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Au moins 2 clés de déchiffrement sont nécessaires pour générer les opérations de dépouillement des urnes.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le bureau de vote :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• Contrôle le déroulement des opérations électorales,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• S\'assure de la régularité du scrutin et du secret du vote,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• Procède au dépouillement,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• Signe les procès-verbaux des élections.',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'L\'employeur ou son représentant pourra assister au déroulement des opérations électorales à condition de respecter une stricte neutralité.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Un représentant désigné par chaque liste pourra également assister aux opérations électorales dans ces mêmes conditions.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 10
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 10 - DATE ET HORAIRES DU VOTE ÉLECTRONIQUE',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'L\'élection peut avoir lieu pendant le temps de travail. Le cas échéant, la participation aux scrutins n\'entraîne aucune perte de salaire pour le salarié.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les plages horaires de vote électronique pour le premier tour sont fixées ainsi :',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Premier tour :',
      bold: true,
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Scellement de la plateforme de vote : ' }),
        new TextRun({ text: `${data.date_ceremonie_cles_1er || ''} à ${data.heure_ceremonie_cles_1er || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Période de vote : ' }),
        new TextRun({ text: `${data.date_premier_tour || ''} à ${data.heure_ouverture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Descellement et proclamation des résultats : ' }),
        new TextRun({ text: `${data.date_cloture_premier_tour || data.date_premier_tour || ''} à ${data.heure_cloture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'S\'il y a lieu de procéder à un second tour, les plages horaires de vote électronique sont fixées ainsi :',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Second tour :',
      bold: true,
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Scellement de la plateforme de vote : ' }),
        new TextRun({ text: `${data.date_ceremonie_cles_2nd || ''} à ${data.heure_ceremonie_cles_2nd || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Période de vote : ' }),
        new TextRun({ text: `${data.date_second_tour || ''} à ${data.heure_ouverture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• Descellement et proclamation des résultats : ' }),
        new TextRun({ text: `${data.date_cloture_second_tour || data.date_second_tour || ''} à ${data.heure_cloture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 200 }
    })
  );
  
  // ARTICLE 11
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 11 - CAMPAGNE ÉLECTORALE ET PROPAGANDE DES CANDIDATS',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: '11.1 Profession de foi',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes en présence (organisations syndicales et candidats libres au second tour) remettent au plus tard à la Direction leur profession de foi aux dates limites indiquées ci-dessous :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• pour le premier tour : ' }),
        new TextRun({ text: `${data.date_limite_depot_candidatures_1er || ''} à ${data.heure_limite_depot_candidatures_1er || '12:00'}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• pour le second tour : ' }),
        new TextRun({ text: `${data.date_limite_depot_candidatures_2nd || ''} à ${data.heure_limite_depot_candidatures_2nd || '12:00'}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Les professions de foi seront transmises, à l\'initiative de la Direction des Ressources Humaines, au prestataire retenu pour le vote électronique le même jour que celui de la date limite de dépôt des listes de candidats.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les professions de foi de chaque liste présentée seront affichées sur le site de vote sécurisé Electis solution (prestataire retenu pour la fourniture d\'un site de vote en ligne).',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Celles-ci devront respecter les prérequis suivants :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• format PDF de 500 Ko au maximum',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• transmission du logo en JPG',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• format de l\'ensemble du fichier : A4',
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Les professions de foi et logos doivent être envoyés selon les mêmes modalités que celles prévues pour le dépôt des listes de candidats, c\'est-à-dire par voie dématérialisée à l\'adresse ' }),
        new TextRun({ text: data.mail_de_la_personne_en_charge_de_lorganisation_des_elections || '', bold: true })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Les listes en présence assurent leur propagande électorale dans le cadre des dispositions légales en vigueur et notamment de l\'article L. 2142-4 du code du travail.',
      spacing: { after: 200 }
    }),
    new Paragraph({
      text: '11.2 Fin de la campagne électorale',
      bold: true,
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'En référence aux articles 49 et suivants du code électoral, il est convenu qu\'aucune diffusion de tract ou autre support de propagande électorale ne se fera durant la période de vote. En conséquence, la propagande électorale prendra fin 72H avant chaque scrutin.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 12
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 12 - DURÉE DES MANDATS',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'La durée des mandats est légalement fixée à 4 années.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le nombre de mandats successifs est limité à 3, conformément aux dispositions légales en vigueur.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 13
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 13 - DÉPOUILLEMENT',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'La date de dépouillement des élections est fixée :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• pour le premier tour : le ' }),
        new TextRun({ text: `${data.date_cloture_premier_tour || data.date_premier_tour || ''} à ${data.heure_cloture_premier_tour || ''}`, bold: true })
      ],
      spacing: { left: 360 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '• pour le second tour, le cas échéant : le ' }),
        new TextRun({ text: `${data.date_cloture_second_tour || data.date_second_tour || ''} à ${data.heure_cloture_second_tour || ''}`, bold: true })
      ],
      spacing: { left: 360, after: 100 }
    }),
    new Paragraph({
      text: 'Le processus de dépouillement est le suivant :',
      spacing: { after: 50 }
    }),
    new Paragraph({
      text: '• clôture du site internet de vote,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• déchiffrement des suffrages à l\'aide des clés des membres du bureau de vote,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• calcul automatique des résultats et attribution des sièges,',
      spacing: { left: 360 }
    }),
    new Paragraph({
      text: '• téléchargement des listes d\'émargement, des procès-verbaux, impression et envoi des procès-verbaux aux membres du bureau pour signature.',
      spacing: { left: 360, after: 200 }
    })
  );
  
  // ARTICLE 14
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 14 - PROCLAMATION DES RÉSULTATS',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Les résultats sont proclamés par le Service Ressources Humaines. Un procès-verbal par urne original est établi permettant de faire état des résultats de chaque scrutin. Ces procès-verbaux sont signés par les membres du bureau de vote.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Chaque liste ayant présenté des candidats peut se faire remettre une copie de ce procès-verbal sur simple demande.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Dès le lendemain des élections professionnelles, les résultats seront communiqués à l\'ensemble des salariés par tous moyens.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Une fois cette action réalisée, ELECTIS SOLUTIONS est invité à télétransmettre les élections auprès du CTEP dans les quinze jours suivant la tenue des élections professionnelles. Puis l\'entreprise pourra téléverser chaque PV sur la plateforme du CTEP.',
      spacing: { after: 200 }
    })
  );
  
  // ARTICLE 15
  paragraphs.push(
    new Paragraph({
      text: 'ARTICLE 15 - DÉPÔT ET PUBLICITÉ DU PROTOCOLE',
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Le présent protocole d\'accord sera notifié, par lettre recommandée avec avis de réception ou par courrier remis en mains propres contre récépissé, à l\'ensemble des organisations syndicales ayant participé à la négociation du présent accord, signataires ou non.',
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'En application du décret n° 2018-362 du 15 mai 2018 relatif à la procédure de dépôt des accords collectifs, les formalités de dépôt seront effectuées par le représentant de ' }),
        new TextRun({ text: data.nom_entreprise || '', bold: true }),
        new TextRun({ text: '. Ce dernier déposera le présent accord sur la plateforme nationale « Télé Accords » à l\'adresse suivante : https://accords-depot.travail.gouv.fr' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Le déposant déposera un exemplaire de l\'accord au secrétariat greffe du Conseil de Prud\'hommes de ' }),
        new TextRun({ text: data.ville || '', bold: true }),
        new TextRun({ text: '. Les parties rappellent que, dans un acte distinct du présent accord, elles pourront convenir qu\'une partie du présent accord ne fera pas l\'objet de la publication prévue à l\'article L. 2231-5-1 du Code du Travail. A défaut d\'un tel acte, le présent accord sera publié dans une version intégrale, ne comportant pas les noms et prénoms des négociateurs et des signataires.' })
      ],
      spacing: { after: 100 }
    }),
    new Paragraph({
      text: 'Le présent accord sera porté à la connaissance des salariés par tous moyen.',
      spacing: { after: 200 }
    })
  );
  
  // SIGNATURES
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
        new TextRun({ text: data.date_signature_pap || '', bold: true }),
        new TextRun({ text: '.' })
      ],
      spacing: { before: 400, after: 200 }
    }),
    new Paragraph({
      text: 'Pour la Société :',
      bold: true,
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

function sanitizeFilename(filename) {
  return (filename || 'document').replace(/[^a-z0-9]/gi, '_').toLowerCase();
}
