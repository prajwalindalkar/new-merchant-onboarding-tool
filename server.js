const express = require('express');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign, UnderlineType, PageBreak
} = require('docx');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────
const border  = { style: BorderStyle.SINGLE, size: 1, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder };

const p  = (children, opts = {}) => new Paragraph({ children, spacing: { after: 120 }, ...opts });
const t  = (text, opts = {}) => new TextRun({ text: text || '', font: 'Times New Roman', size: 22, ...opts });
const th = (text) => new TextRun({ text: text || '', font: 'Times New Roman', size: 20, bold: true });

function cell(children, w, opts = {}) {
  return new TableCell({
    borders,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    width: { size: w, type: WidthType.DXA },
    ...opts,
    children: Array.isArray(children) ? children
      : [new Paragraph({ children: [t(String(children || ''))], spacing: { after: 0 } })]
  });
}
function hCell(text, w) {
  return new TableCell({
    borders,
    shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ children: [th(text)], spacing: { after: 0 } })]
  });
}

function sigTable(partners) {
  const n = partners.length;
  const colW = Math.floor(9360 / n);
  const cwArr = partners.map(() => colW);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: cwArr,
    borders: noBorders,
    rows: [
      new TableRow({
        children: partners.map(() => new TableCell({
          borders: noBorders,
          width: { size: colW, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({ children: [t('____________Sign____________')], spacing: { after: 60 } })]
        }))
      }),
      new TableRow({
        children: partners.map(pp => new TableCell({
          borders: noBorders,
          width: { size: colW, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({ children: [t(pp.name)], spacing: { after: 60 } })]
        }))
      }),
    ]
  });
}

// ─────────────────────────────────────────────────────────
// BUILD RESOLUTION
// ─────────────────────────────────────────────────────────
async function buildResolution(d) {
  const { firmName, regAddress, principalSame, principalAddress,
          resolutionDate, resolutionTime, authSignatoryName,
          presentPartners, isRegistered, deedDate,
          naturalPersons, pepDeclaration } = d;

  const principalText = principalSame
    ? 'same \u2611 (tick if applicable)'
    : principalAddress;

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        // ── PAGE 1: RESOLUTION ──────────────────────────────
        p([t('(on partnership firm\u2019s letterhead)', { italics: true })], { alignment: AlignmentType.CENTER }),
        p([t('TO WHOMSOEVER IT MAY CONCERN', { bold: true })], { alignment: AlignmentType.CENTER }),
        p([]),

        new Paragraph({
          children: [
            t('RESOLUTION OF THE PARTNERS PASSED AT THE MEETING OF THE PARTNERS OF ', { bold: true }),
            t(firmName.toUpperCase(), { bold: true }),
            t(' |NAME OF THE PARTNERSHIP FIRM| (\u201cFIRM\u201d) HELD ON ', { bold: true }),
            t(resolutionDate, { bold: true }),
            t('\n[DATE] AT ', { bold: true }),
            t(resolutionTime + ' [TIME]', { bold: true }),
            t('  having its <registered office address> at\n'),
            t(regAddress, { bold: true }),
            t(' and having its <principal place of operation/office>\nat '),
            t(principalText),
            t(','),
          ],
          spacing: { after: 200 }
        }),
        p([]),

        p([t('PRESENT:', { bold: true })]),
        ...presentPartners.map((pp, i) =>
          p([t((i + 1) + '. ' + pp.name)])
        ),
        p([t('(List of partner present during the resolution)', { italics: true })]),
        p([]),

        new Paragraph({
          children: [
            t('RESOLVED THAT ', { bold: true }),
            t('Mr/Mrs '),
            t(authSignatoryName + ' [Name of Partner ], < Partner>'),
            t(',be and is hereby authorized, to act on behalf of the Firm and to execute/sign all necessary applications/ documents for the purpose of opening and operating a business account with PhonePe Limited.'),
          ],
          spacing: { after: 200 }
        }),
        p([]),

        new Paragraph({
          children: [
            t('RESOLVED FURTHER THAT ', { bold: true }),
            t('all acts, deeds, and things done by the said Partner(s) in this regard shall be binding upon the Firm and all its partners and shall remain in force. '),
            t('RESOLVED FURTHER THAT ', { bold: true }),
            t('this resolution shall remain in force until a written notice of its withdrawal or amendment is served upon and acknowledged by PhonePe Limited, and that a certified true copy of this resolution be furnished to PhonePe Limited for their records'),
          ],
          spacing: { after: 240 }
        }),
        p([]),

        sigTable(presentPartners),
        p([]),
        p([t('Seal of the Firm')], { alignment: AlignmentType.CENTER }),

        // ── PAGE 2: DECLARATION ─────────────────────────────
        new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
        p([]),

        new Paragraph({
          children: [new TextRun({ text: 'Declaration', font: 'Times New Roman', size: 24, bold: true, underline: { type: UnderlineType.SINGLE } })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }),

        new Paragraph({
          children: [t('I/we, the undersigned individuals, hereby personally, jointly, and severally undertake and declare that:')],
          spacing: { after: 160 }
        }),

        new Paragraph({
          children: [t('1. Our firm ['), t(firmName, { bold: true }), t('] is constituted as a partnership firm and it is')],
          spacing: { after: 100 }
        }),

        // Registered checkbox
        p([t(isRegistered ? '\u2611 Registered' : '\u2610 Registered')], { indent: { left: 720 } }),
        new Paragraph({
          children: [t('Whether there is/are any natural persons(s), acting alone or through other juridical persons, with more than 10% ownership/entitlement to capital/profits, or who exercise control through other means.')],
          indent: { left: 1440 }, spacing: { after: 80 }
        }),
        ...(isRegistered ? [
          p([t(naturalPersons === 'yes' ? '\u2611 Yes' : '\u2610 Yes')], { indent: { left: 1440 } }),
          p([t(naturalPersons === 'no'  ? '\u2611 No'  : '\u2610 No')],  { indent: { left: 1440 } }),
        ] : [
          p([t('\u2610 Yes'), t('   '), t('\u2610 No')], { indent: { left: 1440 } }),
        ]),

        // Unregistered checkbox
        p([t(isRegistered ? '\u2610 Unregistered' : '\u2611 Unregistered')], { indent: { left: 720 } }),
        new Paragraph({
          children: [t('Whether there is/are any natural persons(s), acting alone or through other juridical persons, with more than 15% ownership/entitlement to capital/profits, or who exercise control through other means.')],
          indent: { left: 1440 }, spacing: { after: 80 }
        }),
        ...(!isRegistered ? [
          p([t(naturalPersons === 'yes' ? '\u2611 Yes' : '\u2610 Yes')], { indent: { left: 1440 } }),
          p([t(naturalPersons === 'no'  ? '\u2611 No'  : '\u2610 No')],  { indent: { left: 1440 } }),
        ] : [
          p([t('\u2610 Yes'), t('   '), t('\u2610 No')], { indent: { left: 1440 } }),
        ]),

        p([]),
        new Paragraph({
          children: [t('2. If there is/are no natural person(s) as per the responses provided under declaration (1) hereinabove, then we declare that the authorised signatory identified in the Resolution (defined below) is the senior managing official and be considered as the beneficial owner of our firm.')],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [t('3. ' + (pepDeclaration ? '\u2611' : '\u2610') + ' No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a '), t('Politically Exposed Person (PEP) (as defined by RBI).', { bold: true })],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [t('4. The contents of the resolution of the partners dated ['), t(resolutionDate, { bold: true }), t('] (\u201cResolution\u201d) are true and valid.')],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [t('5. The list of partners constituting the partnership firm and their respective details, as provided in the partnership deed dated ['), t(deedDate, { bold: true }), t('] (\u201cPartnership Deed\u201d), are true, complete, current, and valid.')],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [t('6. The percentage of ownership and/or entitlement to capital or profits of the partners as specified in the Partnership Deed is true, complete, current, and valid.')],
          spacing: { after: 200 }
        }),

        new Paragraph({
          children: [t('In consideration of PhonePe agreeing to rely on the Resolution, the Partnership Deed, and this declaration, I/we hereby personally, jointly, and severally undertake to indemnify and hold PhonePe harmless against all damages, liabilities, claims, demands, actions, proceedings, losses, costs (including legal costs), expenses, and all other liabilities of whatsoever nature or description, arising out of or in connection with PhonePe\u2019s reliance on the said Resolution, Partnership Deed, and this declaration.')],
          spacing: { after: 240 }
        }),
        p([]),

        sigTable(presentPartners),
      ]
    }]
  });

  return Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────
// BUILD BO DECLARATION
// ─────────────────────────────────────────────────────────
async function buildBO(d) {
  const { firmName, regAddress, principalSame, principalAddress, isRegistered,
          boDate, deedDate, resolutionDate, boCategory,
          pepDeclaration, boAuthName, partners } = d;

  const principalCell = principalSame
    ? '\u2611 Same   OR ___________________________[ADDRESS]'
    : '\u2610 Same   OR ' + principalAddress;

  const entityTypeCell = isRegistered
    ? '\u2611 Partnership Firm / LLP\n\u2610 Unregistered Partnership Firm'
    : '\u2610 Partnership Firm / LLP\n\u2611 Unregistered Partnership Firm';

  const entityRows = [
    ['I',   'Name of the entity',                  firmName],
    ['II',  'Registered address',                  regAddress],
    ['III', 'Principal place of operation/office', principalCell],
    ['IV',  'Type of entity',                      entityTypeCell],
  ];

  // BO table col widths — total 10080
  const boCols = [400, 1100, 1800, 680, 680, 980, 980, 680, 780];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        // ── PAGE 1 ────────────────────────────────────────
        p([t('(To be printed on Partnership Firm\u2019s Letterhead)', { italics: true })], { alignment: AlignmentType.CENTER }),
        p([]),

        new Paragraph({
          children: [new TextRun({ text: 'DECLARATION OF BENEFICIAL OWNERSHIP (BO) and LIST OF PARTNERS', font: 'Times New Roman', size: 26, bold: true, underline: { type: UnderlineType.SINGLE } })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 }
        }),
        new Paragraph({
          children: [t('(Not applicable for Individual/ Sole Proprietor/ HUF/ Government Departments / Public Sector Undertaking / Local Authority (Municipal Corporations, Gram Panchayats, etc.))', { italics: true, size: 18 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }),

        // Entity info table
        new Table({
          width: { size: 10080, type: WidthType.DXA },
          columnWidths: [500, 2500, 7080],
          rows: entityRows.map(([num, label, val]) => new TableRow({
            children: [
              cell(num,   500),
              cell(label, 2500),
              cell(
                val.split('\n').map(line => new Paragraph({ children: [t(line)], spacing: { after: 0 } })),
                7080
              ),
            ]
          }))
        }),
        p([]),

        new Paragraph({
          children: [
            t('The Legal Entity as stated above hereby confirms and declares the following on the below date: ', { bold: true }),
            t(boDate, { bold: true })
          ],
          spacing: { after: 100 }
        }),
        new Paragraph({
          children: [t('Tick the box as applicable ((Refer note A \u201cRBI guidelines for identification of Beneficial owners\u201d)', { italics: true })],
          spacing: { after: 100 }
        }),

        new Paragraph({
          children: [
            t((boCategory === 'cat1' ? '\u2611' : '\u2610') + ' Category 1', { bold: true }),
            t(' - We hereby declare that following persons/ entity noted in the below table own 10%/15%# or more interest or possess the right to control management/policy decisions (Refer Notes).')
          ],
          spacing: { after: 80 }
        }),
        new Paragraph({
          children: [
            t((boCategory === 'cat2' ? '\u2611' : '\u2610') + ' Category 2', { bold: true }),
            t(' - We hereby declare that no natural person is identified as per category 1 (above). (Mention the details of the natural person(s) holding the position of senior management official in the entity.)')
          ],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [
            t((pepDeclaration ? '\u2611' : '\u2610') + ' No personnel, director, officer, any family member or close associate of the Merchant and its beneficial owners, is a '),
            t('Politically Exposed Person (PEP) (as defined by RBI).', { bold: true })
          ],
          spacing: { after: 160 }
        }),

        p([t('The details of beneficial owner(s) is/are as follows:')]),

        // BO data table
        new Table({
          width: { size: 10080, type: WidthType.DXA },
          columnWidths: boCols,
          rows: [
            new TableRow({
              tableHeader: true,
              children: [
                hCell('S.N.',          boCols[0]),
                hCell('Name',          boCols[1]),
                hCell('Residential Address and PIN code', boCols[2]),
                hCell('Designation',   boCols[3]),
                hCell('DOB',           boCols[4]),
                hCell('Proof of identity (Refer Note C)', boCols[5]),
                hCell('Proof of Address (Refer Note C)',  boCols[6]),
                hCell('Nationality',   boCols[7]),
                hCell('% of interest/ ownership in the entity', boCols[8]),
              ]
            }),
            ...partners.map((partner, i) => new TableRow({
              children: [
                cell(String(i + 1),           boCols[0]),
                cell(partner.name,             boCols[1]),
                cell(partner.address,          boCols[2]),
                cell(partner.designation,      boCols[3]),
                cell(partner.dob,              boCols[4]),
                cell(partner.pan,              boCols[5]),
                cell(partner.poa + ' ' + partner.poaNum, boCols[6]),
                cell(partner.nationality,      boCols[7]),
                cell(partner.share + '%',      boCols[8]),
              ]
            })),
            // spare blank row
            new TableRow({
              children: boCols.map(w => cell(' ', w))
            }),
          ]
        }),
        p([]),

        p([t('List of the Current Partners of the firm operating at the aforementioned address:')]),
        ...partners.map((partner, i) => p([t((i + 1) + '. ' + partner.name)])),
        p([]),
        p([t('We acknowledge and confirm that the information provided above is true and correct to the best of our knowledge and belief.')]),
        p([]),

        p([t('Authorised Signatory/ies: (Refer note B for signature requirement)', { bold: true })]),
        p([t('___________________________(Name, Signature with Stamp)')]),
        p([t(boAuthName)]),
        p([]),

        // ── PAGE 2: NOTES ────────────────────────────────
        new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),
        p([]),
        p([t('#Notes:-', { bold: true })]),
        p([t('A.  RBI guidelines for identification of Beneficial owners', { bold: true })]),
        p([t('Category 1: Controlling ownership interest means:', { bold: true })]),

        new Table({
          width: { size: 5000, type: WidthType.DXA },
          columnWidths: [2500, 2500],
          rows: [
            new TableRow({ children: [hCell('Business entity', 2500), hCell('Shareholding* %', 2500)] }),
            new TableRow({ children: [cell('Partnership Firm', 2500), cell('>10%', 2500)] }),
            new TableRow({ children: [cell('Unregistered Partnership Firm', 2500), cell('>15%', 2500)] }),
          ]
        }),
        p([]),

        new Paragraph({
          children: [
            t('i.      Ownership of/entitlement to '),
            t('more than 10%/15%', { bold: true }),
            t(' of the capital or profits of the juridical person where the juridical person is a '),
            t('partnership firm, LLP/ unregistered firm', { bold: true }),
            t('[ \u2018Control\u2019 shall include the right to control the management or policy decision.]', { bold: true }),
          ],
          indent: { left: 360 }, spacing: { after: 120 }
        }),
        p([t('Category 2:', { bold: true })]),
        p([t('Where no natural person is identified under (i) or (ii) or (iii) of category 1, the beneficial owner is the relevant natural person who holds the position of senior managing official in that entity.')]),
        p([]),
        p([t('B. Signature on the Declaration form:', { bold: true })]),
        p([t('A person who is authorised to sign BO declaration : '), t('Authorised signatory', { bold: true })]),
        p([]),
        p([t('C. Other Instructions', { bold: true })]),
        p([t('1.   Proof of Identity -')]),

        new Table({
          width: { size: 7000, type: WidthType.DXA },
          columnWidths: [3500, 3500],
          rows: [
            new TableRow({ children: [hCell('BO Type', 3500), hCell('Details Required', 3500)] }),
            new TableRow({ children: [cell('Individual (Indian / Foreign National) / Indian Entity', 3500), cell('PAN* (if not available, then form 60 should be provided)', 3500)] }),
            new TableRow({ children: [cell('Foreign entity', 3500), cell('Valid Establishment document issued in the country of incorporation/registration', 3500)] }),
          ]
        }),
        p([]),
        p([t('2.   Proof of Address -')]),
        new Table({
          width: { size: 7000, type: WidthType.DXA },
          columnWidths: [3500, 3500],
          rows: [
            new TableRow({ children: [hCell('BO Type', 3500), hCell('Details Required', 3500)] }),
            new TableRow({ children: [cell('Individual (Indian / Foreign National)', 3500), cell('Voter ID/ Driving License / Passport/ Redacted Aadhar', 3500)] }),
            new TableRow({ children: [cell('Entity (Indian or Foreign)', 3500), cell('Valid Establishment document', 3500)] }),
          ]
        }),
        p([]),
        p([t('3.   PAN Number to be provided for Residents/ Entities registered in India.')]),
        p([t('4.   In case if BO is a minor, and POI or POA as mentioned above is not available, then valid age proof to be provided.')]),
      ]
    }]
  });

  return Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────
// API ROUTES
// ─────────────────────────────────────────────────────────
app.post('/api/generate/resolution', async (req, res) => {
  try {
    const buffer = await buildResolution(req.body);
    const filename = (req.body.firmName || 'firm').replace(/[^a-zA-Z0-9]/g, '_') + '_Partner_Resolution.docx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Resolution error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/generate/bo', async (req, res) => {
  try {
    const buffer = await buildBO(req.body);
    const filename = (req.body.firmName || 'firm').replace(/[^a-zA-Z0-9]/g, '_') + '_BO_Declaration.docx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('BO error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────────────────
// START
// ─────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅  Merchant Tool running at http://localhost:${PORT}`));
