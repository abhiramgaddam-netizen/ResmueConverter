import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import {
  Document,
  Paragraph,
  TextRun,
  AlignmentType,
  Packer,
  Header,
  Footer,
} from 'docx';
import { ResumeData } from '@/types/resume';

function escapeXml(unsafe: string | null | undefined): string {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildSectionHeadingXml(title: string): string {
  return `<w:p><w:pPr><w:spacing w:before="240" w:after="80"/><w:pBdr><w:top w:val="none"/><w:left w:val="none"/><w:right w:val="none"/><w:bottom w:val="single" w:sz="6" w:space="2" w:color="0284C7"/></w:pBdr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0284C7"/><w:sz w:val="23"/></w:rPr><w:t xml:space="preserve">${escapeXml(title.toUpperCase())}</w:t></w:r></w:p>`;
}

function buildBodyParagraphXml(text: string, boldPrefix?: string): string {
  const boldPart = boldPrefix
    ? `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(boldPrefix)} </w:t></w:r>`
    : '';

  return `<w:p><w:pPr><w:spacing w:before="30" w:after="50"/></w:pPr>${boldPart}<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="334155"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function buildBulletParagraphXml(text: string): string {
  return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:before="20" w:after="30"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="0284C7"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">•  </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="334155"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function buildExperienceHeaderXml(role: string, company: string, dates: string, location?: string): string {
  const meta = [company, location, dates].filter(Boolean).join('  |  ');
  return `<w:p><w:pPr><w:spacing w:before="120" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${escapeXml(role)}</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="0" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:i/><w:color w:val="0284C7"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(meta)}</w:t></w:r></w:p>`;
}

export function buildContactInfoXml(
  data: ResumeData,
  removePersonalInfo?: boolean
): string {
  if (removePersonalInfo) return '';
  const contactItems: string[] = [];
  if (data.contactInfo?.email) contactItems.push(data.contactInfo.email);
  if (data.contactInfo?.phone) contactItems.push(data.contactInfo.phone);
  if (data.contactInfo?.location) contactItems.push(data.contactInfo.location);
  if (data.contactInfo?.linkedin) contactItems.push(data.contactInfo.linkedin);
  if (data.contactInfo?.github) contactItems.push(data.contactInfo.github);
  if (contactItems.length === 0) return '';
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="140"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="64748B"/><w:sz w:val="19"/></w:rPr><w:t xml:space="preserve">${escapeXml(contactItems.join('   |   '))}</w:t></w:r></w:p>`;
}

export function buildSummarySectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.professionalSummary || !data.professionalSummary.trim()) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Professional Summary'));
  parts.push(buildBodyParagraphXml(data.professionalSummary.trim()));
  return parts.join('');
}

export function buildSkillsSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.skills || data.skills.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Technical Skills & Competencies'));
  for (const cat of data.skills) {
    if (typeof cat === 'string') {
      parts.push(buildBulletParagraphXml(cat));
    } else if (cat && Array.isArray(cat.skills) && cat.skills.length > 0) {
      parts.push(buildBodyParagraphXml(cat.skills.join(', '), `${cat.category || 'Core'}:`));
    }
  }
  return parts.join('');
}

export function buildExperienceSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.experience || data.experience.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Professional Experience'));
  for (const exp of data.experience) {
    const dates = [exp.startDate, exp.endDate].filter(Boolean).join(' - ');
    parts.push(buildExperienceHeaderXml(exp.role || 'Consultant', exp.company || '', dates, exp.location));
    if (exp.bullets && Array.isArray(exp.bullets) && exp.bullets.length > 0) {
      for (const bullet of exp.bullets) {
        parts.push(buildBulletParagraphXml(bullet));
      }
    }
  }
  return parts.join('');
}

export function buildProjectsSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.projects || data.projects.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Key Projects & Engagements'));
  for (const proj of data.projects) {
    parts.push(
      `<w:p><w:pPr><w:spacing w:before="80" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(proj.name)}</w:t></w:r>${proj.technologies ? `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:i/><w:color w:val="64748B"/><w:sz w:val="19"/></w:rPr><w:t xml:space="preserve"> (${escapeXml(proj.technologies)})</w:t></w:r>` : ''}</w:p>`
    );
    if (proj.description) {
      parts.push(buildBodyParagraphXml(proj.description));
    }
    if (proj.bullets && Array.isArray(proj.bullets) && proj.bullets.length > 0) {
      for (const b of proj.bullets) {
        parts.push(buildBulletParagraphXml(b));
      }
    }
  }
  return parts.join('');
}

export function buildEducationSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.education || data.education.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Education & Credentials'));
  for (const edu of data.education) {
    const dates = [edu.startDate, edu.endDate].filter(Boolean).join(' - ');
    const degree = [edu.degree, edu.fieldOfStudy].filter(Boolean).join(' in ');
    parts.push(
      `<w:p><w:pPr><w:spacing w:before="60" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0F172A"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(edu.institution || 'University')}</w:t></w:r>${dates ? `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="64748B"/><w:sz w:val="19"/></w:rPr><w:t xml:space="preserve">  (${escapeXml(dates)})</w:t></w:r>` : ''}</w:p>`
    );
    if (degree) {
      parts.push(buildBodyParagraphXml(degree));
    }
  }
  return parts.join('');
}

export function buildCertificationsSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.certifications || data.certifications.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Professional Certifications'));
  for (const cert of data.certifications) {
    parts.push(buildBulletParagraphXml(cert));
  }
  return parts.join('');
}

/**
 * Checks whether the paragraph immediately preceding a placeholder already serves as a heading
 * with relevant keywords, preventing redundant default headings.
 */
function checkPrecedingHeading(docXml: string, placeholderMatchIndex: number, keywords: string[]): boolean {
  const xmlBefore = docXml.substring(0, placeholderMatchIndex);
  const lastPEnd = xmlBefore.lastIndexOf('</w:p>');
  if (lastPEnd === -1) return false;
  const lastPStart = xmlBefore.lastIndexOf('<w:p', lastPEnd);
  if (lastPStart === -1) return false;
  const lastPXml = xmlBefore.substring(lastPStart, lastPEnd + 6);
  const textMatch = lastPXml.replace(/<[^>]+>/g, '').toLowerCase().trim();
  return keywords.some((k) => textMatch.includes(k.toLowerCase()));
}

/**
 * Safely normalizes split <w:t> runs inside paragraphs containing {{ ... }}
 * while preserving all <w:pPr> (borders/lines), <w:drawing> (images/logos), <w:pict>, and <w:tab/> elements.
 */
function normalizePlaceholdersInXml(xml: string): string {
  return xml.replace(/(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/gi, (fullP, pOpen, pInner, pClose) => {
    if (!pInner.includes('{')) return fullP;

    // If an intact placeholder already exists in the inner XML, no need to touch the runs
    if (/\{\{\s*[\w\s.-]+\s*\}\}/.test(pInner)) {
      return fullP;
    }

    // Check if combined <w:t> text contains a placeholder
    const matches = pInner.match(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/gi) || [];
    const fullText = (matches as string[])
      .map((t) => t.replace(/<[^>]+>/g, ''))
      .join('');

    if (!/\{\{\s*[\w\s.-]+\s*\}\}/.test(fullText)) {
      return fullP;
    }

    // Placeholder IS split across runs.
    // Non-destructively merge adjacent text runs where the first run has an unclosed '{'
    let normalizedInner = pInner;
    const mergeRunRegex = /(<w:t\b[^>]*>(?:(?!<\/w:t>)[\s\S])*?\{[^{}]*?)\s*<\/w:t>\s*<\/w:r>\s*(?:<w:proofErr[^>]*>\s*)*<w:r\b[^>]*>(?:\s*<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:t\b[^>]*>/gi;

    let prev: string;
    do {
      prev = normalizedInner;
      normalizedInner = normalizedInner.replace(mergeRunRegex, '$1');
    } while (normalizedInner !== prev);

    return `${pOpen}${normalizedInner}${pClose}`;
  });
}

const SECTION_SPECS = [
  {
    id: 'contact_info',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(contact_info|contact|contacts)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['contact', 'address', 'communication'],
    buildXml: (data: ResumeData, _incHeading: boolean, opts?: { removePersonalInfo?: boolean }) =>
      buildContactInfoXml(data, opts?.removePersonalInfo),
  },
  {
    id: 'summary',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(summary|professional_summary|profile|executive_summary|career_summary|about_me)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['summary', 'profile', 'overview', 'about', 'career', 'executive'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildSummarySectionXml(data, incHeading),
  },
  {
    id: 'skills',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(skills|technical_skills|competencies|core_competencies|skill_matrix|expertise)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['skill', 'competenc', 'expertise', 'technolog'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildSkillsSectionXml(data, incHeading),
  },
  {
    id: 'experience',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(experience|professional_experience|work_experience|employment|work_history)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['experience', 'employment', 'history', 'career', 'work'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildExperienceSectionXml(data, incHeading),
  },
  {
    id: 'projects',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(projects|key_projects|project_experience|engagements)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['project', 'engagement', 'initiative'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildProjectsSectionXml(data, incHeading),
  },
  {
    id: 'education',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(education|credentials|academics|qualifications)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['education', 'academic', 'qualification', 'degree', 'credential'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildEducationSectionXml(data, incHeading),
  },
  {
    id: 'certifications',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(certifications|certificates|licenses)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['certif', 'licens', 'accredit'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildCertificationsSectionXml(data, incHeading),
  },
];

export function detectTemplateHeaderType(templateBuffer?: Buffer | null): 'role' | 'name' {
  let docxBuffer = templateBuffer;
  if (!docxBuffer) {
    const defaultPath = path.join(process.cwd(), 'src', 'templates', 'tecnics_resume_template.docx');
    if (fs.existsSync(defaultPath)) {
      docxBuffer = fs.readFileSync(defaultPath);
    }
  }

  if (docxBuffer) {
    try {
      const zip = new PizZip(docxBuffer.toString('binary'));
      const headerFiles = Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f));
      for (const hPath of headerFiles) {
        const hXml = zip.file(hPath)?.asText() || '';
        if (/\{\{\s*(role|target_role|candidate_role|job_title|title|designation)\s*\}\}/i.test(hXml)) {
          return 'role';
        }
      }
    } catch {
      // ignore
    }
  }
  return 'name';
}

export async function generateDocxFromTemplate(
  data: ResumeData,
  templateBuffer?: Buffer | null,
  options?: { removePersonalInfo?: boolean }
): Promise<Buffer> {
  let docxBuffer = templateBuffer;

  // Load default template if no custom buffer provided
  if (!docxBuffer) {
    const defaultPath = path.join(process.cwd(), 'src', 'templates', 'tecnics_resume_template.docx');
    if (fs.existsSync(defaultPath)) {
      docxBuffer = fs.readFileSync(defaultPath);
    }
  }

  if (docxBuffer) {
    try {
      const zip = new PizZip(docxBuffer.toString('binary'));
      const candidateName = data.fullName?.trim() || 'Candidate Profile';
      const targetRole = data.targetRole?.trim() || 'Technology Consultant';
      const removePersonal = Boolean(options?.removePersonalInfo);

      const ROLE_REGEX = /\{\{\s*(role|target_role|candidate_role|job_title|title|designation)\s*\}\}/gi;
      const NAME_REGEX = /\{\{\s*(candidate_name|name|full_name|fullname|candidate)\s*\}\}/gi;

      // 1. Process header XML files (e.g. header1.xml)
      const headerFiles = Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f));
      for (const hPath of headerFiles) {
        const file = zip.file(hPath);
        if (file) {
          let hXml = normalizePlaceholdersInXml(file.asText());
          // Dynamically replace {{ Role }} or {{ candidate_name }} in header
          hXml = hXml.replace(ROLE_REGEX, () => escapeXml(targetRole.toUpperCase()));
          hXml = hXml.replace(NAME_REGEX, () => escapeXml(candidateName.toUpperCase()));
          zip.file(hPath, hXml);
        }
      }

      // Process footer XML files (e.g. footer1.xml)
      const footerFiles = Object.keys(zip.files).filter((f) => /^word\/footer\d+\.xml$/.test(f));
      for (const fPath of footerFiles) {
        const file = zip.file(fPath);
        if (file) {
          let fXml = normalizePlaceholdersInXml(file.asText());
          fXml = fXml.replace(ROLE_REGEX, () => escapeXml(targetRole));
          fXml = fXml.replace(NAME_REGEX, () => escapeXml(candidateName));
          zip.file(fPath, fXml);
        }
      }

      // 2. Process document.xml
      const docFile = zip.file('word/document.xml');
      if (docFile) {
        let docXml = normalizePlaceholdersInXml(docFile.asText());

        // Check if {{ Role }} is explicitly present in document body
        const hasRoleInBody = /\{\{\s*(role|target_role|candidate_role|job_title|title|designation)\s*\}\}/i.test(docXml);

        // Replace {{ candidate_name }} and {{ Role }} placeholders in document body
        docXml = docXml.replace(NAME_REGEX, () => escapeXml(candidateName));
        if (hasRoleInBody) {
          docXml = docXml.replace(ROLE_REGEX, () => escapeXml(targetRole));
        }

        // Inline contact placeholders
        const emailVal = removePersonal ? '' : (data.contactInfo?.email || '');
        const phoneVal = removePersonal ? '' : (data.contactInfo?.phone || '');
        const locVal = removePersonal ? '' : (data.contactInfo?.location || '');
        const linkedinVal = removePersonal ? '' : (data.contactInfo?.linkedin || '');
        const githubVal = removePersonal ? '' : (data.contactInfo?.github || '');

        docXml = docXml.replace(/\{\{\s*email\s*\}\}/gi, () => escapeXml(emailVal));
        docXml = docXml.replace(/\{\{\s*(phone|mobile|telephone|cell)\s*\}\}/gi, () => escapeXml(phoneVal));
        docXml = docXml.replace(/\{\{\s*(location|address|city)\s*\}\}/gi, () => escapeXml(locVal));
        docXml = docXml.replace(/\{\{\s*linkedin\s*\}\}/gi, () => escapeXml(linkedinVal));
        docXml = docXml.replace(/\{\{\s*github\s*\}\}/gi, () => escapeXml(githubVal));

        // 3. Process individual section placeholders
        const claimedSections = new Set<string>();
        for (const spec of SECTION_SPECS) {
          if (spec.regex.test(docXml)) {
            claimedSections.add(spec.id);
            docXml = docXml.replace(spec.regex, (_fullMatch, _p1, matchOffset) => {
              const hasCustomHeading = checkPrecedingHeading(docXml, matchOffset, spec.keywords);
              const sectionXml = spec.buildXml(data, !hasCustomHeading, options);
              return sectionXml || '';
            });
          }
        }

        // 4. Content / Catch-all placeholder (or document append if unplaced)
        const CONTENT_REGEX = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(content|body|resume_content|all_sections)\s*\}\}[\s\S]*?<\/w:p>/gi;
        const hasContentPlaceholder = CONTENT_REGEX.test(docXml);

        // Build remaining sections for content placeholder or end of doc
        const remainingParts: string[] = [];

        if (!claimedSections.has('contact_info')) {
          const contactBar: string[] = [];
          if (!removePersonal) {
            if (data.contactInfo?.email) contactBar.push(data.contactInfo.email);
            if (data.contactInfo?.phone) contactBar.push(data.contactInfo.phone);
            if (data.contactInfo?.location) contactBar.push(data.contactInfo.location);
            if (data.contactInfo?.linkedin) contactBar.push(data.contactInfo.linkedin);
            if (data.contactInfo?.github) contactBar.push(data.contactInfo.github);
          }
          const includeRole = !hasRoleInBody && Boolean(data.targetRole);
          if (includeRole || contactBar.length > 0) {
            if (includeRole) {
              remainingParts.push(
                `<w:p><w:pPr><w:spacing w:before="40" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0284C7"/><w:sz w:val="23"/></w:rPr><w:t xml:space="preserve">${escapeXml(data.targetRole)}</w:t></w:r></w:p>`
              );
            }
            if (contactBar.length > 0) {
              remainingParts.push(
                `<w:p><w:pPr><w:spacing w:before="0" w:after="140"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="64748B"/><w:sz w:val="19"/></w:rPr><w:t xml:space="preserve">${escapeXml(contactBar.join('   |   '))}</w:t></w:r></w:p>`
              );
            }
          }
        }

        if (!claimedSections.has('summary')) {
          remainingParts.push(buildSummarySectionXml(data, true));
        }
        if (!claimedSections.has('skills')) {
          remainingParts.push(buildSkillsSectionXml(data, true));
        }
        if (!claimedSections.has('experience')) {
          remainingParts.push(buildExperienceSectionXml(data, true));
        }
        if (!claimedSections.has('projects')) {
          remainingParts.push(buildProjectsSectionXml(data, true));
        }
        if (!claimedSections.has('education')) {
          remainingParts.push(buildEducationSectionXml(data, true));
        }
        if (!claimedSections.has('certifications')) {
          remainingParts.push(buildCertificationsSectionXml(data, true));
        }

        const remainingXml = remainingParts.join('');

        if (hasContentPlaceholder) {
          docXml = docXml.replace(CONTENT_REGEX, () => remainingXml);
        } else if (claimedSections.size === 0) {
          // If no placeholders were used at all, append to w:body
          docXml = docXml.replace(/<\/w:body>/, () => `${remainingXml}</w:body>`);
        }

        zip.file('word/document.xml', docXml);
      }

      return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    } catch (err) {
      console.warn('Template XML injection encountered an issue, falling back to clean generator:', err);
    }
  }

  return await generateCleanFallbackDocx(data);
}

async function generateCleanFallbackDocx(data: ResumeData): Promise<Buffer> {
  const children: Paragraph[] = [];

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({
          text: data.fullName || 'Candidate Profile',
          bold: true,
          size: 32,
          color: '0F172A',
          font: 'Calibri',
        }),
      ],
    })
  );

  if (data.targetRole) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({
            text: data.targetRole,
            bold: true,
            size: 22,
            color: '0284C7',
            font: 'Calibri',
          }),
        ],
      })
    );
  }

  if (data.professionalSummary) {
    children.push(
      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: 'PROFESSIONAL SUMMARY',
            bold: true,
            size: 22,
            color: '0284C7',
            font: 'Calibri',
          }),
        ],
      })
    );
    children.push(
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [
          new TextRun({
            text: data.professionalSummary,
            size: 21,
            color: '334155',
            font: 'Calibri',
          }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: 'TECNICS INTEGRATION TECHNOLOGIES',
                    size: 15,
                    color: '94A3B8',
                    font: 'Calibri',
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: 'Tecnics Integration Technologies Private Limited  •  Standardized Profile',
                    size: 15,
                    color: '94A3B8',
                    font: 'Calibri',
                  }),
                ],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}

export async function generateDocxBuffer(
  data: ResumeData,
  templateBuffer?: Buffer | null,
  options?: { removePersonalInfo?: boolean }
): Promise<Buffer> {
  return await generateDocxFromTemplate(data, templateBuffer, options);
}
