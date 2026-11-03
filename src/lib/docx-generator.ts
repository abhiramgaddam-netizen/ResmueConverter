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
import { formatCandidateFirstName } from './name-utils';

export { formatCandidateFirstName };

function escapeXml(unsafe: string | null | undefined): string {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Main Section Heading: Times New Roman, 13pt (sz="26"), Bold, Black
 * e.g. "Professional Summary:", "Education:", "Technical Skills:"
 */
export function buildSectionHeadingXml(title: string): string {
  const cleanTitle = title.endsWith(':') ? title : `${title}:`;
  return `<w:p><w:pPr><w:spacing w:before="160" w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t xml:space="preserve">${escapeXml(cleanTitle)}</w:t></w:r></w:p>`;
}

/**
 * Subheading: Times New Roman, 11pt (sz="22"), Bold, Black
 * e.g. "Migration Projects", "SKILLS@SAP R/3", "Project 1", "Responsibilities:"
 */
export function buildSubHeadingXml(title: string): string {
  return `<w:p><w:pPr><w:spacing w:before="120" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">${escapeXml(title)}</w:t></w:r></w:p>`;
}

/**
 * Bullet Paragraph: Times New Roman, 11pt (sz="22"), indented bullet
 */
export function buildBulletParagraphXml(text: string): string {
  return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:ind w:left="360" w:hanging="180"/><w:spacing w:before="20" w:after="30"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">•\t${escapeXml(text)}</w:t></w:r></w:p>`;
}

/**
 * Tab-Aligned Key-Value Line: Times New Roman, 11pt (sz="22")
 * e.g. "Client:\t\t\tGlobal Support Management Centre"
 */
export function buildKeyValueTabXml(label: string, value: string, tabCount = 2, boldLabel = true): string {
  if (!value || !value.trim()) return '';
  const cleanLabel = label.endsWith(':') ? label : `${label}:`;
  const tabsXml = Array(tabCount).fill('<w:tab/>').join('');
  return `<w:p><w:pPr><w:spacing w:before="20" w:after="30"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${boldLabel ? '<w:b/>' : ''}<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">${escapeXml(cleanLabel)}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>${tabsXml}<w:t xml:space="preserve">${escapeXml(value.trim())}</w:t></w:r></w:p>`;
}

/**
 * Contact Info Line: Times New Roman, 10.5pt (sz="21")
 */
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
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="475569"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(contactItems.join('   |   '))}</w:t></w:r></w:p>`;
}

// 1. Professional Summary + Specialized Skill Sections (Migration Projects, SKILLS@SAP R/3, etc.)
export function buildSummarySectionXml(data: ResumeData, includeHeading = true): string {
  const parts: string[] = [];
  const summaryItems = Array.isArray(data.professionalSummary)
    ? data.professionalSummary
    : typeof data.professionalSummary === 'string' && data.professionalSummary.trim()
    ? data.professionalSummary.split('\n').filter(Boolean)
    : [];

  if (summaryItems.length > 0) {
    if (includeHeading) parts.push(buildSectionHeadingXml('Professional Summary:'));
    for (const item of summaryItems) {
      parts.push(buildBulletParagraphXml(item));
    }
  }

  // Render specialized sections ("Migration Projects", "SKILLS@SAP R/3", "Net Weaver Skills", "Hana Skills")
  if (data.specializedSkills && data.specializedSkills.length > 0) {
    for (const group of data.specializedSkills) {
      if (group.title && group.bullets?.length > 0) {
        parts.push(buildSubHeadingXml(group.title));
        for (const b of group.bullets) {
          parts.push(buildBulletParagraphXml(b));
        }
      }
    }
  }

  return parts.join('');
}

// 2. Education
export function buildEducationSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.education || data.education.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Education:'));
  for (const edu of data.education) {
    const line = [
      edu.degree,
      edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
      edu.institution ? `from ${edu.institution}` : '',
      edu.endDate ? `(${edu.endDate})` : '',
    ]
      .filter(Boolean)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (line) parts.push(buildBulletParagraphXml(line));
  }
  return parts.join('');
}

// 3. Technical Skills (Operating Systems, SAP Software, Database, etc.)
export function buildSkillsSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.skills || data.skills.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Technical Skills:'));
  for (const cat of data.skills) {
    if (cat && Array.isArray(cat.skills) && cat.skills.length > 0) {
      const label = cat.category || 'Technical Skills';
      const tabs = label.length <= 12 ? 2 : 1;
      parts.push(buildKeyValueTabXml(label, cat.skills.join(', '), tabs, false));
    }
  }
  return parts.join('');
}

// 4. Awards and Achievements
export function buildAwardsSectionXml(data: ResumeData, includeHeading = true): string {
  const awards = data.awards || [];
  if (awards.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Awards and Achievements:'));
  for (const award of awards) {
    parts.push(buildBulletParagraphXml(award));
  }
  return parts.join('');
}

// 5. Professional Experience (Project 1 -> Client -> Role -> Duration -> Environment -> Working Environment -> Responsibilities)
export function buildExperienceSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.experience || data.experience.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Professional Experience:'));

  data.experience.forEach((exp, idx) => {
    const projTitle = exp.projectNumber || `Project ${idx + 1}`;
    parts.push(buildSubHeadingXml(exp.projectName ? `${projTitle}: ${exp.projectName}` : projTitle));

    const duration = [exp.startDate, exp.endDate].filter(Boolean).join(' to ');
    if (exp.company) parts.push(buildKeyValueTabXml('Client:', exp.company, 3, true));
    if (exp.role) parts.push(buildKeyValueTabXml('Role:', exp.role, 3, true));
    if (duration) parts.push(buildKeyValueTabXml('Duration:', duration, 2, true));
    if (exp.environment) parts.push(buildKeyValueTabXml('Environment:', exp.environment, 2, true));
    if (exp.domain) parts.push(buildKeyValueTabXml('Domain:', exp.domain, 2, true));
    if (exp.onsiteOffshore) parts.push(buildKeyValueTabXml('Location:', exp.onsiteOffshore, 2, true));
    if (exp.description) parts.push(buildBulletParagraphXml(`Project Description: ${exp.description}`));

    // Conditional Working Environment section
    if (exp.workingEnvironment && exp.workingEnvironment.length > 0) {
      parts.push(buildSubHeadingXml('Working Environment:'));
      for (const we of exp.workingEnvironment) {
        parts.push(buildBulletParagraphXml(we));
      }
    }

    // Responsibilities section
    if (exp.bullets && exp.bullets.length > 0) {
      parts.push(buildSubHeadingXml('Responsibilities:'));
      for (const b of exp.bullets) {
        parts.push(buildBulletParagraphXml(b))
      }
    }
  });

  return parts.join('');
}

// 6. Professional Certifications
export function buildCertificationsSectionXml(data: ResumeData, includeHeading = true): string {
  if (!data.certifications || data.certifications.length === 0) return '';
  const parts: string[] = [];
  if (includeHeading) parts.push(buildSectionHeadingXml('Professional Certifications:'));
  for (const cert of data.certifications) {
    parts.push(buildBulletParagraphXml(cert));
  }
  return parts.join('');
}

function buildAdditionalSectionsXml(data: ResumeData): string {
  return (data.additionalSections || []).filter((section) => section.title && section.items?.length)
    .map((section) => buildSectionHeadingXml(section.title) + section.items.map(buildBulletParagraphXml).join('')).join('');
}

function appendBeforeSectPr(docXml: string, newXml: string): string {
  if (!newXml.trim()) return docXml;
  const sectPrIdx = docXml.lastIndexOf('<w:sectPr');
  if (sectPrIdx >= 0) return docXml.slice(0, sectPrIdx) + newXml + docXml.slice(sectPrIdx);
  return docXml.replace(/<\/w:body>/, () => `${newXml}</w:body>`);
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

    let normalizedInner = pInner;
    // Strip proofErr and gramErr tags so spell-check markers between placeholder chars don't split tokens
    normalizedInner = normalizedInner.replace(/<w:(?:proofErr|gramErr)[^>]*>/gi, '');

    // Non-destructively merge adjacent text runs where the first run has an unclosed '{'
    const mergeRunRegex = /(<w:t\b[^>]*>(?:(?!<\/w:t>)[\s\S])*?\{+[^}]*?)\s*<\/w:t>\s*<\/w:r>\s*(?:<w:(?:proofErr|gramErr)[^>]*>\s*)*<w:r\b[^>]*>(?:\s*<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:t\b[^>]*>/gi;

    let prev: string;
    do {
      prev = normalizedInner;
      normalizedInner = normalizedInner.replace(mergeRunRegex, (_m: string, g1: string) => g1);
    } while (normalizedInner !== prev);

    // Normalize any '{ {' into '{{' and '} }' into '}}' and common Word typo splits
    normalizedInner = normalizedInner.replace(/\{\s+\{/g, '{{').replace(/\}\s+\}/g, '}}');
    normalizedInner = normalizedInner.replace(/role\s+_title/gi, 'role_title');
    normalizedInner = normalizedInner.replace(/client\s+_name/gi, 'client_name');

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
    id: 'education',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(education|credentials|academics|qualifications)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['education', 'academic', 'qualification', 'degree', 'credential'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildEducationSectionXml(data, incHeading),
  },
  {
    id: 'skills',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(skills|technical_skills|competencies|core_competencies|skill_matrix|expertise)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['skill', 'competenc', 'expertise', 'technolog'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildSkillsSectionXml(data, incHeading),
  },
  {
    id: 'awards',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(awards|awards_and_achievements|achievements|honors)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['award', 'achievement', 'appreciation', 'recognition', 'honor'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildAwardsSectionXml(data, incHeading),
  },
  {
    id: 'experience',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(experience|professional_experience|work_experience|employment|work_history)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['experience', 'employment', 'history', 'career', 'work'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildExperienceSectionXml(data, incHeading),
  },
  {
    id: 'certifications',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(certifications|certificates|licenses)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['certif', 'licens', 'accredit'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildCertificationsSectionXml(data, incHeading),
  },
  {
    id: 'projects',
    regex: /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(projects|key_projects|project_experience|engagements)\s*\}\}[\s\S]*?<\/w:p>/gi,
    keywords: ['project', 'engagement', 'initiative'],
    buildXml: (data: ResumeData, incHeading: boolean) => buildExperienceSectionXml(data, incHeading),
  },
];

export function processGranularTemplate(
  docXml: string,
  data: ResumeData
): { xml: string; isGranularTemplate: boolean; claimedSections: Set<string> } {
  let xml = docXml;
  const claimedSections = new Set<string>();

  // Check if template contains granular placeholders
  const hasGranularSummary = /\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}/i.test(xml);
  const hasGranularSpecialized = /\{\{\s*(?:specialized_skills_bullets|specialized_skills)\s*\}\}/i.test(xml);
  const hasGranularSkills = /\{\{\s*(?:operating_systems|operating_system|os|sap_software|sap_skills|sap|database|databases|db|other_technical_skills|other_skills|additional_skills)\s*\}\}/i.test(xml);
  const hasGranularExp = /\{\{\s*(?:project_number|client_name|role_title|duration|environment|working_environment_bullets|responsibilities_bullets)\s*\}\}/i.test(xml);
  const hasGranularEdu = /\{\{\s*(?:education_bullets|education_list)\s*\}\}/i.test(xml);
  const hasGranularAwards = /\{\{\s*(?:awards_bullets|awards_list)\s*\}\}/i.test(xml);
  const hasGranularCerts = /\{\{\s*(?:certifications_bullets|certifications_list)\s*\}\}/i.test(xml);

  const isGranularTemplate = Boolean(
    hasGranularSummary ||
    hasGranularSpecialized ||
    hasGranularSkills ||
    hasGranularExp ||
    hasGranularEdu ||
    hasGranularAwards ||
    hasGranularCerts
  );

  if (!isGranularTemplate) {
    return { xml, isGranularTemplate: false, claimedSections };
  }

  // A) Professional Summary
  if (hasGranularSummary) {
    claimedSections.add('summary');
  }
  const summaryItems = Array.isArray(data.professionalSummary)
    ? data.professionalSummary
    : typeof data.professionalSummary === 'string' && data.professionalSummary.trim()
    ? data.professionalSummary.split('\n').filter(Boolean)
    : [];

  let specXml = '';
  if (data.specializedSkills && data.specializedSkills.length > 0) {
    specXml = data.specializedSkills
      .map((g) => {
        const h = buildSubHeadingXml(g.title);
        const b = (g.bullets || []).map((bullet) => buildBulletParagraphXml(bullet)).join('');
        return h + b;
      })
      .join('');
  }

  if (summaryItems.length > 0) {
    const sumMatch = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}[\s\S]*?<\/w:p>/i);
    if (sumMatch) {
      const sumProto = sumMatch[0];
      let bulletsXml = summaryItems.map((b) => {
        return sumProto.replace(/\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}/gi, escapeXml(b));
      }).join('');
      if (!hasGranularSpecialized && specXml) {
        bulletsXml += specXml;
      }
      xml = xml.replace(sumMatch[0], bulletsXml);
    } else {
      let bulletsXml = summaryItems.map((b) => buildBulletParagraphXml(b)).join('');
      if (!hasGranularSpecialized && specXml) {
        bulletsXml += specXml;
      }
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}[\s\S]*?<\/w:p>/gi,
        bulletsXml
      );
    }
  } else {
    xml = xml.replace(
      /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Professional\s+Summary:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}[\s\S]*?<\/w:p>/gi,
      ''
    );
    xml = xml.replace(
      /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Professional\s+Summary:?[\s\S]*?<\/w:p>/gi,
      ''
    );
    xml = xml.replace(
      /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:professional_summary_bullets|summary_bullets)\s*\}\}[\s\S]*?<\/w:p>/gi,
      ''
    );
  }

  // B) Specialized Skills (e.g. Migration Projects, SKILLS@SAP R/3, etc.)
  if (hasGranularSpecialized) {
    if (specXml) {
      // Replace preceding bracketed/unbracketed heading + placeholder if present
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[?Specialized\s+Skill[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:specialized_skills_bullets|specialized_skills)\s*\}\}[\s\S]*?<\/w:p>/gi,
        specXml
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:specialized_skills_bullets|specialized_skills)\s*\}\}[\s\S]*?<\/w:p>/gi,
        specXml
      );
    } else {
      // Completely remove heading and placeholder if candidate has NO specialized skills
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[?Specialized\s+Skill[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:specialized_skills_bullets|specialized_skills)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[?Specialized\s+Skill[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:specialized_skills_bullets|specialized_skills)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }
  } else {
    // If not in template, clean up any bracketed dummy heading if left
    xml = xml.replace(
      /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[?Specialized\s+Skill[\s\S]*?<\/w:p>/gi,
      ''
    );
  }

  // C) Education
  const hasEduInTemplate = /\{\{\s*(?:education_bullets|education_list)\s*\}\}/i.test(xml);
  if (hasEduInTemplate) {
    claimedSections.add('education');
    if (data.education && data.education.length > 0) {
      const eduMatch = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:education_bullets|education_list)\s*\}\}[\s\S]*?<\/w:p>/i);
      if (eduMatch) {
        const eduProto = eduMatch[0];
        const eduXml = data.education
          .map((edu) => {
            const line = [
              edu.degree,
              edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
              edu.institution ? `from ${edu.institution}` : '',
              edu.endDate ? `(${edu.endDate})` : '',
            ]
              .filter(Boolean)
              .join(' ')
              .trim();
            return eduProto.replace(/\{\{\s*(?:education_bullets|education_list)\s*\}\}/gi, escapeXml(line));
          })
          .join('');
        xml = xml.replace(eduMatch[0], eduXml);
      } else {
        const eduXml = data.education
          .map((edu) => {
            const line = [
              edu.degree,
              edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
              edu.institution ? `from ${edu.institution}` : '',
              edu.endDate ? `(${edu.endDate})` : '',
            ]
              .filter(Boolean)
              .join(' ')
              .trim();
            return buildBulletParagraphXml(line);
          })
          .join('');
        xml = xml.replace(
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:education_bullets|education_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
          eduXml
        );
      }
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Education:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:education_bullets|education_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Education:?[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:education_bullets|education_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }
  }

  // D) Technical Skills
  const hasSkillsInTemplate = /\{\{\s*(?:operating_systems|operating_system|os|sap_software|sap_skills|sap|database|databases|db|other_technical_skills|other_skills|additional_skills)\s*\}\}/i.test(xml);
  if (hasSkillsInTemplate) {
    claimedSections.add('skills');
    const skillsList = data.skills || [];
    const usedCategories = new Set<string>();

    function matchCategory(patterns: string[]): string | null {
      for (const cat of skillsList) {
        if (usedCategories.has(cat.category)) continue;
        const low = cat.category.toLowerCase();
        if (patterns.some((p) => low.includes(p))) {
          usedCategories.add(cat.category);
          return cat.skills.join(', ');
        }
      }
      return null;
    }

    const osVal = matchCategory(['operating system', 'os', 'platform']);
    if (osVal) {
      xml = xml.replace(/\{\{\s*(?:operating_systems|operating_system|os)\s*\}\}/gi, () => escapeXml(osVal));
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Operating\s+Systems?:?[\s\S]*?\{\{\s*(?:operating_systems|operating_system|os)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:operating_systems|operating_system|os)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }

    const sapVal = matchCategory(['sap', 'erp', 'netweaver', 'hana']);
    if (sapVal) {
      xml = xml.replace(/\{\{\s*(?:sap_software|sap_skills|sap)\s*\}\}/gi, () => escapeXml(sapVal));
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?SAP\s+Software:?[\s\S]*?\{\{\s*(?:sap_software|sap_skills|sap)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:sap_software|sap_skills|sap)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }

    const dbVal = matchCategory(['database', 'db', 'rdbms', 'sql', 'nosql']);
    if (dbVal) {
      xml = xml.replace(/\{\{\s*(?:database|databases|db)\s*\}\}/gi, () => escapeXml(dbVal));
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Database:?[\s\S]*?\{\{\s*(?:database|databases|db)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:database|databases|db)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }

    const remainingCats = skillsList.filter((c) => !usedCategories.has(c.category));
    if (remainingCats.length > 0) {
      const otherXml = remainingCats
        .map((c) => {
          const tabs = c.category.length <= 12 ? 2 : 1;
          return buildKeyValueTabXml(c.category, c.skills.join(', '), tabs, false);
        })
        .join('');
      const otherSkillRegex = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:other_technical_skills|other_skills|additional_skills)\s*\}\}[\s\S]*?<\/w:p>/gi;
      if (otherSkillRegex.test(xml)) {
        xml = xml.replace(otherSkillRegex, () => otherXml);
      } else {
        // Template did not have {{ other_technical_skills }} tag - insert after the last rendered skill line!
        const lastSkillLineMatch = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?(?:Database|Operating\s+Systems?|SAP\s+Software):?[\s\S]*?<\/w:p>/gi);
        if (lastSkillLineMatch) {
          const lastLine = lastSkillLineMatch[lastSkillLineMatch.length - 1];
          xml = xml.replace(lastLine, () => `${lastLine}${otherXml}`);
        } else {
          // If no specific skill line was matched, look for Technical Skills heading:
          const techHeadingRegex = /(<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Technical\s+Skills:?[\s\S]*?<\/w:p>)/i;
          if (techHeadingRegex.test(xml)) {
            xml = xml.replace(techHeadingRegex, (_m, g1) => `${g1}${otherXml}`);
          } else {
            claimedSections.delete('skills'); // Let the fallback section builder render all skills so nothing is lost!
          }
        }
      }
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:other_technical_skills|other_skills|additional_skills)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }

    if (skillsList.length === 0) {
      xml = xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Technical\s+Skills:?[\s\S]*?<\/w:p>/gi, '');
    }
  }

  // E) Awards and Achievements
  const awardsList = data.awards || [];
  const hasAwardsInTemplate = /\{\{\s*(?:awards_bullets|awards_list)\s*\}\}/i.test(xml);
  if (hasAwardsInTemplate) {
    claimedSections.add('awards');
    if (awardsList.length > 0) {
      const awardsMatch = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:awards_bullets|awards_list)\s*\}\}[\s\S]*?<\/w:p>/i);
      if (awardsMatch) {
        const awardProto = awardsMatch[0];
        const awardsXml = awardsList.map((a) => {
          return awardProto.replace(/\{\{\s*(?:awards_bullets|awards_list)\s*\}\}/gi, () => escapeXml(a));
        }).join('');
        xml = xml.replace(awardsMatch[0], () => awardsXml);
      } else {
        const awardsXml = awardsList.map((a) => buildBulletParagraphXml(a)).join('');
        xml = xml.replace(
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:awards_bullets|awards_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
          () => awardsXml
        );
      }
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Awards\s+and\s+Achievements:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:awards_bullets|awards_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Awards\s+and\s+Achievements:?[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:awards_bullets|awards_list)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }
  }

  // F) Professional Experience Block: DYNAMIC PROTOTYPE CLONING
  const expPattern = /(<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?(?:project_number|\{\{\s*project)[\s\S]*?<\/w:p>[\s\S]*?<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>)/i;
  const expMatch = xml.match(expPattern);
  const fallbackPattern = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Project\s+\{\{\s*project_number\s*\}\}[\s\S]*?<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>/gi;

  if (expMatch || fallbackPattern.test(xml)) {
    claimedSections.add('experience');
    claimedSections.add('projects');

  if (data.experience && data.experience.length > 0) {
    if (expMatch) {
      const prototypeBlockXml = expMatch[1];

      const generatedBlocks = data.experience.map((exp, idx) => {
        let bXml = prototypeBlockXml;

        // 1. Project number
        const projNum = exp.projectNumber || `Project ${idx + 1}`;
        bXml = bXml.replace(/\{\{\s*project_number\s*\}\}/gi, escapeXml(exp.projectName ? `${projNum}: ${exp.projectName}` : projNum));

        // 2. Client / Company name
        bXml = bXml.replace(/\{\{\s*(?:client_name|company|client)\s*\}\}/gi, escapeXml(exp.company || ''));

        // 3. Role / Title
        bXml = bXml.replace(/\{\{\s*(?:role_title|role|job_title)\s*\}\}/gi, escapeXml(exp.role || ''));

        // 4. Duration (replaces all occurrences in the block, e.g. if moved to right side)
        const duration = [exp.startDate, exp.endDate].filter(Boolean).join(' to ');
        bXml = bXml.replace(/\{\{\s*duration\s*\}\}/gi, escapeXml(duration));

        // 5. Environment
        if (exp.environment && exp.environment.trim()) {
          bXml = bXml.replace(/\{\{\s*environment\s*\}\}/gi, escapeXml(exp.environment.trim()));
        } else {
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Environment:?[\s\S]*?\{\{\s*environment\s*\}\}[\s\S]*?<\/w:p>/gi, '');
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*environment\s*\}\}[\s\S]*?<\/w:p>/gi, '');
        }

        // 6. Working Environment (conditional)
        if (exp.workingEnvironment && exp.workingEnvironment.length > 0) {
          const weMatch = bXml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*working_environment_bullets\s*\}\}[\s\S]*?<\/w:p>/i);
          if (weMatch) {
            const bulletProto = weMatch[0];
            const bulletsXml = exp.workingEnvironment.map((item) => {
              return bulletProto.replace(/\{\{\s*working_environment_bullets\s*\}\}/gi, escapeXml(item));
            }).join('');
            bXml = bXml.replace(weMatch[0], bulletsXml);
          }
        } else {
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Working\s+Environment:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*working_environment_bullets\s*\}\}[\s\S]*?<\/w:p>/gi, '');
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Working\s+Environment:?[\s\S]*?<\/w:p>/gi, '');
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*working_environment_bullets\s*\}\}[\s\S]*?<\/w:p>/gi, '');
        }

        // 7. Responsibilities bullets
        const enrichedBullets = [
          ...(exp.description ? [`Project Description: ${exp.description}`] : []),
          ...(exp.domain ? [`Domain: ${exp.domain}`] : []),
          ...(exp.onsiteOffshore ? [`Onsite/Offshore: ${exp.onsiteOffshore}`] : []),
          ...(exp.bullets || []),
        ];
        if (enrichedBullets.length > 0) {
          const respMatch = bXml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>/i);
          if (respMatch) {
            const bulletProto = respMatch[0];
            const bulletsXml = enrichedBullets.map((item) => {
              return bulletProto.replace(/\{\{\s*responsibilities_bullets\s*\}\}/gi, escapeXml(item));
            }).join('');
            bXml = bXml.replace(respMatch[0], bulletsXml);
          }
        } else {
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Responsibilities:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>/gi, '');
          bXml = bXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>/gi, '');
        }

        return bXml;
      }).join('');

      xml = xml.replace(expMatch[0], () => generatedBlocks);
    } else {
      const expXml = buildExperienceSectionXml(data, false);
      const fallbackPattern = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Project\s+\{\{\s*project_number\s*\}\}[\s\S]*?<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*responsibilities_bullets\s*\}\}[\s\S]*?<\/w:p>/gi;
      if (fallbackPattern.test(xml)) {
        xml = xml.replace(fallbackPattern, () => expXml);
      }
    }
  } else {
    // If candidate has NO experience, remove the prototype block and heading
    if (expMatch) {
      xml = xml.replace(expMatch[0], '');
    }
    xml = xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Professional\s+Experience:?[\s\S]*?<\/w:p>/gi, '');
  }
}

  // G) Professional Certifications
  const certsList = data.certifications || [];
  const hasCertsInTemplate = /\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}/i.test(xml);

  if (hasCertsInTemplate) {
    claimedSections.add('certifications');
    if (certsList.length > 0) {
      const certsMatch = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}[\s\S]*?<\/w:p>/i);
      if (certsMatch) {
        const certProto = certsMatch[0];
        const certsXml = certsList.map((c) => {
          return certProto.replace(/\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}/gi, () => escapeXml(c));
        }).join('');
        xml = xml.replace(certsMatch[0], () => certsXml);
      } else {
        const certsXml = certsList.map((c) => buildBulletParagraphXml(c)).join('');
        xml = xml.replace(
          /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}[\s\S]*?<\/w:p>/gi,
          () => certsXml
        );
      }
    } else {
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Professional\s+Certifications:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Certifications:?[\s\S]*?<\/w:p>\s*<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Professional\s+Certifications:?[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?Certifications:?[\s\S]*?<\/w:p>/gi,
        ''
      );
      xml = xml.replace(
        /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(?:certifications_bullets|certifications_list|certifications|certificates)\s*\}\}[\s\S]*?<\/w:p>/gi,
        ''
      );
    }
  }

  // Cleanup: Any bracketed placeholder instruction paragraph like [Specialized Skill Highlights / Migration Projects - if applicable]
  xml = xml.replace(
    /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[[^\]]*(?:if applicable|Other Skills|Highlights)[^\]]*\][\s\S]*?<\/w:p>/gi,
    ''
  );

  return { xml, isGranularTemplate: true, claimedSections };
}

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
      const rawName = data.fullName?.trim() || 'Candidate Profile';
      const candidateName = formatCandidateFirstName(rawName);
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

          // A) Prevent floating logo collision: Replace tight text wrapping on floating drawings with wrapNone
          hXml = hXml.replace(/<wp:wrapTight[^>]*>[\s\S]*?<\/wp:wrapTight>/gi, '<wp:wrapNone/>');

          // B) Remove colliding trailing tab right after candidate name or role
          hXml = hXml.replace(
            /(<w:t[^>]*>\{\{\s*(?:candidate_name|name|full_name|role|target_role)[\s\S]*?<\/w:t><\/w:r>)\s*<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:tab\/>\s*<\/w:r>/gi,
            '$1'
          );

          // C) Candidate Name Heading: Calibri 22pt (sz="44", szCs="44")
          hXml = hXml.replace(/<w:sz w:val="(?:1[0-9]|2[0-9]|3[0-9]|4[0-9])"\/>/gi, '<w:sz w:val="44"/>');
          hXml = hXml.replace(/<w:szCs w:val="(?:1[0-9]|2[0-9]|3[0-9]|4[0-9])"\/>/gi, '<w:szCs w:val="44"/>');
          hXml = hXml.replace(/<w:rFonts\b[^>]*\/>/gi, '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>');

          // Dynamically replace {{ Role }} or {{ candidate_name }} in header
          // Preserves first-letter capitalized name e.g. "Abhiram"
          hXml = hXml.replace(ROLE_REGEX, () => escapeXml(targetRole));
          hXml = hXml.replace(NAME_REGEX, () => escapeXml(candidateName));
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

        // 2b. Process Granular Placeholders (if present in custom template)
        const granularResult = processGranularTemplate(docXml, data);
        docXml = granularResult.xml;
        const isGranularHandled = granularResult.isGranularTemplate;
        const claimedSections = new Set<string>(granularResult.claimedSections);

        // 3. Process individual section placeholders
        for (const spec of SECTION_SPECS) {
          if (claimedSections.has(spec.id)) continue;
          if (spec.regex.test(docXml)) {
            claimedSections.add(spec.id);
            const testXml = spec.buildXml(data, true, options);
            if (!testXml || !testXml.trim()) {
              // Section has NO content: clean up preceding heading + placeholder
              const kwPattern = spec.keywords.join('|');
              const headingAndPlaceholderRegex = new RegExp(
                `<w:p\\b[^>]*>(?:(?!<\\/w:p>)[\\s\\S])*?(?:${kwPattern})[\\s\\S]*?<\\/w:p>\\s*` +
                `<w:p\\b[^>]*>(?:(?!<\\/w:p>)[\\s\\S])*?\\{\\{\\s*${spec.id}\\s*\\}\\}[\\s\\S]*?<\\/w:p>`,
                'gi'
              );
              if (headingAndPlaceholderRegex.test(docXml)) {
                docXml = docXml.replace(headingAndPlaceholderRegex, '');
              } else {
                docXml = docXml.replace(spec.regex, '');
              }
            } else {
              docXml = docXml.replace(spec.regex, (_fullMatch, _p1, matchOffset) => {
                const hasCustomHeading = checkPrecedingHeading(docXml, matchOffset, spec.keywords);
                const sectionXml = spec.buildXml(data, !hasCustomHeading, options);
                return sectionXml || '';
              });
            }
          }
        }

        // 4. Content / Catch-all placeholder (or document append if unplaced)
        const CONTENT_REGEX = /<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\{\{\s*(content|body|resume_content|all_sections)\s*\}\}[\s\S]*?<\/w:p>/gi;
        const hasContentPlaceholder = CONTENT_REGEX.test(docXml);

        // Build remaining sections in the exact classic Tecnics order:
        // 1. Professional Summary (and specialized skill sections)
        // 2. Education
        // 3. Technical Skills
        // 4. Awards and Achievements
        // 5. Professional Experience (Project 1, 2, ...)
        // 6. Professional Certifications
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
                `<w:p><w:pPr><w:spacing w:before="40" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">${escapeXml(data.targetRole)}</w:t></w:r></w:p>`
              );
            }
            if (contactBar.length > 0) {
              remainingParts.push(
                `<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="475569"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">${escapeXml(contactBar.join('   |   '))}</w:t></w:r></w:p>`
              );
            }
          }
        }

        if (!claimedSections.has('summary')) {
          remainingParts.push(buildSummarySectionXml(data, true));
        }
        if (!claimedSections.has('education')) {
          remainingParts.push(buildEducationSectionXml(data, true));
        }
        if (!claimedSections.has('skills')) {
          remainingParts.push(buildSkillsSectionXml(data, true));
        }
        if (!claimedSections.has('awards')) {
          remainingParts.push(buildAwardsSectionXml(data, true));
        }
        if (!claimedSections.has('experience')) {
          remainingParts.push(buildExperienceSectionXml(data, true));
        }
        if (!claimedSections.has('certifications')) {
          remainingParts.push(buildCertificationsSectionXml(data, true));
        }
        remainingParts.push(buildAdditionalSectionsXml(data));

        const remainingXml = remainingParts.join('');

        if (hasContentPlaceholder) {
          docXml = docXml.replace(CONTENT_REGEX, () => remainingXml);
        } else if (claimedSections.size === 0 && !isGranularHandled) {
          // If no placeholders were used at all and NOT a granular template, append to w:body
          docXml = appendBeforeSectPr(docXml, remainingXml);
        } else if (remainingXml.trim()) {
          // ZERO DATA LOSS: Append any sections present in candidate data that lacked a template placeholder
          docXml = appendBeforeSectPr(docXml, remainingXml);
        }

        if (isGranularHandled) {
          // Clean up any lingering optional instructional headings or unused placeholders
          docXml = docXml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?\[[^\]]*(?:if applicable|Other Skills|Highlights)[^\]]*\][\s\S]*?<\/w:p>/gi, '');
          docXml = docXml.replace(/\{\{\s*[\w\s.-]+\s*\}\}/gi, '');
        }

        zip.file('word/document.xml', docXml);
      }

      return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    } catch (err) {
      console.warn('Template XML injection encountered an issue, falling back to clean generator:', err);
    }
  }

  return await generateCleanFallbackDocx(data, options);
}

async function generateCleanFallbackDocx(
  data: ResumeData,
  options?: { removePersonalInfo?: boolean }
): Promise<Buffer> {
  const children: Paragraph[] = [];
  const headingName = formatCandidateFirstName(data.fullName || 'Candidate Profile');

  const createSectionHeading = (title: string) =>
    new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [
        new TextRun({
          text: title.endsWith(':') ? title : `${title}:`,
          bold: true,
          size: 26, // 13pt
          color: '000000',
          font: 'Times New Roman',
        }),
      ],
    });

  const createSubHeading = (title: string) =>
    new Paragraph({
      spacing: { before: 140, after: 40 },
      children: [
        new TextRun({
          text: title,
          bold: true,
          size: 22, // 11pt
          color: '000000',
          font: 'Times New Roman',
        }),
      ],
    });

  const createBullet = (text: string) =>
    new Paragraph({
      spacing: { before: 20, after: 30 },
      indent: { left: 360, hanging: 240 },
      children: [
        new TextRun({
          text: `•\t${text}`,
          size: 22, // 11pt
          color: '000000',
          font: 'Times New Roman',
        }),
      ],
    });

  const createKeyValue = (label: string, value: string, boldValue = false) =>
    new Paragraph({
      spacing: { before: 20, after: 30 },
      children: [
        new TextRun({
          text: label.endsWith(':') ? `${label}\t` : `${label}:\t`,
          bold: true,
          size: 22,
          color: '000000',
          font: 'Times New Roman',
        }),
        new TextRun({
          text: value,
          bold: boldValue,
          size: 22,
          color: '000000',
          font: 'Times New Roman',
        }),
      ],
    });

  // 1. Candidate Name (Calibri 22pt bold)
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({
          text: headingName,
          bold: true,
          size: 44, // 22pt
          color: '000000',
          font: 'Calibri',
        }),
      ],
    })
  );

  // 2. Target Role (Times New Roman 11pt bold)
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
            color: '000000',
            font: 'Times New Roman',
          }),
        ],
      })
    );
  }

  // 3. Contact Info (if not removed)
  if (!options?.removePersonalInfo && data.contactInfo) {
    const contactParts = [
      data.contactInfo.email,
      data.contactInfo.phone,
      data.contactInfo.location,
      data.contactInfo.linkedin,
    ].filter(Boolean);
    if (contactParts.length > 0) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 120 },
          children: [
            new TextRun({
              text: contactParts.join('   |   '),
              size: 20,
              color: '475569',
              font: 'Times New Roman',
            }),
          ],
        })
      );
    }
  }

  // 4. Professional Summary
  const summaryItems = Array.isArray(data.professionalSummary)
    ? data.professionalSummary
    : typeof data.professionalSummary === 'string' && data.professionalSummary.trim()
    ? data.professionalSummary.split('\n').filter(Boolean)
    : [];

  if (summaryItems.length > 0) {
    children.push(createSectionHeading('Professional Summary:'));
    for (const item of summaryItems) {
      children.push(createBullet(item));
    }
  }

  // 5. Specialized Skills
  if (data.specializedSkills && data.specializedSkills.length > 0) {
    for (const group of data.specializedSkills) {
      if (group.title && group.bullets?.length > 0) {
        children.push(createSubHeading(group.title));
        for (const b of group.bullets) {
          children.push(createBullet(b));
        }
      }
    }
  }

  // 6. Education
  if (data.education && data.education.length > 0) {
    children.push(createSectionHeading('Education:'));
    for (const edu of data.education) {
      const line = [
        edu.degree,
        edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
        edu.institution ? `from ${edu.institution}` : '',
        edu.endDate ? `(${edu.endDate})` : '',
      ]
        .filter(Boolean)
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (line) children.push(createBullet(line));
    }
  }

  // 7. Technical Skills
  if (data.skills && data.skills.length > 0) {
    children.push(createSectionHeading('Technical Skills:'));
    for (const cat of data.skills) {
      if (cat && Array.isArray(cat.skills) && cat.skills.length > 0) {
        children.push(createKeyValue(cat.category || 'Skills', cat.skills.join(', ')));
      }
    }
  }

  // 8. Awards and Achievements
  if (data.awards && data.awards.length > 0) {
    children.push(createSectionHeading('Awards and Achievements:'));
    for (const award of data.awards) {
      children.push(createBullet(award));
    }
  }

  // 9. Professional Experience
  if (data.experience && data.experience.length > 0) {
    children.push(createSectionHeading('Professional Experience:'));
    data.experience.forEach((exp, idx) => {
      const projTitle = exp.projectNumber || `Project ${idx + 1}`;
      children.push(createSubHeading(exp.projectName ? `${projTitle}: ${exp.projectName}` : projTitle));

      const duration = [exp.startDate, exp.endDate].filter(Boolean).join(' to ');
      if (exp.company) children.push(createKeyValue('Client:', exp.company, true));
      if (exp.role) children.push(createKeyValue('Role:', exp.role, true));
      if (duration) children.push(createKeyValue('Duration:', duration, true));
      if (exp.environment) children.push(createKeyValue('Environment:', exp.environment, true));
      if (exp.domain) children.push(createKeyValue('Domain:', exp.domain, true));
      if (exp.onsiteOffshore) children.push(createKeyValue('Location:', exp.onsiteOffshore, true));
      if (exp.description) children.push(createBullet(`Project Description: ${exp.description}`));

      if (exp.workingEnvironment && exp.workingEnvironment.length > 0) {
        children.push(createSubHeading('Working Environment:'));
        for (const we of exp.workingEnvironment) {
          children.push(createBullet(we));
        }
      }

      if (exp.bullets && exp.bullets.length > 0) {
        children.push(createSubHeading('Responsibilities:'));
        for (const b of exp.bullets) {
          children.push(createBullet(b));
        }
      }
    });
  }

  // 10. Professional Certifications
  if (data.certifications && data.certifications.length > 0) {
    children.push(createSectionHeading('Professional Certifications:'));
    for (const cert of data.certifications) {
      children.push(createBullet(cert));
    }
  }

  // 11. Additional Sections
  if (data.additionalSections && data.additionalSections.length > 0) {
    for (const sec of data.additionalSections) {
      if (sec.title && sec.items?.length > 0) {
        children.push(createSectionHeading(sec.title));
        for (const item of sec.items) {
          children.push(createBullet(item));
        }
      }
    }
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
                    size: 18,
                    color: '64748B',
                    font: 'Times New Roman',
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
                    size: 16,
                    color: '94A3B8',
                    font: 'Times New Roman',
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
